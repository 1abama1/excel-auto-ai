"""Real private-photo Qt run. No screenshots, copied photos or personal output."""

import argparse
import json
import re
import time
from pathlib import Path

from PySide6.QtCore import QEventLoop, QTimer
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QComboBox, QPushButton

from document_excel.domain.errors import AppError
from document_excel.domain.models import RegistrationStatus, parse_date
from document_excel.infrastructure.excel import readiness
from document_excel.infrastructure.storage import SettingsRepository
from document_excel.presentation.main_window import MainWindow


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("directory", type=Path)
    parser.add_argument("specification", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    spec = args.specification.read_text(encoding="utf8")
    expected = {}
    for field, label in [
        ("birth_date", "Дата рождения"),
        ("passport_issued_date", "Дата выдачи"),
        ("expiration_date", "Срок действия"),
        ("passport_number", "Номер документа"),
        ("passport_issued_by", "Орган выдачи"),
    ]:
        value = re.search(r"\|\s*" + label + r"\s*\|\s*([^|]+)", spec).group(1).strip().split(" —")[0]
        expected[field] = (
            str(parse_date(value))
            if field.endswith("date")
            else value[2:]
            if field == "passport_number"
            else value
        )
    args.output.parent.mkdir(parents=True, exist_ok=True)
    app = QApplication([])
    repo = SettingsRepository(args.output.parent / "private-gui-app-data")
    settings = repo.load()
    settings.default_template = str(root / "assets/lermontov.xlsx")
    repo.save(settings)
    window = MainWindow(repo)
    window.show()
    errors = []
    window.error = lambda error: errors.append(type(error).__name__)

    def wait_workers():
        deadline = time.monotonic() + 180
        loop = QEventLoop()
        timer = QTimer()
        timer.timeout.connect(
            lambda: (
                loop.quit()
                if not window.active_worker and not window.background_workers or time.monotonic() > deadline
                else None
            )
        )
        timer.start(50)
        loop.exec()
        timer.stop()
        assert time.monotonic() < deadline and not errors, errors

    window.case_actions.add_images([args.directory / "1.png", args.directory / "2.png"])
    window.case_actions.start_ocr()
    wait_workers()
    case = window.current_case
    initial_status = {name: e.status.value for name, e in case.evidence.items()}
    manual_selections = []
    if case.evidence.get("passport_number") and case.evidence["passport_number"].status.value == "CONFLICT":

        def select_controlled_candidate():
            dialog = QApplication.activeModalWidget()
            choices = dialog.findChild(QComboBox)
            for i in range(choices.count()):
                c = choices.itemData(i)
                if c.normalized_value == expected["passport_number"] and c.source_method == "MRZ":
                    choices.setCurrentIndex(i)
                    next(
                        b
                        for b in dialog.findChildren(QPushButton)
                        if b.text().startswith("Использовать выбранное")
                    ).click()
                    manual_selections.append("passport_number")
                    return
            dialog.reject()

        QTimer.singleShot(100, select_controlled_candidate)
        window.show_field_details("passport_number")
    before = case.digest()
    for i in (1, 0, 1, 0):
        window.image_list.setCurrentIndex(i)
        wait_workers()
    unchanged = case.digest() == before
    assert unchanged
    rows = {}
    for name in (
        "surname",
        "first_name",
        "patronymic",
        "full_name",
        "birth_date",
        "passport_number",
        "passport_issued_by",
        "passport_issued_date",
        "expiration_date",
        "personal_identifier",
        "gender",
        "citizenship",
    ):
        e = case.evidence.get(name)
        selected = getattr(case.data, name)
        rows[name] = {
            "found": bool(selected),
            "status": e.status.value if e else "NOT_RECOGNIZED",
            "methods": sorted({c.source_method for c in e.candidates}) if e else [],
            "sides": sorted({c.source_side.value for c in e.candidates}) if e else [],
            "needs_review": True,
        }
        if name in expected:
            rows[name]["selected_matches_control"] = str(selected) == expected[name]
            rows[name]["matching_candidate_found"] = (
                any(c.normalized_value == expected[name] for c in e.candidates) if e else False
            )
    assert (
        case.registration_status == RegistrationStatus.NOT_PRESENT
        and not case.data.registration_address.display()
    )
    address_dialog = []

    def inspect_address():
        dialog = QApplication.activeModalWidget()
        address_dialog.append(dialog is not None)
        dialog.reject()

    QTimer.singleShot(100, inspect_address)
    window.case_actions.address_dialog(True)
    blocked = False
    try:
        window.workbooks.plan(root / "assets/lermontov.xlsx", case, window.profile)
    except AppError:
        blocked = True
    result = {
        "fields": rows,
        "initial_status": initial_status,
        "manual_selection_using_user_control": manual_selections,
        "photo_switch_preserved_data": unchanged,
        "identity_status": case.identity_status,
        "name_parts_distinct": len(
            {p for p in (case.data.surname, case.data.first_name, case.data.patronymic) if p}
        )
        == 3,
        "name_order_correct": case.data.full_name
        == " ".join(p for p in (case.data.surname, case.data.first_name, case.data.patronymic) if p),
        "gender_semantic": case.data.gender in {"MALE", "FEMALE", "UNKNOWN"},
        "citizenship_semantic": case.data.citizenship == "KGZ",
        "mrz_checks": [r.checks for r in case.mrz_reports.values()],
        "address_requested": bool(address_dialog),
        "export_blocked_without_address": blocked,
        "readiness_issues": len(readiness(case, window.profile)),
        "private_values_saved": False,
        "actual_contract_exported": False,
    }
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf8")
    print(json.dumps(result, ensure_ascii=True, indent=2))
    window.close()
    QTest.qWait(100)
    window.deleteLater()
    QTest.qWait(100)
    app.processEvents()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
