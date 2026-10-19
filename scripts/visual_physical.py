"""Synthetic real-Qt OCR flow and supplied-template export. No customer photos."""

import hashlib
import json
import sys
import time
from pathlib import Path
from uuid import uuid4

from PIL import Image, ImageDraw, ImageFont
from PySide6.QtCore import QEventLoop, QTimer
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication

from document_excel.domain.models import Address, FieldEvidence, FieldStatus, RegistrationStatus
from document_excel.infrastructure.physical_verification import synthetic_plastic_images
from document_excel.infrastructure.storage import SettingsRepository
from document_excel.presentation.main_window import MainWindow
from document_excel.presentation.theme import apply_theme


def main():
    root = Path(__file__).resolve().parents[1]
    output = root / "test-artifacts" / "physical-visual" / uuid4().hex[:8]
    output.mkdir(parents=True)
    source = root / "assets" / "lermontov.xlsx"
    source_digest = hashlib.sha256(source.read_bytes()).hexdigest()
    repository = SettingsRepository(output / "app-data")
    settings = repository.load()
    settings.default_template = str(source)
    settings.theme = "dark"
    repository.save(settings)
    app = QApplication([])
    apply_theme(app, settings)
    window = MainWindow(repository)
    window.resize(1366, 768)
    window.show()
    errors = []
    window.error = lambda error: errors.append(str(error))
    paths = synthetic_plastic_images(output)
    official_fields = "--official-fields" in sys.argv
    if official_fields:
        with Image.open(paths[1]) as original:
            image = original.copy()
        w, h = image.size
        draw = ImageDraw.Draw(image)
        draw.rectangle(
            (round(0.265 * w), round(0.313 * h), round(0.63 * w), round(0.35 * h)), fill=(228, 240, 245)
        )
        draw.text(
            (round(0.27 * w), round(0.32 * h)),
            "СЕМ 123456",
            font=ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 36),
            fill=(12, 25, 35),
        )
        image.save(paths[1])
    window.case_actions.add_images(paths)
    QTest.qWait(500)
    window.case_actions.start_ocr()
    deadline = time.monotonic() + 180
    loop = QEventLoop()
    poll = QTimer()

    def check_finished():
        if not window.active_worker and not window.background_workers or time.monotonic() >= deadline:
            loop.quit()

    poll.timeout.connect(check_finished)
    poll.start(100)
    loop.exec()
    poll.stop()
    if time.monotonic() >= deadline:
        print(
            {
                "step": window.current_step,
                "active": bool(window.active_worker),
                "background": len(window.background_workers),
                "stage": window.substage_label.text(),
            },
            flush=True,
        )
        window.cancel_work()
    assert not errors, errors
    assert time.monotonic() < deadline
    case = window.current_case
    assert case.data.document_type == "kg_physical_id"
    assert any(report.valid for report in case.mrz_reports.values())
    assert case.registration_status == RegistrationStatus.NOT_PRESENT
    assert window.profile.id == "lermontov-combined"
    if official_fields:
        assert any(c.normalized_value == "СЕМ 123456" for c in case.evidence["passport_issued_by"].candidates)

        def capture_field():
            dialog = QApplication.activeModalWidget()
            dialog.grab().save(str(output / "authority-details.png"))
            dialog.accept()

        QTimer.singleShot(150, capture_field)
        window.show_field_details("passport_issued_by")
    for width, height in ((1366, 768), (1000, 700)):
        window.resize(width, height)
        QTest.qWait(200)
        window.go_step(0)
        window.grab().save(str(output / f"upload-{width}.png"))
        window.go_step(2)
        window.grab().save(str(output / f"review-{width}.png"))

    def details():
        dialog = QApplication.activeModalWidget()
        dialog.grab().save(str(output / "mrz-details.png"))
        dialog.accept()

    QTimer.singleShot(150, details)
    window.case_actions.show_mrz()
    case.data.registration_address = Address(region="Тестовая область", street="ул. Пример, 1")
    case.registration_status = RegistrationStatus.MANUALLY_ENTERED
    case.evidence["registration_address"] = FieldEvidence(source="ручной ввод", status=FieldStatus.CONFIRMED)
    for evidence in case.evidence.values():
        assert evidence.status not in {FieldStatus.CONFLICT, FieldStatus.INVALID}
        evidence.status = FieldStatus.CONFIRMED
    case.confirmed_digest = case.digest()
    window.go_step(3)
    window.refresh_form()
    window.grab().save(str(output / "excel-photos.png"))
    plan = window.workbooks.plan(source, case, window.profile)
    result = window.workbooks.write(plan, output / "synthetic-physical-contract.xlsx", case, window.profile)
    case.output_path = result
    window.last_result = result
    window.go_step(4)
    window.grab().save(str(output / "done.png"))
    report = {
        "synthetic_only": True,
        "profile": window.profile.id,
        "mrz_checks": [r.checks for r in case.mrz_reports.values()],
        "source_unchanged": hashlib.sha256(source.read_bytes()).hexdigest() == source_digest,
        "embedded_document_photos": len(plan.photos.insertions),
        "result": str(result),
        "screenshots": 8 if official_fields else 7,
        "official_authority_test": official_fields,
        "authority_matches": case.data.passport_issued_by == "СЕМ 123456" if official_fields else None,
    }
    (output / "report.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    window.close()
    QTest.qWait(300)
    window.deleteLater()
    QTest.qWait(100)
    QTimer.singleShot(100, app.quit)
    app.exec()
    print(output)


if __name__ == "__main__":
    main()
