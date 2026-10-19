"""Render real Qt screens with synthetic documents and verify the safe XLSX pipeline.

Run with --scale 1, 1.25, 1.5 or 2 to exercise Qt's device scale factor.
The report distinguishes logical window sizes from rendered pixel sizes.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import time
from pathlib import Path
from uuid import uuid4


def synthetic_photos(directory: Path) -> tuple[Path, Path]:
    from PIL import Image, ImageDraw, ImageFont

    directory.mkdir(parents=True, exist_ok=True)
    font_path = next(
        (
            p
            for p in [
                Path("C:/Windows/Fonts/arial.ttf"),
                Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
            ]
            if p.exists()
        ),
        None,
    )
    font = ImageFont.truetype(str(font_path), 36) if font_path else ImageFont.load_default(size=36)
    small = ImageFont.truetype(str(font_path), 25) if font_path else font
    lines = [
        [
            "KYRGYZ REPUBLIC / SYNTHETIC SAMPLE",
            "ФАМИЛИЯ / SURNAME: ТЕСТОВ / TESTOV",
            "ИМЯ / GIVEN NAME: ПРИМЕР / PRIMER",
            "ОТЧЕСТВО: СИНТЕТИЧЕСКИЙ",
            "ДАТА РОЖДЕНИЯ: 01.02.1990",
            "ГРАЖДАНСТВО: КЫРГЫЗСТАН",
            "ПОЛ: МУЖСКОЙ",
        ],
        [
            "KYRGYZ REPUBLIC / SYNTHETIC SAMPLE",
            "ОРГАН ВЫДАЧИ: ТЕСТОВЫЙ ОРГАН",
            "ДАТА ВЫДАЧИ: 03.04.2020",
            "ПЕРСОНАЛЬНЫЙ НОМЕР: 00123456789012",
            "НОМЕР ДОКУМЕНТА: AA0000001",
            "СРОК ДЕЙСТВИЯ: 03.04.2030",
            "АДРЕС РЕГИСТРАЦИИ: ТЕСТОВАЯ ОБЛАСТЬ",
            "ТЕСТОВЫЙ ГОРОД, УЛ. ПРИМЕР, Д. 1",
        ],
    ]
    result = []
    for index, content in enumerate(lines):
        image = Image.new("RGB", (1600, 2300), "#e8eef4")
        draw = ImageDraw.Draw(image)
        draw.text((90, 60), "09:41       SYNTHETIC PHONE SCREEN", fill="#465568", font=font)
        draw.text((120, 200), "Цифровой документ · СИНТЕТИЧЕСКИЙ ПРИМЕР", fill="#24374d", font=font)
        draw.rounded_rectangle((90, 390, 1510, 1570), radius=24, fill="white", outline="#a6b5c5", width=4)
        draw.text((145, 435), "НЕ ДОКУМЕНТ · ONLY FOR LOCAL TESTING", fill="#b45309", font=small)
        for row, text in enumerate(content):
            draw.text((145, 540 + row * 111), text, fill="#0f172a", font=font)
        for row in range(17):
            for col in range(17):
                if (row * 5 + col * 3) % 7 < 3:
                    x, y = 620 + col * 20, 1780 + row * 20
                    draw.rectangle((x, y, x + 18, y + 18), fill="#24374d")
        draw.text((560, 2160), "ТЕСТОВЫЙ QR / NOT VALID", fill="#465568", font=small)
        path = directory / ("synthetic-front.png" if index == 0 else "synthetic-back.png")
        image.save(path)
        result.append(path)
    return result[0], result[1]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--scale", type=float, default=1.0, choices=[1.0, 1.25, 1.5, 2.0])
    parser.add_argument("--skip-ocr", action="store_true", help="Use only for repeated layout/DPI checks")
    parser.add_argument(
        "--output", type=Path, default=Path(__file__).resolve().parents[1] / "test-artifacts" / "redesign"
    )
    args = parser.parse_args()
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    os.environ["QT_SCALE_FACTOR"] = str(args.scale)

    from openpyxl import load_workbook
    from PySide6.QtCore import QPoint
    from PySide6.QtTest import QTest
    from PySide6.QtWidgets import QAbstractScrollArea, QApplication, QLabel, QMessageBox, QPushButton

    from document_excel.domain.models import (
        Address,
        DocumentSide,
        FieldEvidence,
        FieldStatus,
        RegistrationStatus,
        data_value,
    )
    from document_excel.infrastructure.excel import create_demo_template
    from document_excel.infrastructure.ocr import TesseractEngine
    from document_excel.infrastructure.storage import SettingsRepository
    from document_excel.presentation.dialogs import AddressDialog, ExportDialog
    from document_excel.presentation.main_window import MainWindow
    from document_excel.presentation.theme import apply_theme

    output = args.output.resolve() / f"scale-{round(args.scale * 100)}"
    output.mkdir(parents=True, exist_ok=True)
    front, back = synthetic_photos(output / "synthetic-assets")
    app = QApplication([])
    repository = SettingsRepository(output / "app-data")
    settings = repository.load()
    settings.auto_ocr = False
    settings.default_template = ""
    settings.default_profile = ""
    settings.theme = "light"
    settings.history_enabled = False
    settings.languages = ["rus", "eng"]
    settings.preprocessing = "auto"
    repository.save(settings)
    apply_theme(app, settings)
    window = MainWindow(repository)
    errors = []

    def error(text):
        errors.append(str(text))
        window.notice.setText(str(text))

    window.error = error
    # The registration editor is captured separately below, avoiding an unattended modal after OCR.
    window.case_actions.address_dialog = lambda *a, **kw: None
    window.resize(1366, 768)
    window.show()
    app.processEvents()

    report = {
        "synthetic_only": True,
        "scale_factor": args.scale,
        "screenshots": [],
        "layouts": [],
        "errors": errors,
        "ocr": {"ran": False},
    }

    def settle(milliseconds=50):
        app.processEvents()
        QTest.qWait(milliseconds)
        app.processEvents()

    def wait_for(predicate, timeout=60000):
        deadline = time.monotonic() + timeout / 1000
        while not predicate():
            if time.monotonic() > deadline:
                raise TimeoutError("Synthetic visual QA operation timed out")
            settle(20)

    def capture(name, widget=None):
        target = widget or window
        settle()
        path = output / f"{name}.png"
        rendered = target.grab()
        if not rendered.save(str(path)):
            raise RuntimeError("Could not save synthetic screenshot")
        report["screenshots"].append(
            {
                "name": name,
                "path": str(path),
                "logical_size": [target.width(), target.height()],
                "rendered_pixels": [rendered.width(), rendered.height()],
                "device_pixel_ratio": rendered.devicePixelRatio(),
            }
        )

    def layout_record(name, requested):
        buttons = []
        for button in window.findChildren(QPushButton):
            if not button.isVisible():
                continue
            top_left = button.mapTo(window, QPoint(0, 0))
            bottom_right = button.mapTo(window, button.rect().bottomRight())
            ancestor = button.parentWidget()
            in_scroll_panel = False
            while ancestor and ancestor is not window:
                in_scroll_panel |= isinstance(ancestor, QAbstractScrollArea)
                ancestor = ancestor.parentWidget()
            buttons.append(
                {
                    "text": button.text() or button.toolTip(),
                    "object_name": button.objectName(),
                    "rect": [top_left.x(), top_left.y(), button.width(), button.height()],
                    "inside_window": window.rect().contains(top_left)
                    and window.rect().contains(bottom_right),
                    "in_scroll_panel": in_scroll_panel,
                    "currently_painted": not button.visibleRegion().isEmpty(),
                    "primary": button is window.primary_button,
                }
            )
        report["layouts"].append(
            {
                "state": name,
                "requested_logical_size": list(requested),
                "actual_logical_size": [window.width(), window.height()],
                "minimum_size_hint": [window.minimumSizeHint().width(), window.minimumSizeHint().height()],
                "buttons": buttons,
            }
        )

    capture("01-empty-upload")
    window.case_actions.add_images([front])
    wait_for(lambda: not window.background_workers)
    window.current_case.images[0].role = DocumentSide.FRONT
    window.refresh_photos()
    capture("02-one-photo")
    window.case_actions.add_images([back])
    wait_for(lambda: not window.background_workers)
    window.current_case.images[1].role = DocumentSide.BACK
    window.refresh_photos()
    capture("03-two-photos")
    layout_record("two-photos", (1366, 768))

    if not args.skip_ocr:
        diagnostic = TesseractEngine(window.settings).diagnose()
        if diagnostic.error:
            raise RuntimeError(diagnostic.error)
        window.case_actions.start_ocr()
        assert window.active_worker is not None
        capture("04-real-ocr-running")
        report["ocr"]["ran"] = True
        report["ocr"]["progress_range_during_capture"] = [
            window.progress.minimum(),
            window.progress.maximum(),
        ]
        wait_for(lambda: window.active_worker is None, timeout=120000)
        report["ocr"]["evidence_fields"] = sorted(window.current_case.evidence)
        report["ocr"]["source_texts"] = len(window.case_ocr_texts.get(window.current_case.id, {}))
        report["ocr"]["warnings_count"] = len(window.current_case.errors)
        report["ocr"]["synthetic_field_checks_before_manual_review"] = {
            name: data_value(window.current_case.data, name) == value
            for name, value in {
                "surname": "ТЕСТОВ",
                "first_name": "ПРИМЕР",
                "patronymic": "СИНТЕТИЧЕСКИЙ",
                "birth_date": "01.02.1990",
                "personal_identifier": "00123456789012",
                "passport_type": "AA",
                "passport_number": "0000001",
                "passport_issued_date": "03.04.2020",
                "expiration_date": "03.04.2030",
            }.items()
        }
    else:
        report["ocr"]["reason"] = "Explicit --skip-ocr for repeat layout/DPI validation"

    case = window.current_case
    window.go_step(2)
    for name, value in {
        "full_name": "ТЕСТОВ ПРИМЕР СИНТЕТИЧЕСКИЙ",
        "surname": "ТЕСТОВ",
        "first_name": "ПРИМЕР",
        "patronymic": "СИНТЕТИЧЕСКИЙ",
        "birth_date": "01.02.1990",
        "personal_identifier": "00123456789012",
        "passport_type": "AA",
        "passport_number": "0000001",
        "passport_issued_by": "ТЕСТОВЫЙ ОРГАН",
        "passport_issued_date": "03.04.2020",
        "expiration_date": "03.04.2030",
    }.items():
        window.apply_field_edit(name, value)
    case.data.document_type = "kg_id"
    window.refresh_form()
    window.notice.setText("Данные извлечены. Проверьте выделенные поля.")
    capture("05-review-data")
    window.apply_field_edit("birth_date", "31.02.2000")
    window.apply_field_edit("personal_identifier", "123")
    window.review_form.reveal("birth_date")
    capture("06-review-errors")
    window.apply_field_edit("birth_date", "01.02.1990")
    window.apply_field_edit("personal_identifier", "00123456789012")

    address_dialog = AddressDialog(Address(), True, window)
    address_dialog.entries["country"].setText("Кыргызстан")
    address_dialog.entries["region"].setText("Тестовая область")
    address_dialog.entries["city"].setText("Тестовый город")
    address_dialog.entries["street"].setText("ул. Пример")
    address_dialog.entries["house"].setText("1")
    address_dialog.show()
    capture("07-manual-registration", address_dialog)
    address_dialog.accept()
    assert address_dialog.result() == address_dialog.DialogCode.Accepted
    case.data.registration_address = address_dialog.address()
    case.registration_status = RegistrationStatus.MANUALLY_ENTERED
    case.evidence["registration_address"] = FieldEvidence(
        source="ручной ввод", status=FieldStatus.MANUALLY_CORRECTED
    )
    address_dialog.close()
    case.invalidate()
    window.refresh_form()

    window.go_step(3)
    capture("08-excel-selection")
    template = output / "synthetic-assets" / f"demo-template-{uuid4().hex[:8]}.xlsx"
    create_demo_template(template)
    original_hash = hashlib.sha256(template.read_bytes()).hexdigest()
    case.template_path = template
    window.template_sheets[template] = window.workbooks.sheets(template)
    window.refresh_form()
    original_question = QMessageBox.question
    try:
        QMessageBox.question = lambda *a, **kw: QMessageBox.StandardButton.Yes
        window.case_actions.confirm_data()
    finally:
        QMessageBox.question = original_question
    assert case.confirmed, "Synthetic case could not be confirmed"
    capture("09-excel-mapping")
    plan = window.workbooks.plan(template, case, window.profile)
    preview = ExportDialog(plan, window)
    preview.confirm.setChecked(True)
    preview.show()
    capture("10-export-preview", preview)
    preview.accept()
    assert preview.result() == preview.DialogCode.Accepted
    preview.close()

    result = output / f"synthetic-contract-{uuid4().hex[:8]}.xlsx"
    window.excel_actions.save_plan(plan, result, case, window.profile.model_copy(deep=True))
    wait_for(lambda: window.active_worker is None)
    settle()
    assert case.output_path == result and result.exists(), "Actual XLSX write did not complete"
    window.go_step(4)
    capture("11-success")
    workbook = load_workbook(result, data_only=False)
    try:
        report["xlsx"] = {
            "source_unchanged": hashlib.sha256(template.read_bytes()).hexdigest() == original_hash,
            "output_exists": result.exists(),
            "output_path": str(result),
            "changed_cells": len(plan.changes),
            "formula_preserved": workbook["Акт расч"]["L21"].value == "=G21-F21",
            "pin_preserved_as_text": workbook["Дог."]["C22"].value == "00123456789012",
            "full_name_written": workbook["Дог."]["A18"].value == case.data.full_name,
        }
    finally:
        workbook.close()

    for size in [(1366, 768), (1600, 900), (1920, 1080), (1000, 700)]:
        window.resize(*size)
        for step, name in [(0, "upload"), (2, "review"), (3, "excel"), (4, "success")]:
            window.go_step(step)
            settle()
            layout_record(name, size)
            capture(f"layout-{size[0]}x{size[1]}-{name}")
    window.navigate(3)
    window.resize(1366, 768)
    capture("12-queue")
    report["visible_technical_text_at_finish"] = [
        label.text()
        for label in window.findChildren(QLabel)
        if label.isVisible()
        and any(t in label.text().lower() for t in ["tesseract", "timeout", "кандидатов:"])
    ]
    report["all_main_actions_inside_window"] = all(
        button["inside_window"]
        for layout in report["layouts"]
        for button in layout["buttons"]
        if not button["in_scroll_panel"]
    )
    report_path = output / "visual-report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    wait_for(lambda: not window.background_workers)
    window.close()
    settle()
    print(
        json.dumps(
            {
                "report": str(report_path),
                "screenshots": len(report["screenshots"]),
                "ocr": report["ocr"],
                "xlsx": report["xlsx"],
                "all_main_actions_inside": report["all_main_actions_inside_window"],
            },
            ensure_ascii=False,
        )
    )
    return (
        0
        if all(
            report["xlsx"][k]
            for k in ["source_unchanged", "formula_preserved", "pin_preserved_as_text", "full_name_written"]
        )
        and not errors
        else 1
    )


if __name__ == "__main__":
    raise SystemExit(main())
