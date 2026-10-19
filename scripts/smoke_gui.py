"""Render the actual Qt widgets with synthetic data; no desktop screen capture."""

from __future__ import annotations

import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
ROOT = Path(__file__).resolve().parents[1]
os.environ["DOCUMENT_EXCEL_DATA_DIR"] = str(ROOT / "test-artifacts" / "app-data")

from PIL import Image, ImageDraw, ImageFont
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication

from document_excel.domain.models import Address, ClientData, FieldEvidence, FieldStatus, RegistrationStatus
from document_excel.infrastructure.storage import SettingsRepository
from document_excel.presentation.main_window import MainWindow
from document_excel.presentation.theme import apply_theme


def main():
    output = ROOT / "test-artifacts"
    output.mkdir(exist_ok=True)
    image = Image.new("RGB", (1400, 900), "#faf8f0")
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 34)
    for i, line in enumerate(
        [
            "СИНТЕТИЧЕСКИЙ ПРИМЕР / НЕ ДОКУМЕНТ",
            "KYRGYZ REPUBLIC · PASSPORT",
            "ФАМИЛИЯ: ТЕСТОВ",
            "ИМЯ: ПРИМЕР",
            "ДАТА РОЖДЕНИЯ: 01.02.1990",
            "ПИН: 00123456789012",
            "НОМЕР ПАСПОРТА: AA0000001",
            "ОРГАН ВЫДАЧИ: МКК 000123",
            "АДРЕС РЕГИСТРАЦИИ: ТЕСТОВЫЙ ГОРОД",
        ]
    ):
        draw.text((60, 65 + i * 85), line, fill="#24374d", font=font)
    photo = output / "synthetic-document.png"
    image.save(photo)
    app = QApplication(sys.argv)
    repo = SettingsRepository(output / "app-data")
    settings = repo.load()
    settings.theme = "dark" if "--dark" in sys.argv else "light"
    repo.save(settings)
    apply_theme(app, settings)
    window = MainWindow(repo)
    window.settings.auto_ocr = False
    window.resize(1366, 768)
    window.case_actions.add_images([photo])
    case = window.current_case
    case.data = ClientData(
        full_name="ТЕСТОВ ПРИМЕР",
        surname="ТЕСТОВ",
        first_name="ПРИМЕР",
        birth_date="1990-02-01",
        personal_identifier="00123456789012",
        passport_type="AA",
        passport_number="0000001",
        passport_issued_by="МКК",
        passport_department_code="000123",
        document_type="kg_id",
        registration_address=Address(raw="Тестовый город, ул. Пример, д. 1"),
    )
    case.registration_status = RegistrationStatus.PARTIALLY_FOUND
    for name in [
        "full_name",
        "surname",
        "first_name",
        "birth_date",
        "personal_identifier",
        "passport_number",
    ]:
        case.evidence[name] = FieldEvidence(source="OCR", status=FieldStatus.VALID)
    case.evidence["registration_address"] = FieldEvidence(source="OCR", status=FieldStatus.REVIEW)
    window.refresh_form()
    window.go_step(2)
    window.show()
    if "--department" in sys.argv:
        window.review_form.reveal("passport_department_code")

    def capture():
        name = "gui-dark.png" if settings.theme == "dark" else "gui-light.png"
        if "--department" in sys.argv:
            name = "gui-department.png"
        window.grab().save(str(output / name))
        window.close()
        QTimer.singleShot(400, app.quit)

    QTimer.singleShot(1200, capture)
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
