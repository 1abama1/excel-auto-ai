"""Local integration check with synthetic inputs; never prints template values."""

from __future__ import annotations

import json
import os
from datetime import date
from pathlib import Path
from uuid import uuid4

from document_excel.domain.models import (
    Address,
    ClientCase,
    ClientData,
    ContractData,
    DocumentSide,
    RegistrationStatus,
    SourceImage,
    ToolItem,
)
from document_excel.infrastructure.excel import WorkbookService, file_hash, legacy_profile

ROOT = Path(__file__).resolve().parents[1]


def main():
    source = ROOT / "assets/lermontov.xlsx"
    output = ROOT / "test-artifacts" / "supplied-template" / uuid4().hex[:8]
    output.mkdir(parents=True)
    original_hash = file_hash(source)
    service = WorkbookService()
    sheets, profile = service.describe(source, legacy_profile())
    case = ClientCase(
        data=ClientData(
            full_name="SYNTHETIC TEST CLIENT",
            personal_identifier="00123456789012",
            passport_type="ID",
            passport_number="0000001",
            passport_issued_by="MKK",
            passport_department_code="000123",
            passport_issued_date=date(2020, 3, 4),
            birth_date=date(1990, 1, 2),
            whatsapp_phone="+996 555 000001",
            registration_address=Address(region="TEST REGION", street="TEST STREET 1"),
            living_address=Address(region="TEST REGION", street="TEST STREET 2"),
            object_address="TEST OBJECT ADDRESS",
            document_type="kg_id",
        ),
        contract=ContractData(
            tools=[
                ToolItem(
                    name=f"TEST {i + 1}", inventory_number=f"000{i + 1}", quantity=1, price_per_day=100 + i
                )
                for i in range(7)
            ],
            deposit_amount=500,
            purchase_price=999,
        ),
        registration_status=RegistrationStatus.MANUALLY_ENTERED,
    )
    from PIL import Image, ImageDraw

    for side, size, color in [
        (DocumentSide.FRONT, (1200, 750), "lightblue"),
        (DocumentSide.BACK, (900, 1200), "lightgreen"),
    ]:
        path = output / f"synthetic-{side.value.lower()}.png"
        image = Image.new("RGB", size, color)
        draw = ImageDraw.Draw(image)
        draw.rectangle((8, 8, size[0] - 9, size[1] - 9), outline="black", width=5)
        draw.text((40, 40), "SYNTHETIC " + side.value + " / NOT A DOCUMENT", fill="black", font_size=32)
        image.save(path)
        case.images.append(SourceImage(path=path, role=side))
    case.confirmed_digest = case.digest()
    plan = service.plan(source, case, profile)
    target = service.write(plan, output / "synthetic-contract.xlsx", case, profile)
    assert file_hash(source) == original_hash
    report = {
        "source_unchanged": True,
        "source_sha256": original_hash,
        "sheet_count": len(sheets),
        "profile": profile.id,
        "changed_cells": len(plan.changes),
        "new_sheets": len(plan.new_sheets),
        "embedded_photos": len(plan.photos.insertions),
        "output": str(target),
    }
    (output / "report.json").write_text(json.dumps(report, indent=2, ensure_ascii=True), encoding="utf-8")
    # Render only synthetic input rows, without original clients from other tabs.
    os.environ["QT_QPA_PLATFORM"] = "offscreen"
    from PySide6.QtCore import QTimer
    from PySide6.QtWidgets import QApplication, QFileDialog

    from document_excel.infrastructure.storage import SettingsRepository
    from document_excel.presentation.main_window import MainWindow
    from document_excel.presentation.theme import apply_theme

    app = QApplication([])
    window = MainWindow(SettingsRepository(output / "app-data"))
    window.current_case.data = case.data.model_copy(deep=True)
    window.current_case.contract = case.contract.model_copy(deep=True)
    window.current_case.registration_status = case.registration_status
    window.current_case.images = [s.model_copy(deep=True) for s in case.images]
    apply_theme(app, window.settings)
    window.resize(1120, 804)
    QFileDialog.getOpenFileName = lambda *_args, **_kwargs: (str(source), "Excel (*.xlsx)")
    errors = []
    window.error = lambda error: errors.append(type(error).__name__)
    window.show()

    def capture():
        if window.background_workers:
            QTimer.singleShot(100, capture)
            return
        assert not errors, errors
        assert window.profile.id == "lermontov-combined"
        assert window.template_sheets[source] == sheets
        window.go_step(3)
        window.refresh_form()
        window.grab().save(str(output / "excel-selected.png"))
        window.close()
        app.quit()

    QTimer.singleShot(100, window.excel_actions.choose_template)
    QTimer.singleShot(1000, capture)
    QTimer.singleShot(20000, app.quit)
    app.exec()
    print(json.dumps(report, ensure_ascii=True))


if __name__ == "__main__":
    main()
