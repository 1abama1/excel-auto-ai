"""Opt-in real CPU engines -> Qt worker, exclusively generated pixels."""

import os

import pytest
from PySide6.QtCore import QTimer

from document_excel.domain.models import FieldEvidence, FieldStatus
from document_excel.infrastructure.paddle_ocr import PaddleOCREngine
from document_excel.infrastructure.rapid_ocr import RapidOCREngine
from document_excel.infrastructure.storage import Settings, SettingsRepository
from document_excel.presentation.main_window import MainWindow


@pytest.mark.paddle
@pytest.mark.parametrize("engine", ["paddle", "rapidocr"])
def test_real_cpu_worker_keeps_gui_responsive_and_manual_data(
    qtbot, tmp_path, image_factory, monkeypatch, engine
):
    root = os.environ.get(
        "DOCUMENT_EXCEL_RAPID_MODELS" if engine == "rapidocr" else "DOCUMENT_EXCEL_PADDLE_MODELS"
    )
    if not root:
        pytest.skip("Set the corresponding model directory to opt in to real offline CPU inference")
    settings = Settings(
        ocr_engine=engine, paddle_model_directory=root, rapid_model_directory=root, timeout=120
    )
    diagnostic = (RapidOCREngine if engine == "rapidocr" else PaddleOCREngine)(settings).diagnose()
    assert not diagnostic.error, diagnostic.error
    monkeypatch.setattr(MainWindow, "diagnose", lambda self: None)
    window = MainWindow(SettingsRepository(tmp_path / "app-data"))
    window.settings = settings
    errors = []
    monkeypatch.setattr(window, "error", errors.append)
    qtbot.addWidget(
        window,
        before_close_func=lambda _: qtbot.waitUntil(lambda: not window.background_workers, timeout=30000),
    )
    window.show()
    path = image_factory(
        lines=[
            "ТҮНДҮК / ЦИФРОВОЕ / KYRGYZ REPUBLIC",
            "ФАМИЛИЯ: СИНТЕТИКОВ",
            "ИМЯ: ПРИМЕР",
            "ДАТА РОЖДЕНИЯ: 01.02.1990",
            "ПИН: 00123456789012",
            "АДРЕС РЕГИСТРАЦИИ: ТЕСТОВАЯ ОБЛАСТЬ, ПРИМЕР 1",
        ]
    )
    window.case_actions.add_images([path])
    window.current_case.data.notes = "SYNTHETIC MANUAL NOTE"
    window.current_case.evidence["notes"] = FieldEvidence(
        source="ручной ввод", status=FieldStatus.MANUALLY_CORRECTED
    )
    ticks = []
    timer = QTimer()
    timer.timeout.connect(lambda: ticks.append(True))
    timer.start(10)
    window.case_actions.start_ocr()
    qtbot.waitUntil(lambda: window.active_worker is None or bool(errors), timeout=120000)
    timer.stop()
    assert not errors, errors
    assert len(ticks) > 5
    assert window.current_case.data.personal_identifier == "00123456789012"
    assert window.current_case.data.notes == "SYNTHETIC MANUAL NOTE"
    assert not window.current_case.confirmed
    assert all(
        c.engine_name == engine for c in window.current_case.evidence["personal_identifier"].candidates
    )
    window.grab().save(str(tmp_path / f"{engine}-live.png"))
