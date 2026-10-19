"""Explicit opt-in local control. Never captures images or writes recognized text."""

import json
import os
import re
import time
from pathlib import Path

import pytest
from PySide6.QtCore import QTimer

from document_excel.application.fast_recognition import clear_cache
from document_excel.domain.models import ClientCase, FieldEvidence, FieldStatus
from document_excel.domain.resolution import displayed_candidate
from document_excel.infrastructure.paddle_ocr import _POOL, PaddleOCREngine
from document_excel.infrastructure.storage import Settings, SettingsRepository
from document_excel.presentation.main_window import MainWindow


@pytest.mark.paddle
def test_private_control_gui(qtbot, tmp_path, monkeypatch):
    paths, request = os.environ.get("LOCAL_CONTROL_IMAGES"), os.environ.get("LOCAL_CONTROL_REQUEST")
    if not paths or not request:
        pytest.skip("Explicit local images and control request are required")
    text = Path(request).read_text(encoding="utf-8")
    expected = {}
    for name, heading in (("surname", "Фамилия"), ("first_name", "Имя")):
        group = re.search(r"\*\*" + heading + r":\*\*(.*?)(?=\*\*|\Z)", text, re.S)[1]
        for script, label in (("cyrillic", "Кириллица"), ("latin", "Латиница")):
            expected[name + "_" + script] = re.search(label + r":\s*([^\r\n]+)", group)[1].strip()
    expected["patronymic_cyrillic"] = re.search(r"\*\*Отчество:\*\*\s*-\s*([^\r\n]+)", text)[1].strip()
    clear_cache()
    _POOL.close()
    monkeypatch.setattr(MainWindow, "diagnose", lambda self: None)
    window = MainWindow(SettingsRepository(tmp_path / "app-data"))
    window.settings = Settings(
        ocr_engine=os.environ.get("LOCAL_CONTROL_ENGINE", "paddle"),
        paddle_model_directory=os.environ.get("DOCUMENT_EXCEL_PADDLE_MODELS", ""),
        rapid_model_directory=os.environ.get("DOCUMENT_EXCEL_RAPID_MODELS", ""),
        timeout=120,
    )
    qtbot.addWidget(
        window,
        before_close_func=lambda _: qtbot.waitUntil(lambda: not window.background_workers, timeout=30000),
    )
    errors = []
    monkeypatch.setattr(window, "error", lambda *_: errors.append(True))
    window.show()
    window.case_actions.add_images([Path(p) for p in json.loads(paths)])
    original = window.current_case
    original.data.notes = "SYNTHETIC MANUAL NOTE"
    original.evidence["notes"] = FieldEvidence(source="ручной ввод", status=FieldStatus.MANUALLY_CORRECTED)
    calls = []
    native = PaddleOCREngine._read

    def read(self, *a, **kw):
        calls.append(kw.get("mode", a[4] if len(a) > 4 else "document"))
        return native(self, *a, **kw)

    monkeypatch.setattr(PaddleOCREngine, "_read", read)
    ticks = []
    timer = QTimer()
    timer.timeout.connect(lambda: ticks.append(time.perf_counter()))
    timer.start(10)
    started = time.perf_counter()
    window.case_actions.start_ocr()
    window.batch.cases.append(ClientCase())
    window.select_case(1)
    qtbot.waitUntil(lambda: window.active_worker is None or bool(errors), timeout=120000)
    cold = time.perf_counter() - started
    timer.stop()
    checks = {k: getattr(original.data, k) == v for k, v in expected.items()}
    assert not errors and all(checks.values()), checks
    assert len(ticks) > 20
    assert not window.current_case.data.full_name
    assert original.data.notes == "SYNTHETIC MANUAL NOTE"
    assert not original.confirmed
    window.select_case(0)
    provenance = {}
    for name in ("surname", "first_name", "patronymic"):
        candidate = displayed_candidate(original.evidence[name], getattr(original.data, name))
        provenance[name] = bool(
            candidate and str(candidate.source_side) == "FRONT" and candidate.source_method == "VISUAL_OCR"
        )
    assert all(provenance.values()), provenance
    pid = _POOL.process.pid
    count = len(calls)
    for index in (1, 0, 1, 0):
        window.image_list.setCurrentIndex(index)
    qtbot.wait(100)
    assert len(calls) == count
    started = time.perf_counter()
    window.case_actions.start_ocr()
    qtbot.waitUntil(lambda: window.active_worker is None, timeout=10000)
    warm = time.perf_counter() - started
    assert len(calls) == count and _POOL.process.pid == pid
    started = time.perf_counter()
    window.refresh_form()
    refresh = time.perf_counter() - started
    # Read one name from a deliberately narrow previous word box. The saved
    # FRONT group hint must expand it, and other fields must remain intact.
    candidate = next(c for c in original.evidence["surname"].candidates if c.source_method == "MRZ")
    candidate = candidate.model_copy(
        update={"source_bbox": (*candidate.source_bbox[:2], 10, candidate.source_bbox[3])}
    )
    snapshot = original.data.first_name
    window.case_actions.repeat_field("surname_cyrillic", candidate)
    qtbot.waitUntil(lambda: window.active_worker is None, timeout=30000)
    repeat_checks = {
        "surname": original.data.surname_cyrillic == expected["surname_cyrillic"],
        "other_name": original.data.first_name == snapshot,
    }
    assert all(repeat_checks.values()), repeat_checks
    assert _POOL.process.pid == pid
    assert all(mode != "document" for mode in calls[count:])
    report = {
        "cold_seconds": cold,
        "warm_seconds": warm,
        "refresh_seconds": refresh,
        "heartbeat_ticks": len(ticks),
        "max_heartbeat_gap_seconds": max((b - a for a, b in zip(ticks, ticks[1:], strict=False)), default=0),
        "ocr_calls": count,
        "repeat_field_calls": len(calls) - count,
        "checks": checks,
        "provenance": provenance,
        "repeat_checks": repeat_checks,
        "model_process_reused": True,
    }
    Path("test-artifacts/private-gui.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    _POOL.close()
