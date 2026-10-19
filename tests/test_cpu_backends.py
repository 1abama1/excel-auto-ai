"""CPU-only admission, offline bundles and safe reverse-side recovery."""

import hashlib
import json
import threading

import numpy as np
import pytest
from test_fast_recognition import run, setup  # noqa: F401
from test_mrz_physical import td1

from document_excel.application import fast_recognition as fast
from document_excel.domain.errors import AppError
from document_excel.domain.models import DocumentSide
from document_excel.infrastructure import paddle_ocr, rapid_models, rapid_ocr
from document_excel.infrastructure.mrz import MRZRecognizer, TD1Parser
from document_excel.infrastructure.ocr_contract import OcrResult, TextBlock
from document_excel.infrastructure.ocr_factory import create_engine
from document_excel.infrastructure.storage import Settings, SettingsRepository


@pytest.mark.parametrize(
    "engine_class, label", [(paddle_ocr.PaddleOCREngine, "paddle"), (rapid_ocr.RapidOCREngine, "rapidocr")]
)
def test_mrz_rows_keep_the_actual_cpu_backend_provenance(monkeypatch, engine_class, label):
    engine = engine_class(Settings())
    monkeypatch.setattr(MRZRecognizer, "line_boxes", lambda _: [])
    monkeypatch.setattr(
        engine,
        "recognize_mrz_line",
        lambda *_: OcrResult(
            [TextBlock("SYNTHETIC", 0.9, (0, 0, 100, 10), (1, 1, 1))],
            "id",
            "mrz",
            100,
            10,
            label,
            "PP-OCRv5 test",
        ),
    )
    result = engine.recognize_mrz(np.zeros((30, 100, 3), np.uint8), "id", threading.Event())
    assert result.engine_name == label and result.model_version == "PP-OCRv5 test"
    assert len(result.blocks) == 3


def test_cpu_engine_selection_is_explicit_and_settings_round_trip(tmp_path):
    settings = Settings(ocr_engine="rapidocr", rapid_model_directory="local-onnx", paddle_cpu_threads=2)
    assert isinstance(create_engine(settings), rapid_ocr.RapidOCREngine)
    assert isinstance(create_engine(Settings(ocr_engine="paddle")), paddle_ocr.PaddleOCREngine)
    repo = SettingsRepository(tmp_path)
    repo.save(settings)
    assert repo.load() == settings


def test_legacy_gpu_preference_becomes_cpu_without_losing_settings(tmp_path):
    repo = SettingsRepository(tmp_path)
    repo.path.write_text(json.dumps({"paddle_device": "gpu:0", "theme": "dark", "paddle_cpu_threads": 2}))
    loaded = repo.load()
    assert loaded.paddle_device == "cpu" and loaded.theme == "dark"


def test_worker_refuses_gpu_before_loading_any_model(tmp_path):
    with pytest.raises(AppError, match="CPU"):
        paddle_ocr._create_backend(str(tmp_path), "gpu:0", 2)


def test_backend_dispatch_never_silently_falls_back(tmp_path, monkeypatch):
    marker = object()
    monkeypatch.setattr(rapid_ocr, "_RapidBackend", lambda *a: marker)
    (tmp_path / "manifest.json").write_text(json.dumps({"backend": "rapidocr"}))
    assert paddle_ocr._create_backend(str(tmp_path), "cpu", 2) is marker
    monkeypatch.setattr(paddle_ocr, "_Backend", lambda *a: "synthetic-paddle")
    (tmp_path / "manifest.json").write_text(json.dumps({"ocr_version": "PP-OCRv5"}))
    assert paddle_ocr._create_backend(str(tmp_path), "cpu", 2) == "synthetic-paddle"


def test_offline_onnx_bundle_rejects_changed_weights_and_wrong_manifest(tmp_path, monkeypatch):
    payload = b"SYNTHETIC ONNX PAYLOAD"
    name = "synthetic.onnx"
    monkeypatch.setattr(rapid_models, "MODEL_FILES", {"det": (name, hashlib.sha256(payload).hexdigest())})
    (tmp_path / name).write_bytes(payload)
    manifest = tmp_path / "manifest.json"
    manifest.write_text(json.dumps({"backend": "rapidocr", "ocr_version": "PP-OCRv5"}))
    rapid_models.validate_bundle(tmp_path)
    (tmp_path / name).write_bytes(b"CHANGED MODEL")
    with pytest.raises(AppError):
        rapid_models.validate_bundle(tmp_path)
    (tmp_path / name).write_bytes(payload)
    manifest.write_text(json.dumps({"backend": "paddle", "ocr_version": "PP-OCRv5"}))
    with pytest.raises(AppError):
        rapid_models.validate_bundle(tmp_path)


def test_onnx_provider_must_be_cpu_only():
    from types import SimpleNamespace

    for providers in (["CUDAExecutionProvider", "CPUExecutionProvider"], ["DmlExecutionProvider"]):
        predictor = SimpleNamespace(
            session=SimpleNamespace(
                session=SimpleNamespace(get_providers=lambda providers=providers: providers)
            )
        )
        with pytest.raises(AppError, match="CPUExecutionProvider"):
            rapid_ocr._RapidBackend._check_provider(predictor)
    predictor = SimpleNamespace(
        session=SimpleNamespace(session=SimpleNamespace(get_providers=lambda: ["CPUExecutionProvider"]))
    )
    rapid_ocr._RapidBackend._check_provider(predictor)


@pytest.mark.parametrize("valid", [True, False])
def test_checked_td1_recovers_unknown_side_without_file_order(setup, monkeypatch, valid):  # noqa: F811
    service, engine, source = setup
    source.role = DocumentSide.UNKNOWN
    engine.recognize_mrz_line = lambda *a: None
    monkeypatch.setattr(fast.DocumentSideClassifier, "classify", lambda *a: DocumentSide.UNKNOWN)
    monkeypatch.setattr(
        MRZRecognizer, "line_boxes", lambda *a: [(0, 400, 900, 40), (0, 450, 900, 40), (0, 500, 900, 40)]
    )
    lines = td1()
    if not valid:
        lines[0] = lines[0][:14] + ("1" if lines[0][14] != "1" else "2") + lines[0][15:]
    report = TD1Parser().parse(lines, source.id)
    assert report.valid == valid
    calls = []

    def recognize(*a):
        calls.append(True)
        return report.model_copy(deep=True)

    monkeypatch.setattr(MRZRecognizer, "recognize", recognize)
    output = run(service, source)
    assert output.sides[source.id] == (DocumentSide.BACK if valid else DocumentSide.UNKNOWN)
    assert len(calls) == 1
    assert not any(
        c.original_cyrillic
        for c in output.candidates
        if c.field_name in {"surname", "first_name", "patronymic"}
    )
