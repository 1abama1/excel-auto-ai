"""Synthetic adapter, offline bundle, settings, cancellation and lifecycle regressions."""

from __future__ import annotations

import json
import threading
from pathlib import Path

import numpy as np
import pytest

from document_excel.domain.errors import AppError, Cancelled
from document_excel.infrastructure.ocr import TesseractEngine
from document_excel.infrastructure.ocr_factory import create_engine
from document_excel.infrastructure.paddle_models import CYRILLIC, ENGLISH, MODELS, sha256, validate_bundle
from document_excel.infrastructure.paddle_ocr import PaddleOCREngine, _ProcessPool, decode_prediction
from document_excel.infrastructure.storage import Settings


def test_selection_is_explicit_and_unavailable_paddle_does_not_fall_back(tmp_path):
    from document_excel.infrastructure.rapid_ocr import RapidOCREngine

    assert isinstance(create_engine(Settings()), RapidOCREngine)
    assert isinstance(create_engine(Settings(ocr_engine="paddle")), PaddleOCREngine)
    assert isinstance(create_engine(Settings(ocr_engine="tesseract")), TesseractEngine)
    engine = PaddleOCREngine(Settings(paddle_model_directory=str(tmp_path)))
    assert engine.diagnose().error
    with pytest.raises(AppError):
        engine.recognize(np.zeros((30, 100, 3), np.uint8), "synthetic", "original", threading.Event())


def test_v3_filtered_polygons_confidence_and_source():
    prediction = {
        "rec_texts": ["ТЕСТ", "TEST"],
        "rec_scores": [0.98, 0.5],
        "rec_polys": [
            [[10.2, 20], [75, 18.5], [77.3, 40], [10, 40]],
            [[90, 60], [150, 60], [150, 80], [90, 80]],
        ],
        "dt_polys": [np.zeros((4, 2))] * 3,
    }
    result = decode_prediction(prediction, "source-uuid", "crop", 140, 100)
    assert result.blocks[0].bounding_box == (10, 18, 68, 22)
    assert result.blocks[1].bbox == (90, 60, 50, 20)
    assert result.blocks[0].script == "cyrillic" and result.blocks[1].script == "latin"
    assert result.blocks[0].confidence == 0.98
    assert result.source_image == "source-uuid" and result.engine_name == "paddle"
    assert "PP-OCRv5" in result.model_version


@pytest.mark.parametrize(
    "scores,polys", [([0.9], []), ([float("nan")], [np.zeros((4, 2))]), ([0.9], [np.zeros((3, 2))])]
)
def test_malformed_v3_output_is_rejected(scores, polys):
    with pytest.raises(ValueError):
        decode_prediction(
            {"rec_texts": ["TEST"], "rec_scores": scores, "rec_polys": polys}, "x", "v", 100, 100
        )


def synthetic_bundle(directory: Path):
    records = {}
    for model in MODELS:
        target = directory / model
        target.mkdir()
        (target / "inference.json").write_text("{}")
        (target / "inference.pdiparams").write_bytes(b"SYNTHETIC MODEL")
        chars = (
            "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯҮӨҢүөң0123456789"
            if model == CYRILLIC
            else "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<"
        )
        # JSON is valid YAML, avoiding a dependency on Paddle in these unit tests.
        (target / "inference.yml").write_text(
            json.dumps({"Global": {"model_name": model}, "PostProcess": {"character_dict": list(chars)}})
        )
        records[model] = {
            name: sha256(target / name) for name in ("inference.json", "inference.yml", "inference.pdiparams")
        }
    manifest = directory / "manifest.json"
    manifest.write_text(json.dumps({"ocr_version": "PP-OCRv5", "models": records}))
    return manifest


def test_bundle_integrity_and_ky_alphabet(tmp_path):
    pytest.importorskip("yaml")
    synthetic_bundle(tmp_path)
    assert validate_bundle(tmp_path)["languages"] == ["ru", "ky", "en", "digits"]
    (tmp_path / ENGLISH / "inference.pdiparams").write_bytes(b"CORRUPTED")
    with pytest.raises(AppError):
        validate_bundle(tmp_path)


def test_native_ascii_path_is_preserved():
    from document_excel.infrastructure.paddle_models import native_model_root

    assert native_model_root("C:/local-models/paddle") == "C:/local-models/paddle"


def test_native_windows_model_path_uses_existing_short_alias(monkeypatch):
    from document_excel.infrastructure import paddle_models

    monkeypatch.setattr(paddle_models.sys, "platform", "win32")
    monkeypatch.setattr(paddle_models, "_windows_short_path", lambda _: "C:/MODELS~1/paddle")
    assert paddle_models.native_model_root("C:/модели/paddle") == "C:/MODELS~1/paddle"


def test_native_unicode_path_without_short_names_uses_verified_local_weights(tmp_path, monkeypatch):
    from document_excel.infrastructure import paddle_models

    pytest.importorskip("yaml")
    source = tmp_path / "модели"
    source.mkdir()
    synthetic_bundle(source)
    monkeypatch.setattr(paddle_models.sys, "platform", "win32")
    monkeypatch.setattr(paddle_models, "_windows_short_path", lambda path: path)
    monkeypatch.setenv("PUBLIC", str(tmp_path / "public"))
    result = Path(paddle_models.native_model_root(str(source)))
    assert str(result).isascii() and result != source
    assert validate_bundle(result)["version"] == "PP-OCRv5"
    # Partial or corrupted caches are repaired from the verified local bundle.
    (result / ENGLISH / "inference.pdiparams").write_bytes(b"BROKEN")
    assert Path(paddle_models.native_model_root(str(source))) == result
    assert validate_bundle(result)["version"] == "PP-OCRv5"
    assert {p.name for p in result.iterdir()} == set(MODELS) | {"manifest.json"}


def test_wrong_model_name_cannot_silently_replace_cyrillic(tmp_path):
    pytest.importorskip("yaml")
    manifest = synthetic_bundle(tmp_path)
    data = json.loads(manifest.read_text())
    config = tmp_path / CYRILLIC / "inference.yml"
    config.write_text(config.read_text().replace(CYRILLIC, "PP-OCRv5_server_rec"))
    data["models"][CYRILLIC]["inference.yml"] = sha256(config)
    manifest.write_text(json.dumps(data))
    with pytest.raises(AppError):
        validate_bundle(tmp_path)


def test_cancellation_before_model_loading():
    event = threading.Event()
    event.set()
    with pytest.raises(Cancelled):
        PaddleOCREngine(Settings()).recognize(np.zeros((30, 100, 3), np.uint8), "x", "v", event)


class FakeProcess:
    alive = True
    terminated = False
    closed = False

    def is_alive(self):
        return self.alive

    def terminate(self):
        self.alive, self.terminated = False, True

    def join(self, timeout):
        pass

    def close(self):
        self.closed = True


class WaitingConnection:
    def __init__(self, cancel=None):
        self.cancel = cancel
        self.closed = False

    def send(self, payload):
        pass

    def poll(self, timeout):
        if self.cancel:
            self.cancel.set()
        return False

    def close(self):
        self.closed = True


@pytest.mark.parametrize("cancelled", [True, False])
def test_cancel_and_timeout_terminate_inference_and_release_model_memory(cancelled, monkeypatch):
    pool, event = _ProcessPool(), threading.Event()
    process, connection = FakeProcess(), WaitingConnection(event if cancelled else None)
    pool.process, pool.connection, pool.key = process, connection, ("model", "cpu", 2)
    with pytest.raises(Cancelled if cancelled else AppError):
        pool.request(pool.key, ("synthetic",), event, 10 if cancelled else -1)
    assert process.terminated and process.closed and connection.closed
    assert pool.process is None and pool.key is None
    assert pool.lock.acquire(blocking=False)
    pool.lock.release()


def test_settings_engine_and_model_directory_round_trip(tmp_path):
    from document_excel.infrastructure.storage import SettingsRepository

    repo = SettingsRepository(tmp_path)
    settings = Settings(ocr_engine="paddle", paddle_model_directory="local models", paddle_cpu_threads=4)
    repo.save(settings)
    assert repo.load() == settings


def test_gui_engine_selection_round_trip(qtbot):
    from document_excel.presentation.dialogs import SettingsDialog

    dialog = SettingsDialog(Settings())
    qtbot.addWidget(dialog)
    assert dialog.value().ocr_engine == "rapidocr"
    dialog.engine.setCurrentIndex(dialog.engine.findData("tesseract"))
    assert dialog.value().ocr_engine == "tesseract"


def test_line_reader_trims_empty_margins_and_retains_actual_coordinates():
    from document_excel.infrastructure.paddle_ocr import _Backend

    shapes = []

    class Recognizer:
        def predict(self, pixels, batch_size=None):
            shapes.append(pixels.shape)
            return iter([{"rec_text": "SYNTHETIC", "rec_score": 0.83}])

    backend = object.__new__(_Backend)
    backend.recognizers = {ENGLISH: Recognizer()}
    pixels = np.full((90, 1200, 3), 255, np.uint8)
    pixels[25:60, 100:300] = 0
    read = backend.line(pixels, ENGLISH, "synthetic", "crop")
    assert shapes == [(41, 206, 3)]
    assert read.blocks[0].bbox == (97, 22, 206, 41)
    assert read.width == 1200 and read.blocks[0].confidence == 0.83
    empty = backend.line(np.full_like(pixels, 255), ENGLISH, "blank", "crop")
    assert not empty.blocks and len(shapes) == 1


def test_native_stage_timing_wraps_paddlex_cpu_pipeline_not_forwarding_facade(monkeypatch):
    import sys
    from types import SimpleNamespace

    from document_excel.infrastructure.paddle_ocr import _Backend, _TimedPredictor

    class Predictor:
        def __call__(self, *_):
            yield {}

    inner = SimpleNamespace(text_det_model=Predictor(), text_rec_model=Predictor())

    class Facade:
        _pipeline = inner

        def __getattr__(self, name):
            return getattr(self._pipeline, name)

    facade = Facade()

    def predict(pixels):
        list(inner.text_det_model(pixels))
        list(inner.text_rec_model(pixels))
        yield {
            "rec_texts": ["TEST"],
            "rec_scores": [0.99],
            "rec_polys": [[[0, 0], [40, 0], [40, 10], [0, 10]]],
        }

    wrapper = SimpleNamespace(paddlex_pipeline=facade, predict=predict)
    monkeypatch.setitem(
        sys.modules, "paddleocr", SimpleNamespace(PaddleOCR=lambda **_: wrapper, TextRecognition=Predictor)
    )
    backend = _Backend("synthetic", "cpu", 2)
    assert isinstance(inner.text_det_model, _TimedPredictor)
    assert "text_det_model" not in vars(facade)
    assert backend.recognizers[CYRILLIC] is inner.text_rec_model
    result = backend.document(np.zeros((20, 60, 3), np.uint8), "synthetic", "original")
    assert result.timings["text_detection_seconds"] > 0
    assert result.timings["text_recognition_seconds"] > 0


def test_cached_result_is_isolated_by_client_image_and_variant():
    from collections import OrderedDict

    from document_excel.infrastructure.paddle_ocr import _Backend

    calls = []

    class Pipeline:
        def predict(self, pixels):
            calls.append(pixels.copy())
            return iter(
                [
                    {
                        "rec_texts": ["TEST"],
                        "rec_scores": [0.91],
                        "rec_polys": [[[0, 0], [50, 0], [50, 10], [0, 10]]],
                    }
                ]
            )

    backend = object.__new__(_Backend)
    backend.cache, backend.pipeline = OrderedDict(), Pipeline()
    backend.stage_times = {}
    pixels = np.zeros((20, 80, 3), np.uint8)
    first = backend.recognize(pixels, "old-client-image", "original", "document")
    second = backend.recognize(pixels.copy(), "new-client-image", "crop", "document")
    assert len(calls) == 2
    assert first.image_id == "old-client-image" and second.image_id == "new-client-image"
    assert first.variant == "original" and second.variant == "crop"
    changed = pixels.copy()
    changed[0, 0] = 255
    backend.recognize(changed, "new-client-image", "crop", "document")
    assert len(calls) == 3
    backend.recognize(changed, "new-client-image", "crop", "document")
    assert len(calls) == 3
    backend.recognize(changed, "new-client-image", "crop", "document", "changed-profile")
    assert len(calls) == 4
    backend.recognize(np.zeros((1, 1), np.uint8), "", "", "clear")
    assert not backend.cache
