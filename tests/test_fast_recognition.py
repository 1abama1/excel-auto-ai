import threading
from types import SimpleNamespace

import numpy as np
import pytest
from test_full_name_groups import front

from document_excel.application import fast_recognition as fast
from document_excel.application.workflows import RecognitionService
from document_excel.domain.errors import AppError, Cancelled
from document_excel.domain.models import ClientCase, DocumentSide, SourceImage
from document_excel.infrastructure.extraction import apply_candidates
from document_excel.infrastructure.physical_id import PlasticCard
from document_excel.infrastructure.storage import Settings


@pytest.fixture
def setup(tmp_path, monkeypatch):
    fast.clear_cache()
    pixels = np.full((640, 1000, 3), 255, np.uint8)
    monkeypatch.setattr(
        fast, "prepare_plastic_card", lambda _: PlasticCard(pixels, np.eye(3), (1000, 640), False)
    )

    class Engine:
        preferred_input = "color"
        calls = 0

        def recognize(self, image, image_id, variant, cancel):
            self.calls += 1
            result = front()
            result.image_id = image_id
            return result

    engine = Engine()
    service = RecognitionService(Settings(), engine)
    service.preprocessor = SimpleNamespace(
        prepare=lambda *_: SimpleNamespace(original=pixels, quality=SimpleNamespace(warnings=[]))
    )
    path = tmp_path / "synthetic.png"
    path.write_bytes(b"synthetic-image-v1")
    source = SourceImage(path=path, case_id="case-a")
    yield service, engine, source
    fast.clear_cache()


def run(service, source):
    return service.run([source], threading.Event(), lambda *_: None, document_type="kg_physical_id")


def test_one_document_pass_and_cached_reopen_across_services(setup):
    service, engine, source = setup
    first = run(service, source)
    assert engine.calls == 1
    case = ClientCase(images=[source])
    apply_candidates(case, first.candidates, first.document_type)
    assert source.role == DocumentSide.FRONT
    # Resolution's derived role does not force another inference.
    second_service = RecognitionService(service.settings, engine)
    second_service.preprocessor = service.preprocessor
    second = run(second_service, source)
    assert engine.calls == 1 and second.candidates == first.candidates
    second.candidates.clear()
    assert run(service, source).candidates


def test_empty_digital_crop_retries_untouched_original_once(setup, monkeypatch):
    from document_excel.infrastructure.digital_region import DigitalIDRegion
    from document_excel.infrastructure.ocr_contract import OcrResult, TextBlock

    service, engine, source = setup
    pixels = service.preprocessor.prepare().original
    monkeypatch.setattr(
        fast.DigitalIDRegionDetector,
        "prepare",
        lambda *_, **__: DigitalIDRegion(
            pixels[100:200, 100:200], (1000, 640), (100, 100, 100, 100), True, 0.9, True
        ),
    )
    seen = []

    def read(image, image_id, variant, cancel):
        seen.append(image.shape)
        blocks = (
            []
            if len(seen) == 1
            else [
                TextBlock("Дата рождения", 0.99, (200, 300, 200, 20), (1, 1, 1)),
                TextBlock("07.11.1990", 0.99, (200, 330, 200, 25), (1, 1, 2)),
                TextBlock("KYRGYZ REPUBLIC", 0.99, (100, 40, 300, 20), (1, 1, 3)),
            ]
        )
        return OcrResult(blocks, image_id, variant, image.shape[1], image.shape[0])

    engine.recognize = read
    output = service.run([source], threading.Event(), lambda *_: None, document_type="kg_digital_id")
    assert seen == [(100, 100, 3), (640, 1000, 3)]
    assert output.processing_metrics[source.id]["original_fallback_attempted"]
    assert any(c.field_name == "birth_date" and c.normalized_value == "1990-11-07" for c in output.candidates)


@pytest.mark.parametrize("technical,code", [(True, "OCR_ENGINE_ERROR"), (False, "NO_TEXT_DETECTED")])
def test_digital_technical_failure_and_empty_text_have_different_codes(setup, technical, code):
    from document_excel.domain.errors import OcrFailure
    from document_excel.infrastructure.ocr_contract import OcrResult

    service, engine, source = setup

    def read(image, image_id, variant, cancel):
        if technical:
            raise AppError("Synthetic offline model failure")
        return OcrResult([], image_id, variant, image.shape[1], image.shape[0])

    engine.recognize = read
    with pytest.raises(OcrFailure) as error:
        service.run([source], threading.Event(), lambda *_: None, document_type="kg_digital_id")
    assert error.value.code == code and not fast._CACHE


@pytest.mark.parametrize(
    "change",
    ["client", "image_id", "bytes", "rotation", "processed", "parameters", "profile", "preprocessing"],
)
def test_cache_invalidates_only_when_relevant_input_changes(setup, monkeypatch, change):
    service, engine, source = setup
    run(service, source)
    if change == "client":
        source.case_id = "case-b"
    elif change == "image_id":
        source.id = "replacement"
    elif change == "bytes":
        source.path.write_bytes(b"synthetic-image-v2")
    elif change == "rotation":
        source.rotation = 90
    elif change == "processed":
        source.processed_path = source.path
        source.processed_confirmed = True
    elif change == "parameters":
        service.settings.paddle_cpu_threads = 4
    elif change == "profile":
        original = fast.load_document_profiles

        def changed():
            profiles = original()
            profiles[0].signatures.append("SYNTHETIC SIGNATURE")
            return profiles

        monkeypatch.setattr(fast, "load_document_profiles", changed)
    else:
        monkeypatch.setattr(fast, "PREPROCESSING_VERSION", "synthetic-new-preprocessing")
    run(service, source)
    assert engine.calls == 2


def test_model_files_are_part_of_cache_key(setup, tmp_path):
    service, engine, source = setup
    engine.root = tmp_path / "model"
    engine.root.mkdir()
    model = engine.root / "inference.pdiparams"
    model.write_bytes(b"synthetic-model-v1")
    run(service, source)
    model.write_bytes(b"synthetic-model-v2-longer")
    run(service, source)
    assert engine.calls == 2


def test_cached_result_still_honors_cancellation_and_clear(setup):
    service, engine, source = setup
    run(service, source)
    cancel = threading.Event()
    cancel.set()
    with pytest.raises(Cancelled):
        service.run([source], cancel, lambda *_: None, document_type="kg_physical_id")
    fast.clear_cache()
    run(service, source)
    assert engine.calls == 2


def test_unknown_upside_down_text_gets_one_orientation_recovery(setup):
    service, engine, source = setup
    native = engine.recognize

    def recognize(image, image_id, variant, cancel):
        if engine.calls == 0:
            engine.calls += 1
            from document_excel.infrastructure.ocr_contract import OcrResult

            return OcrResult([], image_id, variant, 1000, 640)
        return native(image, image_id, variant, cancel)

    engine.recognize = recognize
    output = run(service, source)
    assert output.sides[source.id] == DocumentSide.FRONT and engine.calls == 2


def test_optional_retry_failure_keeps_primary_fields_and_is_not_cached(setup, monkeypatch):
    service, engine, source = setup

    def unavailable(*_args, **_kwargs):
        raise AppError("Synthetic inference unavailable")

    monkeypatch.setattr(fast, "layout_candidates", unavailable)
    output = run(service, source)
    assert output.candidates and output.warnings and not fast._CACHE
    assert any(c.normalized_value == "АЛЕКСАНДР" for c in output.candidates)


def test_memory_cache_is_bounded_and_evicts_old_clients(setup):
    service, engine, source = setup
    for i in range(20):
        source.case_id = f"synthetic-case-{i}"
        run(service, source)
    assert len(fast._CACHE) <= 16
    source.case_id = "synthetic-case-0"
    run(service, source)
    assert engine.calls == 21


@pytest.mark.parametrize("mode", ["grayscale", "adaptive", "denoised", "sharpened", "contrast"])
def test_explicit_preprocessing_keeps_geometry_and_one_main_pass(setup, mode):
    service, engine, source = setup
    service.settings.preprocessing = mode
    native = engine.recognize
    shapes = []

    def recognize(image, *args):
        shapes.append(image.shape)
        return native(image, *args)

    engine.recognize = recognize
    run(service, source)
    assert shapes == [(640, 1000)] and engine.calls == 1
