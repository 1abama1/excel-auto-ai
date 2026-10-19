"""Opt-in OCR of generated phone screens; no private fixtures or values."""

import os
import threading

import pytest

from document_excel.application.workflows import RecognitionService
from document_excel.domain.models import ClientCase, SourceImage
from document_excel.infrastructure.digital_verification import synthetic_digital_images
from document_excel.infrastructure.extraction import apply_candidates
from document_excel.infrastructure.storage import Settings


@pytest.mark.paddle
@pytest.mark.parametrize("backend", ["rapidocr", "paddle"])
@pytest.mark.parametrize("variant", ["normal", "small", "overlay", "wrapped_names"])
def test_generated_digital_screens_preserve_critical_fields(tmp_path, backend, variant):
    rapid = os.environ.get("DOCUMENT_EXCEL_RAPID_MODELS", "")
    paddle = os.environ.get("DOCUMENT_EXCEL_PADDLE_MODELS", "")
    if not (rapid if backend == "rapidocr" else paddle):
        pytest.skip("Supply local model directories for offline inference")
    paths, expected = synthetic_digital_images(tmp_path, **({variant: True} if variant != "normal" else {}))
    case = ClientCase()
    case.images = [SourceImage(path=p, case_id=case.id) for p in paths]
    service = RecognitionService(
        Settings(ocr_engine=backend, rapid_model_directory=rapid, paddle_model_directory=paddle, timeout=120)
    )
    output = service.run(case.images, threading.Event(), lambda *_: None)
    apply_candidates(case, output.candidates, output.document_type, output.label_seen)
    assert output.document_type == "kg_digital_id"
    for field, wanted in expected.items():
        assert str(getattr(case.data, field)) == wanted, (backend, variant, field)
    assert case.data.registration_address.raw and "Примерная" in case.data.registration_address.raw
    assert case.registration_status != "NOT_PRESENT"
