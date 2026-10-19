"""An office-code zero must be observed in an independent optical read."""

import threading

import numpy as np
import pytest
from test_fast_recognition import setup  # noqa: F401

from document_excel.application import fast_recognition as fast
from document_excel.domain.errors import AppError, Cancelled
from document_excel.domain.models import FieldCandidate, FieldStatus
from document_excel.infrastructure.identifier_ocr import refine_mkk_authority
from document_excel.infrastructure.ocr_contract import OcrResult, TextBlock


def candidate(value="MKKO00123"):
    return FieldCandidate(
        field_name="passport_issued_by",
        raw_value=value,
        normalized_value=value,
        confidence=0.92,
        source_image_id="synthetic",
        source_bbox=(10, 20, 160, 24),
    )


def result(text, confidence=0.98):
    return OcrResult(
        [TextBlock(text, confidence, (0, 0, 160, 24), (1, 1, 1))],
        "synthetic",
        "authority-line",
        160,
        24,
        "rapidocr",
        "English PP-OCRv5",
    )


def test_authority_zero_is_read_from_one_roi_and_other_fields_are_preserved():
    issued = candidate()
    pin = issued.model_copy(
        update={
            "field_name": "personal_identifier",
            "raw_value": "00123456789012",
            "normalized_value": "00123456789012",
        }
    )
    calls = []

    def read(pixels, *_):
        calls.append(pixels.shape)
        return result("AUTHORITY: MKK000123")

    found = refine_mkk_authority(
        np.zeros((100, 300, 3), np.uint8), result("unused"), [issued, pin], read, threading.Event()
    )
    values = {c.field_name: c for c in found}
    assert calls == [(30, 168, 3)]
    assert values["passport_issued_by"].normalized_value == "MKK"
    assert values["passport_department_code"].normalized_value == "000123"
    assert values["personal_identifier"] == pin
    assert values["passport_issued_by"].engine_name == "rapidocr"
    assert values["passport_department_code"].validation_status == FieldStatus.REVIEW


@pytest.mark.parametrize(
    "text, confidence",
    [
        ("MKKO00123", 0.99),
        ("MKK000124", 0.99),
        ("MKK100123", 0.99),
        ("MKK000123", 0.4),
        ("MKK00123", 0.99),
        ("OTHER MKK000123", 0.99),
    ],
)
def test_failed_or_different_office_code_is_never_guessed(text, confidence):
    initial = [candidate()]
    found = refine_mkk_authority(
        np.zeros((100, 300, 3), np.uint8),
        result("unused"),
        initial,
        lambda *_: result(text, confidence),
        threading.Event(),
    )
    assert found is initial


@pytest.mark.parametrize("values", [["MKK000123"], ["OTHER00123"], ["MKKO00123", "MKKO00456"], []])
def test_known_codes_other_authorities_and_conflicts_need_no_extra_pass(values):
    def never(*_):
        pytest.fail("Unexpected extra OCR")

    initial = [candidate(v) for v in values]
    assert (
        refine_mkk_authority(
            np.zeros((100, 300, 3), np.uint8), result("unused"), initial, never, threading.Event()
        )
        is initial
    )


@pytest.mark.parametrize("exception", [AppError, Cancelled])
def test_generic_optional_read_failure_keeps_fields_or_propagates_cancel(setup, monkeypatch, exception):  # noqa: F811
    service, engine, source = setup
    engine.recognize_identifier_line = lambda *_: result("unused")

    def failed(*_):
        if exception is Cancelled:
            raise Cancelled()
        raise AppError("Synthetic failure")

    monkeypatch.setattr(fast, "refine_mkk_authority", failed)
    if exception is Cancelled:
        with pytest.raises(Cancelled):
            service.run([source], threading.Event(), lambda *_: None, document_type="kg_id")
    else:
        output = service.run([source], threading.Event(), lambda *_: None, document_type="kg_id")
        assert output.candidates and output.warnings
    assert not fast._CACHE
