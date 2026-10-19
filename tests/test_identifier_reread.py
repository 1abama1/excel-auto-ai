"""Synthetic regression cases for ID glyphs read as numeric 10."""

import threading

import numpy as np
import pytest

from document_excel.domain.errors import Cancelled
from document_excel.domain.models import ClientCase, FieldStatus
from document_excel.infrastructure.extraction import FieldExtractor, apply_candidates
from document_excel.infrastructure.identifier_ocr import refine_identifier
from document_excel.infrastructure.ocr import OcrDiagnostic, OcrResult, TesseractEngine, TextBlock
from document_excel.infrastructure.storage import Settings


def ambiguous_result(value="100000001", inline=False):
    words = [TextBlock("DOCUMENT NUMBER:", 0.95, (20, 20, 180, 20), (1, 1, 1))]
    words.append(
        TextBlock(value, 0.91, (220, 20, 100, 20), (1, 1, 1))
        if inline
        else TextBlock(value, 0.91, (20, 50, 100, 20), (1, 1, 2))
    )
    result = OcrResult(words, "synthetic", "original", 400, 200)
    return result, FieldExtractor().extract(result)


@pytest.mark.parametrize("inline", [False, True])
def test_id_is_read_from_number_pixels_then_split(inline):
    result, candidates = ambiguous_result(inline=inline)
    image = np.zeros((200, 400), dtype=np.uint8)
    observed = []

    def reread(crop, image_id, variant, cancel):
        observed.append(crop.shape)
        return OcrResult(
            [TextBlock("ID0000001", 0.21, (4, 3, 100, 20), (1, 1, 1))], image_id, variant, 108, 26
        )

    output = refine_identifier(image, result, candidates, reread, threading.Event())
    case = ClientCase()
    apply_candidates(case, output, "kg_id")
    assert observed == [(26, 108)]  # The label was not sent to the line reader.
    assert (case.data.passport_type, case.data.passport_number) == ("ID", "0000001")
    for field in ("passport_type", "passport_number"):
        assert case.evidence[field].status == FieldStatus.REVIEW
        assert "identifier-line+split-type" in case.evidence[field].candidates[0].extraction_method
    assert len([c for c in output if c.field_name == "passport_number"]) == 1


@pytest.mark.parametrize(
    "observed", ["100000001", "1D0000001", "ID0000002", "AA0000001", "ID0000001EXTRA", "ID0000001\nUNKNOWN"]
)
def test_no_guessed_letters_or_changed_numeric_suffix(observed):
    result, candidates = ambiguous_result()

    def reread(*_):
        return OcrResult(
            [TextBlock(observed, 0.99, (4, 3, 100, 20), (1, 1, 1))], "synthetic", "original", 108, 26
        )

    assert (
        refine_identifier(np.zeros((200, 400), np.uint8), result, candidates, reread, threading.Event())
        == candidates
    )


@pytest.mark.parametrize("original", ["990000001", "ID0000001", "00123456789012"])
def test_irrelevant_or_already_split_values_need_no_extra_pass(original):
    result, candidates = ambiguous_result(original)

    def never(*_):
        pytest.fail("An unrelated number triggered the ID reader")

    assert (
        refine_identifier(np.zeros((200, 400), np.uint8), result, candidates, never, threading.Event())
        == candidates
    )


def test_identifier_reader_cancellation_propagates():
    result, candidates = ambiguous_result()

    def cancelled(*_):
        raise Cancelled()

    with pytest.raises(Cancelled):
        refine_identifier(np.zeros((200, 400), np.uint8), result, candidates, cancelled, threading.Event())


def test_english_line_reader_does_not_mutate_global_ocr_settings(monkeypatch):
    settings = Settings(languages=["rus", "eng", "kir"], psm=6)
    engine = TesseractEngine(settings)
    monkeypatch.setattr(engine, "diagnose", lambda: OcrDiagnostic("synthetic.exe", "5.0", settings.languages))
    calls = []

    def read(image, **kwargs):
        calls.append((image.shape, kwargs["lang"], kwargs["config"]))
        return {
            "text": ["ID0000001"],
            "conf": [21],
            "left": [22],
            "top": [19],
            "width": [300],
            "height": [60],
            "block_num": [1],
            "par_num": [1],
            "line_num": [1],
        }

    monkeypatch.setattr("document_excel.infrastructure.ocr.pytesseract.image_to_data", read)
    output = engine.recognize_identifier_line(
        np.zeros((26, 108), np.uint8), "synthetic", "original", threading.Event()
    )
    assert calls == [((98, 344), "eng", "--oem 1 --psm 7 -c load_system_dawg=0 -c load_freq_dawg=0")]
    assert output.blocks[0].bbox == (4, 3, 100, 20)
    assert engine.settings.languages == settings.languages and engine.settings.psm == 6
