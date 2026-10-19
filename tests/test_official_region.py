"""Synthetic official codes and identifiers; no real client identifiers."""

import threading
from pathlib import Path

import cv2
import numpy as np
import pytest
from openpyxl import load_workbook

from document_excel.domain.errors import AppError, Cancelled
from document_excel.domain.models import (
    ClientCase,
    ClientData,
    DocumentSide,
    FieldCandidate,
    FieldEvidence,
    FieldStatus,
)
from document_excel.infrastructure.excel import WorkbookService
from document_excel.infrastructure.extraction import apply_candidates
from document_excel.infrastructure.ocr import OcrResult, TesseractEngine, TextBlock
from document_excel.infrastructure.physical_id import PlasticCard
from document_excel.infrastructure.region_ocr import DocumentNumberExtractor, FieldRegion, RegionBasedOCR
from document_excel.infrastructure.storage import Settings


def read(value, image_id="synthetic", confidence=0.9):
    return OcrResult(
        [TextBlock(value, confidence, (0, 0, 100, 20), (1, 1, 1))], image_id, "original", 200, 40
    )


@pytest.mark.parametrize("raw", ["ID1234567", "ID 1234567", "id1234567"])
def test_document_format_preserves_seven_digits(raw):
    assert DocumentNumberExtractor.literal(raw) == "ID1234567"


@pytest.mark.parametrize(
    "raw",
    ["1234567", "1D1234567", "I1234567", "D1234567", "ID123456", "ID12345678", "ID12O4567", "00123456789012"],
)
def test_missing_or_confused_symbols_do_not_invent_an_identifier(raw):
    assert not DocumentNumberExtractor.resolve([(read(raw), "original")])


@pytest.mark.parametrize("raw", ["1D1234567", "lD1234567", "I01234567", "I1234567", "D1234567"])
def test_position_normalization_requires_literal_optical_corroboration(raw):
    values = DocumentNumberExtractor.resolve([(read(raw), "gray"), (read("ID1234567"), "original")])
    assert len(values) == 2 and {v[2] for v in values} == {"ID1234567"}
    assert values[0][3]
    assert len(DocumentNumberExtractor.resolve([(read(raw), "gray"), (read("ID7654321"), "original")])) == 1


def test_authority_alternatives_preserve_cyrillic_and_mixed_scripts(monkeypatch):
    class Engine:
        def recognize_official_region(self, pixels, image_id, cancel, field, variant):
            return read(
                {
                    "original": "СЕМ 123456",
                    "gray": "СОМ 123456",
                    "clahe": "СЕM 123456",
                    "adaptive": "СЕМ 123456",
                    "upscale": "СОМ 123456",
                }[variant]
            )

    processor = RegionBasedOCR(Engine())
    monkeypatch.setattr(processor, "locate", lambda *_: [FieldRegion((20, 20, 250, 40), "label")])
    pixels = np.full((100, 400, 3), 255, np.uint8)
    values = processor.extract(
        PlasticCard(pixels, np.eye(3), (400, 100), False),
        read(""),
        threading.Event(),
        DocumentSide.BACK,
        field_name="passport_issued_by",
    )
    assert {v.normalized_value for v in values} == {"СЕМ 123456", "СОМ 123456", "СЕM 123456"}
    assert all(v.source_bbox == (20, 20, 250, 40) for v in values)
    case = ClientCase()
    apply_candidates(case, values, "kg_physical_id")
    assert case.evidence["passport_issued_by"].status == FieldStatus.CONFLICT
    case.data.passport_issued_by = "РУЧНОЙ ОРГАН"
    case.evidence["passport_issued_by"].source = "ручной ввод"
    apply_candidates(case, values, "kg_physical_id")
    assert case.data.passport_issued_by == "РУЧНОЙ ОРГАН"
    assert np.all(pixels == 255)


def test_failed_pass_does_not_discard_good_candidates_and_cancel_propagates(monkeypatch):
    class Engine:
        def recognize_official_region(self, pixels, image_id, cancel, field, variant):
            if variant == "gray":
                raise AppError("synthetic failure")
            return read("СЕМ 123456")

    processor = RegionBasedOCR(Engine())
    monkeypatch.setattr(processor, "locate", lambda *_: [FieldRegion((0, 0, 200, 40), "label")])
    card = PlasticCard(np.full((60, 300, 3), 255, np.uint8), np.eye(3), (300, 60), False)
    assert (
        len(
            processor.extract(
                card, read(""), threading.Event(), DocumentSide.BACK, field_name="passport_issued_by"
            )
        )
        == 4
    )
    cancel = threading.Event()
    cancel.set()
    with pytest.raises(Cancelled):
        processor.extract(card, read(""), cancel, DocumentSide.BACK, field_name="passport_issued_by")


def test_barcode_locates_repeat_below_actual_stripes():
    image = np.full((600, 1000, 3), 255, np.uint8)
    for x in range(650, 950, 5):
        image[130:175, x : x + 2] = 0
    region = RegionBasedOCR.barcode_region(image)
    assert region and region.locator == "barcode"
    x, y, w, h = region.bbox
    assert 630 < x < 660 and 170 < y < 185 and w > 290 and h > 40


def test_digits_baseline_locates_authority_when_label_is_unreadable():
    processor = RegionBasedOCR(None)
    result = OcrResult(
        [TextBlock("123456", 0.9, (350, 180, 100, 25), (1, 1, 1))], "synthetic", "original", 1000, 600
    )
    card = PlasticCard(np.full((600, 1000, 3), 255, np.uint8), np.eye(3), (1000, 600), False)
    region = processor.locate(card, result, DocumentSide.BACK, "passport_issued_by")[0]
    assert region.locator == "code-row" and region.bbox[1] == 177 and region.bbox[3] == 31


def test_independent_sources_increase_status_without_confirming_employee():
    candidates = []
    for side, method in [
        (DocumentSide.FRONT, "VISUAL_OCR"),
        (DocumentSide.BACK, "VISUAL_OCR"),
        (DocumentSide.BACK, "MRZ"),
    ]:
        candidates.append(
            FieldCandidate(
                field_name="passport_number",
                raw_value="ID1234567",
                normalized_value="1234567",
                confidence=0.9,
                source_image_id=side.value + method,
                source_side=side,
                source_method=method,
                extraction_method="checked+split-type",
            )
        )
    case = ClientCase(data=ClientData(passport_type="ID", personal_identifier="00123456789012"))
    apply_candidates(case, candidates, "kg_physical_id")
    assert case.evidence["passport_number"].status == FieldStatus.VALID
    assert not case.confirmed and case.data.document_number == "ID1234567"
    assert case.data.personal_identifier == "00123456789012"
    candidates[1].normalized_value = "7654321"
    apply_candidates(case, candidates, "kg_physical_id")
    assert case.evidence["passport_number"].status == FieldStatus.CONFLICT


def test_official_engine_uses_cyrillic_alphabet_without_code_dictionary(monkeypatch):
    engine = TesseractEngine(Settings(languages=["rus", "eng"], psm=6))
    configs = []

    def fake(*args, **kwargs):
        config = kwargs["extra_config"]
        path = Path(config.split('"')[1])
        text = path.read_text(encoding="utf8")
        assert "Е" in text and "О" in text and "СЕМ" not in text and "СОМ" not in text
        assert kwargs["languages"] == ["rus"]
        configs.append(path)
        return read("СЕМ 123456")

    monkeypatch.setattr(engine, "_recognize", fake)
    engine.recognize_official_region(
        np.full((40, 200, 3), 255, np.uint8), "synthetic", threading.Event(), "passport_issued_by", "original"
    )
    assert configs and not configs[0].exists() and engine.settings.psm == 6


def test_authority_number_and_pin_reach_independent_excel_cells(template, case, profile, tmp_path):
    case.data.passport_issued_by = "СЕМ 123456"
    case.data.passport_type = "ID"
    case.data.passport_number = "1234567"
    case.data.personal_identifier = "00123456789012"
    case.evidence["passport_issued_by"] = FieldEvidence(
        source="ручной ввод", status=FieldStatus.MANUALLY_CORRECTED
    )
    case.confirmed_digest = case.digest()
    service = WorkbookService()
    result = service.write(
        service.plan(template, case, profile), tmp_path / "synthetic-fields.xlsx", case, profile
    )
    book = load_workbook(result)
    assert [book["Дог."][cell].value for cell in ("J19", "E19", "G19", "C22")] == [
        "СЕМ 123456",
        "ID",
        "1234567",
        "00123456789012",
    ]
    book.close()


@pytest.mark.parametrize("damage", ["contrast", "glare", "blur"])
def test_image_variants_preserve_original_under_damage(damage):
    image = np.full((70, 300, 3), 230, np.uint8)
    cv2.putText(image, "ID1234567", (12, 45), cv2.FONT_HERSHEY_SIMPLEX, 1, (10, 10, 10), 2)
    if damage == "contrast":
        image = (image * 0.15 + 180).astype(np.uint8)
    if damage == "glare":
        image[:, 80:140] = 255
    if damage == "blur":
        image = cv2.GaussianBlur(image, (9, 9), 3)
    original = image.copy()
    variants = RegionBasedOCR.variants(image)
    assert len(variants) == 5 and variants[0][1] is image
    assert all(v.shape[:2] == image.shape[:2] for _, v in variants)
    assert np.array_equal(original, image)


@pytest.mark.parametrize("code", ["СЕМ", "СОМ"])
def test_real_ocr_reads_both_official_codes_without_substitution(code):
    from PIL import Image, ImageDraw, ImageFont

    engine = TesseractEngine(Settings())
    if engine.diagnose().error:
        pytest.skip("Local Tesseract unavailable")
    font_path = next(
        (
            p
            for p in [
                Path("C:/Windows/Fonts/arial.ttf"),
                Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
            ]
            if p.exists()
        ),
        None,
    )
    if font_path is None:
        pytest.skip("No synthetic test font")
    image = Image.new("RGB", (520, 90), "white")
    ImageDraw.Draw(image).text(
        (15, 12), code + " 123456", font=ImageFont.truetype(str(font_path), 44), fill="black"
    )
    values = RegionBasedOCR(engine).extract(
        PlasticCard(np.asarray(image).copy(), np.eye(3), (520, 90), False),
        read(""),
        threading.Event(),
        DocumentSide.BACK,
        field_name="passport_issued_by",
        bbox=(0, 0, 520, 90),
    )
    assert any(v.normalized_value == code + " 123456" for v in values)
    assert all(v.validation_status == FieldStatus.REVIEW for v in values)
