from __future__ import annotations

import os
import threading
from pathlib import Path

import pytest
from openpyxl import load_workbook

from document_excel.application.workflows import RecognitionService
from document_excel.domain.models import ClientCase, FieldStatus, RegistrationStatus, SourceImage
from document_excel.infrastructure.excel import WorkbookService
from document_excel.infrastructure.extraction import apply_candidates
from document_excel.infrastructure.ocr import TesseractEngine
from document_excel.infrastructure.storage import Settings


@pytest.fixture
def engine_settings():
    configured = os.environ.get("TESSERACT_CMD", "")
    local = Path(__file__).resolve().parents[1] / "tools/tesseract/tesseract.exe"
    if not configured and local.exists():
        configured = str(local)
    settings = Settings(ocr_engine="tesseract", tesseract_path=configured, languages=["eng"], preprocessing="original")
    diagnostic = TesseractEngine(settings).diagnose()
    if diagnostic.error:
        pytest.skip("Real OCR unavailable: " + diagnostic.error)
    return settings


@pytest.mark.ocr
@pytest.mark.parametrize("low_contrast, rotation", [(False, 0), (True, 0), (False, 90)])
def test_real_clean_low_contrast_and_rotated(image_factory, engine_settings, low_contrast, rotation):
    path = image_factory(low_contrast=low_contrast, rotation=rotation)
    source = SourceImage(path=path, rotation=rotation)  # PIL rotates CCW; the UI rotation corrects clockwise.
    result = RecognitionService(engine_settings).run([source], threading.Event(), lambda *_args: None)
    values = {c.field_name: c.normalized_value for c in result.candidates}
    assert values["personal_identifier"] == "00123456789012"
    assert values["birth_date"] == "1990-02-01"
    assert values["passport_type"] == "AA"
    assert values["passport_number"] == "0000001"
    assert result.elapsed > 0


@pytest.mark.ocr
def test_real_missing_registration(image_factory, engine_settings):
    path = image_factory(
        lines=[
            "KYRGYZ REPUBLIC PASSPORT",
            "FULL NAME: SYNTHETIC TEST",
            "PERSONAL NUMBER: 00123456789012",
            "PASSPORT NO: AA0000001",
        ]
    )
    output = RecognitionService(engine_settings).run(
        [SourceImage(path=path)], threading.Event(), lambda *_args: None
    )
    case = ClientCase()
    apply_candidates(case, output.candidates, output.document_type, output.label_seen)
    assert not case.data.registration_address.display()
    assert case.registration_status == RegistrationStatus.UNCERTAIN


@pytest.mark.ocr
def test_real_image_to_excel(image_factory, engine_settings, template, profile, tmp_path):
    source = SourceImage(path=image_factory())
    output = RecognitionService(engine_settings).run([source], threading.Event(), lambda *_args: None)
    case = ClientCase(images=[source], template_path=template, workbook_profile=profile.model_dump(mode="json"))
    apply_candidates(case, output.candidates, output.document_type, output.label_seen)
    assert case.data.personal_identifier == "00123456789012"
    assert not any(e.status == FieldStatus.INVALID for e in case.evidence.values())
    case.registration_status = RegistrationStatus.USER_CONFIRMED
    case.confirmed_digest = case.digest()
    service = WorkbookService()
    path = service.write(service.plan(template, case, profile), tmp_path / "contract.xlsx", case, profile)
    book = load_workbook(path)
    assert book["Дог."]["C22"].value == "00123456789012"
    book.close()


@pytest.mark.ocr
def test_real_three_installed_languages(image_factory, engine_settings):
    engine_settings.languages = ["rus", "eng", "kir"]
    diagnostic = TesseractEngine(engine_settings).diagnose()
    if diagnostic.error:
        pytest.skip(diagnostic.error)
    image = image_factory(
        lines=[
            "КЫРГЫЗ РЕСПУБЛИКАСЫ PASSPORT",
            "ФАМИЛИЯ: ТЕСТОВ",
            "ИМЯ: ПРИМЕР",
            "ПИН: 00123456789012",
            "ДАТА РОЖДЕНИЯ: 01.02.1990",
        ]
    )
    output = RecognitionService(engine_settings).run(
        [SourceImage(path=image)], threading.Event(), lambda *_args: None
    )
    values = {candidate.field_name: candidate.normalized_value for candidate in output.candidates}
    assert values["surname"] == "ТЕСТОВ"
    assert values["personal_identifier"] == "00123456789012"


@pytest.mark.ocr
def test_real_selected_region_uses_original_coordinates(image_factory, engine_settings):
    engine_settings.preprocessing = "perspective"
    image = image_factory()
    output = RecognitionService(engine_settings).run(
        [SourceImage(path=image)], threading.Event(), lambda *_args: None, roi=(50, 530, 1500, 90)
    )
    values = {candidate.field_name: candidate.normalized_value for candidate in output.candidates}
    assert values["personal_identifier"] == "00123456789012"
    assert "birth_date" not in values
    for candidate in output.candidates:
        assert candidate.source_bbox is not None
        x, y, width, height = candidate.source_bbox
        assert 50 <= x < x + width <= 1550
        assert 530 <= y < y + height <= 620


@pytest.mark.ocr
def test_real_identifier_line_recovers_letters_from_synthetic_pixels(image_factory, engine_settings):
    from dataclasses import replace

    class ConfusedPageEngine(TesseractEngine):
        def recognize_with_psm(self, *args, **kwargs):
            result = super().recognize_with_psm(*args, **kwargs)
            # Reproduce the multilingual page's 10/ID error while the dedicated
            # reader still runs real Tesseract against untouched letter pixels.
            result.blocks = [replace(b, text=b.text.replace("ID7654321", "107654321")) for b in result.blocks]
            return result

    path = image_factory(lines=["KYRGYZ REPUBLIC IDENTITY CARD", "DOCUMENT NUMBER: ID7654321"])
    output = RecognitionService(engine_settings, ConfusedPageEngine(engine_settings)).run(
        [SourceImage(path=path)], threading.Event(), lambda *_: None
    )
    case = ClientCase()
    apply_candidates(case, output.candidates, output.document_type)
    assert case.data.passport_type == "ID"
    assert case.data.passport_number == "7654321"
    assert case.evidence["passport_number"].status == FieldStatus.REVIEW
    assert "identifier-line" in case.evidence["passport_number"].candidates[0].extraction_method
