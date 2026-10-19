"""Only synthetic identifiers; type and number are separate contract fields."""

import pytest
from openpyxl import load_workbook

from document_excel.domain.models import (
    ClientCase,
    FieldEvidence,
    FieldStatus,
    RegistrationStatus,
    split_document_identifier,
)
from document_excel.infrastructure.excel import WorkbookService
from document_excel.infrastructure.extraction import FieldExtractor, apply_candidates
from document_excel.infrastructure.ocr import OcrResult, TextBlock
from document_excel.infrastructure.storage import SettingsRepository
from document_excel.presentation.main_window import MainWindow


def extract_identifier(value):
    return FieldExtractor().extract(
        OcrResult(
            [
                TextBlock("№ документа", 0.96, (40, 100, 160, 20), (1, 1, 1)),
                TextBlock(value, 0.96, (40, 130, 160, 20), (1, 1, 2)),
            ],
            "synthetic",
            "original",
            500,
            500,
        )
    )


@pytest.mark.parametrize("value", ["ID0000001", "ID 0000001", "id-0000001", "ID\t0000001"])
def test_printed_prefix_becomes_type_without_losing_leading_zeros(value):
    candidates = extract_identifier(value)
    case = ClientCase()
    apply_candidates(case, candidates, "kg_id")
    assert case.data.passport_type == "ID"
    assert case.data.passport_number == "0000001"
    for field in ("passport_type", "passport_number"):
        evidence = case.evidence[field]
        assert evidence.status != FieldStatus.CONFLICT
        assert evidence.candidates[0].source_image_id == "synthetic"
        assert evidence.candidates[0].source_bbox == (40, 130, 160, 20)
    assert case.evidence["passport_number"].candidates[0].raw_value == value
    assert "OCR мог пропустить" not in case.evidence["passport_number"].message


@pytest.mark.parametrize("value", ["100000001", "00123456789012", "ID00O0001", "AB12CD34", "ID123"])
def test_unclear_or_unstructured_identifier_is_not_guessed(value):
    assert split_document_identifier(value) is None


def test_numeric_ocr_does_not_invent_or_remove_a_type():
    case = ClientCase()
    apply_candidates(case, extract_identifier("100000001"), "kg_id")
    assert case.data.passport_type is None
    assert case.data.passport_number == "100000001"
    assert case.evidence["passport_number"].status == FieldStatus.REVIEW


def test_manual_type_and_number_survive_repeated_ocr():
    case = ClientCase()
    case.data.passport_type = "MANUAL"
    case.data.passport_number = "0000002"
    for field in ("passport_type", "passport_number"):
        case.evidence[field] = FieldEvidence(source="ручной ввод", status=FieldStatus.MANUALLY_CORRECTED)
    apply_candidates(case, extract_identifier("ID0000001"), "kg_id")
    assert case.data.passport_type == "MANUAL"
    assert case.data.passport_number == "0000002"


def test_type_and_number_reach_separate_excel_cells(template, case, profile, tmp_path):
    apply_candidates(case, extract_identifier("ID0000001"), "kg_id")
    case.registration_status = RegistrationStatus.USER_CONFIRMED
    case.confirmed_digest = case.digest()
    service = WorkbookService()
    output = service.write(
        service.plan(template, case, profile), tmp_path / "separate-fields.xlsx", case, profile
    )
    book = load_workbook(output)
    assert book["Дог."]["E19"].value == "ID"
    assert book["Дог."]["G19"].value == "0000001"
    assert book["Дог."]["G19"].data_type == "s"
    book.close()


def test_combined_manual_entry_splits_both_fields_and_undoes_together(qtbot, tmp_path, monkeypatch, case):
    monkeypatch.setattr(MainWindow, "diagnose", lambda self: None)
    window = MainWindow(SettingsRepository(tmp_path / "settings"))
    qtbot.addWidget(window)
    case.data.passport_type, case.data.passport_number = "AB", "0000002"
    window.batch.cases = [case]
    window.select_case(0)
    window.go_step(2)
    window.show()
    window.apply_field_edit("passport_number", "ID0000001")
    assert case.data.passport_type == "ID"
    assert case.data.passport_number == "0000001"
    assert window.review_form.fields["passport_type"].editor.text() == "ID"
    assert window.review_form.fields["passport_number"].editor.text() == "0000001"
    assert not case.confirmed
    window.undo_stack.undo()
    assert (case.data.passport_type, case.data.passport_number) == ("AB", "0000002")
    window.undo_stack.redo()
    assert (case.data.passport_type, case.data.passport_number) == ("ID", "0000001")
