"""Synthetic issuing authorities only; no customer document values."""

import pytest
from openpyxl import load_workbook

from document_excel.domain.models import (
    ClientCase,
    FieldEvidence,
    FieldStatus,
    RegistrationStatus,
    split_mkk_authority,
)
from document_excel.infrastructure.excel import WorkbookService, readiness
from document_excel.infrastructure.extraction import FieldExtractor, apply_candidates
from document_excel.infrastructure.ocr import OcrResult, TextBlock
from document_excel.infrastructure.storage import SettingsRepository
from document_excel.presentation.main_window import MainWindow


def extract_authority(*lines):
    return FieldExtractor().extract(
        OcrResult(
            [TextBlock(text, 0.96, (40, 40 + i * 30, 240, 20), (1, 1, i)) for i, text in enumerate(lines)],
            "synthetic-authority",
            "original",
            600,
            600,
        )
    )


@pytest.mark.parametrize(
    "value, prefix",
    [
        ("МКК 000123", "МКК"),
        ("мкк000123", "МКК"),
        ("MKK: 000123", "MKK"),
        ("МКК № 000123", "МКК"),
        ("MКК-000123", "MКК"),
    ],
)
def test_mkk_digits_fill_department_code_with_leading_zeros(value, prefix):
    candidates = extract_authority("Орган выдачи", value)
    case = ClientCase()
    apply_candidates(case, candidates, "kg_id")
    assert case.data.passport_issued_by == prefix
    assert case.data.passport_department_code == "000123"
    assert case.evidence["passport_issued_by"].candidates[0].raw_value == value
    for field in ("passport_issued_by", "passport_department_code"):
        evidence = case.evidence[field]
        assert evidence.status != FieldStatus.CONFLICT
        assert evidence.candidates[0].source_image_id == "synthetic-authority"
        assert evidence.candidates[0].source_bbox == (40, 70, 240, 20)


def test_code_can_follow_mkk_on_next_line_and_does_not_take_the_pin():
    case = ClientCase()
    apply_candidates(
        case,
        extract_authority(
            "Орган выдачи",
            "МКК",
            "000123",
            "ПИН: 00123456789012",
            "Дата выдачи: 03.04.2020",
        ),
        "kg_id",
    )
    assert case.data.passport_issued_by == "МКК"
    assert case.data.passport_department_code == "000123"
    assert case.data.personal_identifier == "00123456789012"
    assert case.data.passport_issued_date.isoformat() == "2020-04-03"


@pytest.mark.parametrize(
    "value", ["Другой орган 000123", "МКК", "МКК 00O123", "МКК 00123456789012", "МКК 000123 район"]
)
def test_other_text_or_unclear_digits_are_not_silently_split(value):
    assert split_mkk_authority(value) is None


def test_manual_issuer_and_department_code_survive_repeated_ocr():
    case = ClientCase()
    case.data.passport_issued_by, case.data.passport_department_code = "Ручной орган", "000999"
    for field in ("passport_issued_by", "passport_department_code"):
        case.evidence[field] = FieldEvidence(source="ручной ввод", status=FieldStatus.MANUALLY_CORRECTED)
    apply_candidates(case, extract_authority("Орган выдачи: МКК 000123"), "kg_id")
    assert (case.data.passport_issued_by, case.data.passport_department_code) == ("Ручной орган", "000999")


def test_explicit_different_department_code_is_a_conflict(case, profile):
    apply_candidates(
        case, extract_authority("Орган выдачи: МКК 000123", "Код подразделения: 000999"), "kg_id"
    )
    assert case.evidence["passport_department_code"].status == FieldStatus.CONFLICT
    assert readiness(case, profile, require_confirmation=False)


def test_issuing_authority_and_code_go_to_separate_excel_cells(template, case, profile, tmp_path):
    apply_candidates(case, extract_authority("Орган выдачи: МКК 000123"), "kg_id")
    case.registration_status = RegistrationStatus.USER_CONFIRMED
    case.confirmed_digest = case.digest()
    service = WorkbookService()
    output = service.write(
        service.plan(template, case, profile), tmp_path / "issuer-code.xlsx", case, profile
    )
    book = load_workbook(output)
    assert book["Дог."]["J19"].value == "МКК"
    assert book["Дог."]["K19"].value == "000123"
    assert book["Дог."]["K19"].data_type == "s"
    book.close()


def test_manual_full_issuer_splits_and_undoes_both_fields(qtbot, tmp_path, monkeypatch, case):
    monkeypatch.setattr(MainWindow, "diagnose", lambda self: None)
    window = MainWindow(SettingsRepository(tmp_path / "settings"))
    qtbot.addWidget(window)
    case.data.passport_issued_by, case.data.passport_department_code = "Ручной орган", "000999"
    window.batch.cases = [case]
    window.select_case(0)
    window.go_step(2)
    window.show()
    window.apply_field_edit("passport_issued_by", "МКК 000123")
    assert (case.data.passport_issued_by, case.data.passport_department_code) == ("МКК", "000123")
    assert window.review_form.fields["passport_department_code"].editor.text() == "000123"
    assert not case.confirmed
    window.undo_stack.undo()
    assert (case.data.passport_issued_by, case.data.passport_department_code) == ("Ручной орган", "000999")
    window.undo_stack.redo()
    assert (case.data.passport_issued_by, case.data.passport_department_code) == ("МКК", "000123")
