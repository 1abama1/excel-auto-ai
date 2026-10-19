"""Synthetic Kyrgyz identity layouts; no customer documents or identity details."""

from datetime import date
from pathlib import Path

import pytest

from document_excel.domain.models import (
    ClientCase,
    DocumentSide,
    FieldCandidate,
    FieldEvidence,
    FieldStatus,
    RegistrationStatus,
    SourceImage,
    validate_field,
)
from document_excel.infrastructure.extraction import (
    DocumentSideClassifier,
    FieldExtractor,
    apply_candidates,
    load_document_profiles,
)
from document_excel.infrastructure.ocr import OcrResult, TextBlock


def synthetic_result(*lines: str, image_id: str = "synthetic", variant: str = "original") -> OcrResult:
    return OcrResult(
        [TextBlock(text, 0.96, (40, 40 + i * 36, 900, 24), (1, 1, i)) for i, text in enumerate(lines)],
        image_id,
        variant,
        1200,
        1200,
    )


def front() -> OcrResult:
    return synthetic_result(
        "КЫРГЫЗ РЕСПУБЛИКАСЫ / IDENTITY CARD",
        "Фамилиясы / Фамилия / Surname",
        "СИНТЕТИЧЕСКИЙ / SYNTHETIC",
        "Аты / Имя / Name",
        "ТЕСТ / TEST",
        "Атасынын аты / Отчество / Patronymic",
        "ПРИМЕРОВИЧ",
        "Туулган күнү / Дата рождения / Date of birth",
        "01.02.1990",
        image_id="front",
    )


def back() -> OcrResult:
    return synthetic_result(
        "Берген мекеме / Орган выдачи / Authority",
        "ТЕСТОВЫЙ ОРГАН 000",
        "Берилген күнү / Дата выдачи / Date of issue",
        "03.04.2020",
        "Жеке номеру / Персональный номер",
        "00123456789012",
        "Документтик № / № Документа",
        "ID0000001",
        "Колдонуу мөөнөтү / Срок действия",
        "03.04.2030",
        "Катталган жери / Адрес прописки",
        "Кыргызская Республика,",
        "Тестовая область,",
        "г. Пример, Тестовый район,",
        "ул. Образец, д. 1,",
        "кв. 2",
        "QR / Поделиться",
        image_id="back",
    )


def extractor() -> FieldExtractor:
    return FieldExtractor(next(p for p in load_document_profiles() if p.id == "kg_id"))


def test_front_multilingual_name_forms_are_preserved_without_false_conflict():
    values = extractor().extract(front())
    case = ClientCase()
    apply_candidates(case, values, "kg_id")
    assert case.data.surname == "СИНТЕТИЧЕСКИЙ"
    assert case.data.first_name == "ТЕСТ"
    assert case.data.birth_date == date(1990, 2, 1)
    assert case.data.full_name == "СИНТЕТИЧЕСКИЙ ТЕСТ ПРИМЕРОВИЧ"
    evidence = case.evidence["first_name"]
    assert evidence.original_cyrillic == "ТЕСТ"
    assert evidence.original_latin == "TEST"
    assert evidence.status != FieldStatus.CONFLICT
    assert len(evidence.candidates) == 2
    assert {c.source_side for c in values} == {DocumentSide.FRONT}


def test_back_distinguishes_pin_document_number_and_each_date():
    case = ClientCase()
    apply_candidates(case, extractor().extract(back()), "kg_id", label_seen=True)
    assert case.data.personal_identifier == "00123456789012"
    assert case.data.passport_type == "ID"
    assert case.data.passport_number == "0000001"
    assert case.data.passport_issued_date == date(2020, 4, 3)
    assert case.data.expiration_date == date(2030, 4, 3)
    assert case.data.birth_date is None
    assert case.data.passport_issued_by == "ТЕСТОВЫЙ ОРГАН 000"
    address = case.data.registration_address.display()
    assert all(
        fragment in address for fragment in ("Кыргызская", "Тестовая область", "район", "д. 1", "кв. 2")
    )
    assert "QR" not in address
    assert "Поделиться" not in address
    assert case.registration_status == RegistrationStatus.FOUND
    assert {c.source_side for c in case.evidence["passport_number"].candidates} == {DocumentSide.BACK}


def test_two_sides_form_one_case_with_distinct_sources():
    case = ClientCase(
        images=[
            SourceImage(id="front", path=Path("front.png")),
            SourceImage(id="back", path=Path("back.png")),
        ]
    )
    apply_candidates(case, extractor().extract(front()), "kg_id")
    assert case.registration_status == RegistrationStatus.UNCERTAIN
    apply_candidates(case, extractor().extract(back()), "kg_id", label_seen=True)
    assert case.data.birth_date == date(1990, 2, 1)
    assert case.data.passport_issued_date == date(2020, 4, 3)
    assert [image.role for image in case.images] == [DocumentSide.FRONT, DocumentSide.BACK]
    assert case.evidence["birth_date"].candidates[0].source_image_id == "front"
    assert case.evidence["registration_address"].candidates[0].source_image_id == "back"


@pytest.mark.parametrize(
    "text", ["ПАСПОРТ", "Date of birth: 01.02.1990", "Document no: ID0000001", "QR 00123456789012"]
)
def test_side_detection_does_not_guess_from_one_field(text: str):
    assert DocumentSideClassifier().classify(synthetic_result(text)) == DocumentSide.UNKNOWN


def test_side_detection_ambiguous_equal_signals_remains_unknown():
    result = synthetic_result(
        "Имя: ТЕСТ", "Дата рождения: 01.02.1990", "ПИН: 00123456789012", "Дата выдачи: 03.04.2020"
    )
    assert DocumentSideClassifier().classify(result) == DocumentSide.UNKNOWN


def test_same_script_disagreement_is_still_a_conflict():
    case = ClientCase()
    values = extractor().extract(synthetic_result("Имя: ТЕСТ / TEST", "Имя: ИНОЙ / TEST"))
    apply_candidates(case, values, "kg_id")
    assert case.evidence["first_name"].status == FieldStatus.CONFLICT


def test_adjacent_cyrillic_and_latin_value_lines_are_representations():
    case = ClientCase()
    apply_candidates(case, extractor().extract(synthetic_result("Аты / Имя / Name", "ТЕСТ", "TEST")), "kg_id")
    assert case.data.first_name == "ТЕСТ"
    assert case.evidence["first_name"].original_latin == "TEST"
    assert case.evidence["first_name"].status != FieldStatus.CONFLICT


def test_latin_representation_can_be_selected_and_is_not_overwritten_by_new_ocr():
    case = ClientCase()
    values = extractor().extract(front())
    apply_candidates(case, values, "kg_id")
    case.data.first_name = "TEST"
    case.evidence["first_name"].source = "Введено пользователем"
    case.evidence["first_name"].status = FieldStatus.MANUALLY_CORRECTED
    apply_candidates(case, values, "kg_id")
    assert case.data.first_name == "TEST"
    assert case.evidence["first_name"].original_cyrillic == "ТЕСТ"


def test_manual_registration_survives_ocr_of_back():
    case = ClientCase()
    case.data.registration_address.raw = "Введённый адрес"
    case.evidence["registration_address"] = FieldEvidence(status=FieldStatus.MANUALLY_CORRECTED)
    case.registration_status = RegistrationStatus.MANUALLY_ENTERED
    apply_candidates(case, extractor().extract(back()), "kg_id", label_seen=True)
    assert case.data.registration_address.raw == "Введённый адрес"
    assert case.registration_status == RegistrationStatus.MANUALLY_ENTERED


def test_empty_pin_label_does_not_steal_value_of_the_next_label():
    values = extractor().extract(synthetic_result("Жеке номеру", "Документтик №", "ID0000001"))
    assert not any(c.field_name == "personal_identifier" for c in values)
    assert next(c for c in values if c.field_name == "passport_number").normalized_value == "0000001"
    assert next(c for c in values if c.field_name == "passport_type").normalized_value == "ID"


def test_pin_cannot_be_used_as_a_document_number():
    case = ClientCase()
    apply_candidates(case, extractor().extract(synthetic_result("№ Документа: 00123456789012")), "kg_id")
    assert case.data.passport_number is None
    assert case.evidence["passport_number"].status == FieldStatus.INVALID
    assert validate_field("passport_number", "00123456789012") is not None


def test_partly_occluded_invalid_date_is_not_repaired_or_used():
    case = ClientCase()
    apply_candidates(case, extractor().extract(synthetic_result("Дата рождения: 0?.02.1990")), "kg_id")
    assert case.data.birth_date is None
    assert case.evidence["birth_date"].status == FieldStatus.INVALID


def test_source_metadata_defaults_are_backward_compatible_and_manual_role_wins():
    candidate = FieldCandidate(
        field_name="first_name",
        raw_value="ТЕСТ",
        normalized_value="ТЕСТ",
        confidence=0.9,
        source_image_id="front",
    )
    assert candidate.source_side == DocumentSide.UNKNOWN
    image = SourceImage(id="front", path=Path("front.png"), role=DocumentSide.ADDITIONAL)
    case = ClientCase(images=[image])
    apply_candidates(case, extractor().extract(front()), "kg_id")
    assert image.role == DocumentSide.ADDITIONAL


def compact_number_layout(translation: str, value: str) -> OcrResult:
    """Small bilingual card labels on separate OCR lines, with a 13 px value row."""
    rows = [
        ("Жеке номери / Персональный номер", (35, 226, 195, 13)),
        ("00123456789012", (35, 245, 130, 13)),
        ("Документтин № / № документа", (35, 267, 195, 13)),
        (translation, (35, 286, 86, 11)),
        (value, (36, 308, 81, 13)),
        ("Действителен до / Срок действия", (35, 332, 216, 13)),
        ("03.04.2030", (35, 352, 85, 13)),
    ]
    return OcrResult(
        [TextBlock(text, 0.94, bbox, (1, 1, index)) for index, (text, bbox) in enumerate(rows)],
        "compact-synthetic-back",
        "original",
        490,
        514,
    )


@pytest.mark.parametrize("translation", ["/ DocumentNo", "/ Documeni No", "/ LLLLLLLL X"])
def test_number_row_wins_over_a_closer_translation_label(translation: str):
    candidates = extractor().extract(compact_number_layout(translation, "AB0000001"))
    values = [c for c in candidates if c.field_name == "passport_number"]
    assert values
    assert {c.normalized_value for c in values} == {"0000001"}
    assert all(c.source_bbox == (36, 308, 81, 13) for c in values)


def test_small_numeric_ocr_document_number_stays_review_without_invented_prefix():
    case = ClientCase()
    candidates = extractor().extract(compact_number_layout("/ Documeni No", "100000001."))
    apply_candidates(case, candidates, "kg_id")
    assert case.data.passport_number == "100000001"
    assert case.evidence["passport_number"].status == FieldStatus.REVIEW
    assert "серии" in case.evidence["passport_number"].message
    assert case.data.personal_identifier == "00123456789012"
    assert case.data.expiration_date == date(2030, 4, 3)


def test_garbled_alphabetic_label_cannot_fill_document_number_when_value_is_missing():
    case = ClientCase()
    candidates = extractor().extract(
        synthetic_result("№ документа", "/ LLLLLLLL X", "Срок действия: 03.04.2030")
    )
    apply_candidates(case, candidates, "kg_id")
    assert case.data.passport_number is None
    assert case.evidence["passport_number"].status == FieldStatus.INVALID


def test_inline_translation_fragment_yields_to_a_typed_value_below():
    candidates = extractor().extract(synthetic_result("№ документа / Documeni No", "AB0000001"))
    assert {c.normalized_value for c in candidates if c.field_name == "passport_number"} == {"0000001"}


def test_same_field_translated_label_can_precede_date_value():
    candidates = extractor().extract(synthetic_result("Дата выдачи", "Dateofissue", "03.04.2020"))
    assert {c.normalized_value for c in candidates if c.field_name == "passport_issued_date"} == {
        "2020-04-03"
    }


def test_corrupted_final_language_label_does_not_replace_bilingual_name_below():
    result = OcrResult(
        [
            TextBlock("Аты / Имя / Мате", 0.94, (190, 104, 110, 17), (1, 1, 1)),
            TextBlock("ПРИМЕР / PRIMER", 0.94, (191, 126, 151, 15), (1, 1, 2)),
        ],
        "synthetic-front",
        "original",
        490,
        514,
    )
    case = ClientCase()
    apply_candidates(case, extractor().extract(result), "kg_id")
    assert case.data.first_name == "ПРИМЕР"
    assert case.evidence["first_name"].original_latin == "PRIMER"
    assert case.evidence["first_name"].status != FieldStatus.CONFLICT


def test_partly_corrupted_english_label_before_colon_is_not_part_of_name():
    case = ClientCase()
    result = synthetic_result("ИМЯ / GIVEN МАМЕ: ПРИМЕР / PRIMER")
    apply_candidates(case, extractor().extract(result), "kg_id")
    assert case.data.first_name == "ПРИМЕР"
    assert {c.normalized_value for c in case.evidence["first_name"].candidates} == {"ПРИМЕР", "PRIMER"}


def test_intermediate_given_label_fragment_is_not_a_name_candidate():
    result = synthetic_result("ИМЯ / GIVEN / NAME: ПРИМЕР / PRIMER")
    values = extractor().extract(result)
    assert {c.normalized_value for c in values if c.field_name == "first_name"} == {"ПРИМЕР", "PRIMER"}


def test_real_inline_name_is_retained_when_the_following_line_is_its_latin_form():
    case = ClientCase()
    apply_candidates(case, extractor().extract(synthetic_result("Аты / Имя / ПРИМЕР", "PRIMER")), "kg_id")
    assert case.data.first_name == "ПРИМЕР"
    assert case.evidence["first_name"].original_latin == "PRIMER"


def test_ethnicity_is_not_a_second_citizenship_on_bilingual_identity():
    result = synthetic_result(
        "Жарандыгы / Гражданство / Citizenship",
        "ТЕСТОВАЯ РЕСПУБЛИКА",
        "Улуту / Национальность / Nationality",
        "ТЕСТОВАЯ НАЦИОНАЛЬНОСТЬ",
    )
    case = ClientCase()
    apply_candidates(case, extractor().extract(result), "kg_id")
    assert case.data.citizenship == "ТЕСТОВАЯ РЕСПУБЛИКА"
    assert len(case.evidence["citizenship"].candidates) == 1
    assert case.evidence["citizenship"].status != FieldStatus.CONFLICT


def test_generic_passport_nationality_label_still_means_citizenship():
    case = ClientCase()
    apply_candidates(
        case, FieldExtractor().extract(synthetic_result("Nationality: SYNTHETIC REPUBLIC")), "unknown"
    )
    assert case.data.citizenship == "SYNTHETIC REPUBLIC"


def test_citizenship_value_uses_its_word_column_when_ocr_joins_signature_on_left():
    result = OcrResult(
        [
            TextBlock("Жарандыгы / Гражданство / Citizenship", 0.94, (190, 285, 220, 17), (1, 1, 1)),
            TextBlock("Подпись / Signature", 0.94, (40, 310, 123, 17), (1, 1, 2)),
            TextBlock("ТЕСТОВАЯ", 0.94, (191, 310, 75, 17), (1, 1, 2)),
            TextBlock("РЕСПУБЛИКА", 0.94, (275, 310, 110, 17), (1, 1, 2)),
            TextBlock("x", 0.94, (470, 334, 9, 8), (1, 1, 2)),
        ],
        "synthetic-front",
        "original",
        490,
        514,
    )
    case = ClientCase()
    apply_candidates(case, extractor().extract(result), "kg_id")
    assert case.data.citizenship == "ТЕСТОВАЯ РЕСПУБЛИКА"
    assert case.evidence["citizenship"].candidates[0].source_bbox == (191, 310, 194, 17)
