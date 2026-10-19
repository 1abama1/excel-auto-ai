from document_excel.application.workflows import merge_regional_candidates
from document_excel.domain.models import ClientCase, FieldCandidate, FieldStatus
from document_excel.infrastructure.extraction import apply_candidates


def candidate(name, value, variant):
    return FieldCandidate(
        field_name=name,
        raw_value=value,
        normalized_value=value,
        source_image_id="synthetic",
        confidence=0.95,
        variant=variant,
    )


def test_latin_region_does_not_erase_independent_cyrillic_name():
    primary = [candidate("surname", "ТЕСТОВ", "global"), candidate("first_name", "ПРИМЕР", "global")]
    region = [candidate("surname", "TESTOV", "crop"), candidate("first_name", "PRIMER", "crop")]
    merged = merge_regional_candidates(primary, region)
    case = ClientCase()
    apply_candidates(case, merged, "kg_physical_id")
    assert case.data.full_name == "ТЕСТОВ ПРИМЕР"
    assert case.data.first_name_cyrillic == "ПРИМЕР" and case.data.first_name_latin == "PRIMER"


def test_region_disagreement_in_same_script_is_visible_conflict():
    primary = [candidate("surname", "ТЕСТОВ", "global")]
    region = [candidate("surname", "ТЕСТЕВ", "crop")]
    case = ClientCase()
    apply_candidates(case, merge_regional_candidates(primary, region), "kg_physical_id")
    assert case.evidence["surname_cyrillic"].status == FieldStatus.CONFLICT
    assert case.data.surname_cyrillic is None


def test_homoglyphs_are_disambiguated_only_in_explicit_cyrillic_name_row():
    from document_excel.domain.models import DocumentSide
    from document_excel.infrastructure.physical_id import normalize_physical_name_scripts
    top = candidate("first_name", "TECT", "global")
    top.source_bbox = (400, 390, 100, 20)
    latin = candidate("first_name", "TECT", "global")
    latin.source_bbox = (400, 440, 100, 20)
    normalize_physical_name_scripts([top, latin], 1000, 1000, DocumentSide.FRONT)
    assert top.normalized_value == "ТЕСТ" and top.raw_value == "TECT"
    assert top.validation_status == FieldStatus.REVIEW
    assert latin.normalized_value == "TECT"


def test_cyrillic_crop_does_not_invent_non_equivalent_letters():
    from document_excel.infrastructure.field_extractors import FirstNameExtractor
    extractor = FirstNameExtractor()
    assert extractor.value("ТEСТ", "cyrillic") == "ТЕСТ"
    assert extractor.value("TECT", "cyrillic") == "ТЕСТ"
    assert extractor.value("TEST", "cyrillic") is None
    assert extractor.value("Y", "cyrillic") is None
