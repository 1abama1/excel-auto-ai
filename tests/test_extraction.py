import pytest

from document_excel.domain.models import (
    ClientCase,
    FieldCandidate,
    FieldEvidence,
    FieldStatus,
    RegistrationStatus,
)
from document_excel.infrastructure.extraction import (
    FieldExtractor,
    RegistrationAddressResolver,
    apply_candidates,
    classify,
    label_matches,
    load_document_profiles,
)
from document_excel.infrastructure.ocr import OcrResult, TextBlock


def result(*lines):
    return OcrResult(
        [TextBlock(text, 0.96, (20, i * 55, 900, 35), (1, 1, i)) for i, text in enumerate(lines)],
        "source",
        "original",
        1200,
        900,
    )


def candidate(value="Тестовая область, ул. Пример 1", **kwargs):
    return FieldCandidate(
        field_name="registration_address",
        raw_value=value,
        normalized_value=value,
        confidence=0.9,
        source_image_id="source",
        **kwargs,
    )


@pytest.mark.parametrize(
    "candidates, seen, absent, expected",
    [
        ([], False, False, RegistrationStatus.UNCERTAIN),
        ([], True, False, RegistrationStatus.NOT_RECOGNIZED),
        ([], False, True, RegistrationStatus.NOT_PRESENT),
        ([candidate()], True, False, RegistrationStatus.FOUND),
        ([candidate("ул. Пример")], True, False, RegistrationStatus.PARTIALLY_FOUND),
        ([candidate(), candidate("Другая область, ул. Иная 2")], True, False, RegistrationStatus.UNCERTAIN),
    ],
)
def test_registration_resolution(candidates, seen, absent, expected):
    assert (
        RegistrationAddressResolver().resolve(candidates, label_seen=seen, explicitly_absent=absent)
        == expected
    )


def test_labels_dates_pin_and_no_random_phone():
    candidates = FieldExtractor().extract(
        result(
            "FULL NAME: SYNTHETIC TEST",
            "PERSONAL NUMBER: 00123456789012",
            "DATE OF BIRTH: 01.02.1990",
            "123456789",
        )
    )
    values = {c.field_name: c.normalized_value for c in candidates}
    assert values == {
        "full_name": "SYNTHETIC TEST",
        "personal_identifier": "00123456789012",
        "birth_date": "1990-02-01",
    }


def test_spatial_field_below_label():
    candidates = FieldExtractor().extract(result("SURNAME", "SYNTHETIC", "DATE OF BIRTH: 01.02.1990"))
    surname = next(c for c in candidates if c.field_name == "surname")
    assert surname.raw_value == "SYNTHETIC"
    assert surname.extraction_method == "label-spatial"
    assert surname.source_bbox == (20, 55, 900, 35)


def test_overlapping_labels():
    assert [field for field, _ in label_matches("FULL NAME: TEST")] == ["full_name"]


def test_conflicts_and_manual_preservation():
    case = ClientCase()
    first = FieldExtractor().extract(result("PERSONAL NUMBER: 00123456789012"))
    second = FieldExtractor().extract(result("PERSONAL NUMBER: 00123456789013"))
    apply_candidates(case, first + second, "kg_id")
    assert case.evidence["personal_identifier"].status == FieldStatus.CONFLICT
    case.data.personal_identifier = "00123456789014"
    case.evidence["personal_identifier"].source = "ручной ввод"
    case.evidence["personal_identifier"].status = FieldStatus.MANUALLY_CORRECTED
    apply_candidates(case, first, "kg_id")
    assert case.data.personal_identifier == "00123456789014"


def test_invalid_candidate_never_fills_model():
    case = ClientCase()
    apply_candidates(case, FieldExtractor().extract(result("DATE OF BIRTH: 31.02.2000")), "kg_id")
    assert case.data.birth_date is None
    assert case.evidence["birth_date"].status == FieldStatus.INVALID


def test_quality_warning_lowers_field_status():
    c = FieldExtractor().extract(result("SURNAME: SYNTHETIC"), quality_warning=True)[0]
    assert c.validation_status == FieldStatus.REVIEW
    assert c.confidence < 0.82


def test_document_detection_requires_multiple_signals():
    profiles = load_document_profiles()
    assert classify(result("PASSPORT"), profiles) == "unknown"
    assert classify(result("KYRGYZ REPUBLIC PASSPORT"), profiles) == "kg_id"


def test_missing_registration_does_not_become_not_present():
    case = ClientCase()
    apply_candidates(case, [], "kg_id")
    assert case.registration_status == RegistrationStatus.UNCERTAIN


def test_manual_confirmation_does_not_allow_invalid(case, profile):
    from document_excel.infrastructure.excel import readiness

    case.evidence["birth_date"] = FieldEvidence(status=FieldStatus.INVALID)
    assert readiness(case, profile, False)


def test_document_profile_aliases_formats_and_regions():
    from document_excel.infrastructure.extraction import DocumentProfile

    profile = DocumentProfile(
        id="custom",
        name="custom",
        signatures=["CUSTOM", "DOCUMENT"],
        minimum_signatures=2,
        expected_fields=["personal_identifier"],
        regions={"personal_identifier": [0.0, 0.7, 1.0, 0.3]},
        formats={"personal_identifier": "[0-9]{14}"},
        strategy="label-spatial",
        aliases={"personal_identifier": ["CUSTOM CODE"]},
    )
    value = FieldExtractor(profile).extract(result("CUSTOM CODE: 00123456789012"))[0]
    assert value.field_name == "personal_identifier"
    assert value.validation_status == FieldStatus.REVIEW  # Outside the configured region.
    invalid = FieldExtractor(profile).extract(result("CUSTOM CODE: ABC"))[0]
    assert invalid.validation_status == FieldStatus.INVALID
