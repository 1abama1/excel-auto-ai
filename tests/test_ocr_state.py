"""Synthetic race, provenance and replacement regressions."""

import pytest

from document_excel.application.ocr_state import OcrState, event
from document_excel.application.workflows import RecognitionOutput
from document_excel.domain.models import (
    ClientCase,
    DocumentSide,
    FieldCandidate,
    FieldEvidence,
    FieldStatus,
    SourceImage,
)
from document_excel.infrastructure.extraction import apply_candidates


def candidate(field, value, image_id="front", **kwargs):
    return FieldCandidate(
        field_name=field,
        raw_value=value,
        normalized_value=value,
        confidence=0.9,
        source_image_id=image_id,
        source_side=DocumentSide.FRONT,
        **kwargs,
    )


def output(*values):
    return RecognitionOutput(list(values), "kg_physical_id", [], False, 0)


@pytest.fixture
def tracked_case(tmp_path):
    p = tmp_path / "synthetic.png"
    p.write_bytes(b"synthetic image revision one")
    return ClientCase(images=[SourceImage(id="front", path=p, role=DocumentSide.FRONT)])


def test_new_run_removes_disappeared_field_and_old_candidates(tracked_case):
    state = OcrState()
    case = tracked_case
    first = state.begin(case)
    assert state.commit(
        case, first.id, output(candidate("surname", "ПРИМЕР"), candidate("first_name", "ТЕСТ"))
    )
    assert case.data.full_name == "ПРИМЕР ТЕСТ"
    second = state.begin(case)
    assert state.commit(case, second.id, output(candidate("surname", "НОВЫЙ")))
    assert case.data.surname == "НОВЫЙ" and case.data.first_name is None and case.data.full_name is None
    assert case.evidence["first_name"].status == FieldStatus.NOT_RECOGNIZED
    assert all(c.run_id == second.id for e in case.evidence.values() for c in e.candidates)


def test_old_run_cannot_overwrite_newer_result(tracked_case):
    state = OcrState()
    case = tracked_case
    old = state.begin(case)
    new = state.begin(case)
    assert state.commit(case, new.id, output(candidate("surname", "НОВЫЙ")))
    assert not state.commit(case, old.id, output(candidate("surname", "СТАРЫЙ")))
    assert case.data.surname == "НОВЫЙ"
    assert case.diagnostic_events[-1]["event"] == "STALE_RESULT_IGNORED"


@pytest.mark.parametrize("change", ["bytes", "rotation", "remove", "type", "role"])
def test_input_revision_change_rejects_late_result(tracked_case, change):
    state = OcrState()
    case = tracked_case
    run = state.begin(case)
    if change == "bytes":
        case.images[0].path.write_bytes(b"different bytes same path")
    if change == "rotation":
        case.images[0].rotation = 90
    if change == "remove":
        case.images.clear()
    if change == "type":
        case.document_type_manual = True
    if change == "role":
        case.images[0].role = DocumentSide.BACK
    assert not state.commit(case, run.id, output(candidate("surname", "СТАРЫЙ")))
    assert not case.data.surname


def test_manual_override_survives_replacement_and_records_disagreement(tracked_case):
    state = OcrState()
    case = tracked_case
    case.data.surname = "РУЧНОЙ"
    case.evidence["surname"] = FieldEvidence(source="ручной ввод", status=FieldStatus.MANUALLY_CORRECTED)
    run = state.begin(case)
    state.commit(case, run.id, output(candidate("surname", "НОВЫЙ")))
    assert case.data.surname == "РУЧНОЙ" and case.evidence["surname"].manual_conflict


def test_cancelled_and_foreign_run_never_apply(tracked_case):
    state = OcrState()
    case = tracked_case
    run = state.begin(case)
    other = ClientCase(images=[SourceImage(id="other", path=case.images[0].path)])
    assert not state.commit(other, run.id, output(candidate("surname", "ЧУЖОЙ")))
    state.stop(case, run.id, "CANCELLED")
    assert not state.commit(case, run.id, output(candidate("surname", "СТАРЫЙ")))
    assert not case.data.surname and not other.data.surname


def test_missing_identity_is_quarantined_until_explicit_confirmation(tracked_case, tmp_path):
    case = tracked_case
    state = OcrState()
    p = tmp_path / "back.png"
    p.write_bytes(b"another synthetic side")
    case.images.append(SourceImage(id="back", path=p, role=DocumentSide.BACK))
    run = state.begin(case)
    values = output(
        candidate("surname", "ПРИМЕР"),
        candidate("passport_number", "1234567"),
        candidate("passport_number", "7654321", image_id="back"),
    )
    assert state.commit(case, run.id, values)
    assert case.identity_status == "review" and case.pending_identity and case.data.surname is None
    assert state.accept_identity(case)
    assert case.identity_status == "user_confirmed" and case.data.surname == "ПРИМЕР"
    assert case.evidence["passport_number"].status == FieldStatus.CONFLICT


def test_same_document_number_links_sides_without_counting_repeated_crop_votes(tracked_case, tmp_path):
    case = tracked_case
    state = OcrState()
    p = tmp_path / "back.png"
    p.write_bytes(b"synthetic back")
    case.images.append(SourceImage(id="back", path=p, role=DocumentSide.BACK))
    run = state.begin(case)
    assert state.commit(
        case,
        run.id,
        output(
            candidate("passport_number", "1234567"), candidate("passport_number", "1234567", image_id="back")
        ),
    )
    assert case.identity_status == "matched"
    assert case.evidence["passport_number"].status == FieldStatus.REVIEW


def test_atomic_resolution_failure_does_not_publish_partial_data(tracked_case, monkeypatch):
    from document_excel.domain.resolution import FieldResolver

    state = OcrState()
    case = tracked_case
    run = state.begin(case)

    def fail(self, working, *args):
        working.data.surname = "ЧАСТИЧНЫЙ"
        raise ValueError("synthetic failure")

    monkeypatch.setattr(FieldResolver, "resolve", fail)
    with pytest.raises(ValueError):
        state.commit(case, run.id, output(candidate("surname", "НОВЫЙ")))
    assert not case.data.surname and case.ocr_runs[run.id].status == "RUNNING"


def test_diagnostic_events_drop_values_and_paths(tracked_case):
    event(
        tracked_case,
        "FIELD_CANDIDATE_FOUND",
        field="surname",
        raw_value="PRIVATE",
        path="PRIVATE",
        text="PRIVATE",
    )
    assert "PRIVATE" not in str(tracked_case.diagnostic_events)


def test_removed_candidate_does_not_erase_unchanged_other_photo_source(tracked_case, tmp_path):
    case = tracked_case
    p = tmp_path / "support.png"
    p.write_bytes(b"support")
    case.images.append(SourceImage(id="support", path=p, role=DocumentSide.ADDITIONAL))
    apply_candidates(
        case,
        [candidate("surname", "ПРИМЕР"), candidate("first_name", "ТЕСТ", image_id="support")],
        "kg_physical_id",
    )
    state = OcrState()
    run = state.begin(case, [case.images[0]])
    state.commit(case, run.id, output(candidate("surname", "НОВЫЙ")))
    assert case.data.first_name == "ТЕСТ"


@pytest.mark.parametrize(
    "field,value",
    [("gender", "a"), ("gender", "№"), ("citizenship", "№"), ("first_name", "12"), ("surname", "_")],
)
def test_semantically_invalid_tokens_never_become_resolved_fields(field, value):
    case = ClientCase()
    apply_candidates(case, [candidate(field, value)], "kg_physical_id")
    assert getattr(case.data, field) is None
    assert case.evidence[field].status == FieldStatus.INVALID


def test_invalid_high_confidence_does_not_displace_valid_value():
    case = ClientCase()
    bad = candidate("gender", "a")
    bad.confidence = 0.99
    good = candidate("gender", "М")
    good.confidence = 0.4
    apply_candidates(case, [bad, good], "kg_physical_id")
    assert case.data.gender == "MALE" and case.evidence["gender"].status == FieldStatus.REVIEW


def test_wrong_name_region_cannot_repeat_patronymic_as_first_name():
    case = ClientCase()
    case.field_regions["first_name"] = {"image_id": "front", "bbox": (300, 200, 350, 60)}
    apply_candidates(
        case,
        [
            candidate("surname", "ПРИМЕР", source_bbox=(300, 100, 350, 60)),
            candidate("first_name", "ТЕСТОВИЧ", source_bbox=(300, 350, 350, 40)),
            candidate("patronymic", "ТЕСТОВИЧ", source_bbox=(300, 350, 350, 40)),
        ],
        "kg_physical_id",
    )
    assert case.data.first_name is None and case.data.full_name is None


@pytest.mark.parametrize(
    "raw,expected",
    [("М", "MALE"), ("Э / М", "MALE"), ("Ж", "FEMALE"), ("А / F", "FEMALE"), ("unknown", "UNKNOWN")],
)
def test_gender_has_canonical_values(raw, expected):
    case = ClientCase()
    apply_candidates(case, [candidate("gender", raw)], "kg_physical_id")
    assert case.data.gender == expected


@pytest.mark.parametrize("raw", ["03 04 2010", "03.04.2010", "03/04/2010"])
def test_date_parsing_keeps_birth_date_without_age_policy(raw):
    from document_excel.infrastructure.field_extractors import DateExtractor

    assert DateExtractor().value("birth_date", raw) == "2010-04-03"


def test_back_issue_date_does_not_fill_front_birth_date():
    c = candidate("birth_date", "2000-01-02")
    c.source_side = DocumentSide.BACK
    case = ClientCase()
    apply_candidates(case, [c], "kg_physical_id")
    assert case.data.birth_date is None
    assert case.evidence["birth_date"].status == FieldStatus.INVALID


def test_field_reread_rebuilds_full_name_from_current_components(tracked_case):
    case = tracked_case
    state = OcrState()
    run = state.begin(case)
    state.commit(case, run.id, output(candidate("surname", "ПРИМЕР"), candidate("first_name", "ТЕСТ")))
    run = state.begin(case, fields={"first_name"}, scope="field")
    state.commit(case, run.id, output(candidate("first_name", "НОВЫЙ")))
    assert case.data.full_name == "ПРИМЕР НОВЫЙ"


def test_legacy_manual_address_survives_new_optical_candidate():
    from document_excel.domain.models import Address, RegistrationStatus

    case = ClientCase(registration_status=RegistrationStatus.MANUALLY_ENTERED)
    case.data.registration_address = Address(raw="Тестовая область, ул. Пример, 1")
    apply_candidates(case, [candidate("registration_address", "Другая область, ул. Пример, 2")], "kg_id")
    assert case.data.registration_address.raw == "Тестовая область, ул. Пример, 1"
    assert case.registration_status == RegistrationStatus.MANUALLY_ENTERED
    assert case.evidence["registration_address"].manual_conflict


def test_composed_cyrillic_name_cites_used_front_components_only():
    case = ClientCase()
    latin = candidate("surname", "PRIMER", source_method="MRZ")
    latin.source_side = DocumentSide.BACK
    latin.confidence = 0.99
    apply_candidates(
        case,
        [latin, candidate("surname", "ПРИМЕР"), candidate("first_name", "ТЕСТ")],
        "kg_physical_id",
    )
    assert case.data.full_name == "ПРИМЕР ТЕСТ"
    assert case.data.surname_latin == "PRIMER"
    assert {c.source_side for c in case.evidence["full_name"].candidates} == {DocumentSide.FRONT}
