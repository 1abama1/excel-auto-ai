"""Synthetic identities only. Customer controls belong in opt-in local audits."""

import threading

import cv2
import numpy as np
import pytest

from document_excel.domain.models import ClientCase, DocumentSide, FieldCandidate, FieldEvidence, FieldStatus
from document_excel.domain.resolution import displayed_candidate
from document_excel.infrastructure.extraction import apply_candidates
from document_excel.infrastructure.full_name import FullNameExtractor, NameGroups
from document_excel.infrastructure.ocr_contract import OcrResult, TextBlock


def block(text, x, y, width=180, score=0.97):
    return TextBlock(text, score, (x, y, width, 24), (1, 1, y + x))


def front(surname="ТЕСТОВ", first="АЛЕКСАНДР", patronymic="СЕРГЕЕВИЧ"):
    blocks = [
        block("Фамилия / Surname", 400, 80),
        block(surname, 400, 115),
        block("TESTOV", 400, 150),
        block("Имя / Name", 400, 200),
        block(first, 400, 235),
        block("ALEKSANDR", 400, 270),
        block("Отчество / Patronymic", 400, 320),
        block("Пол / Sex", 400, 400),
    ]
    if patronymic:
        blocks.append(block(patronymic, 400, 355))
    return OcrResult(blocks, "synthetic-front", "physical-card", 1000, 640)


@pytest.mark.parametrize("surname", ["ЛИ", "АБДЫРАХМАНАЛИЕВ", "АБДЫРАХМАНАЛИЕВ-СУЛТАНБЕКОВ", "САРЫ КЫЗ"])
def test_names_have_no_length_limit_and_keep_double_names(surname):
    groups = FullNameExtractor().extract(front(surname), DocumentSide.FRONT)
    assert any(c.normalized_value == surname for c in groups.candidates)
    assert all(c.source_side == DocumentSide.FRONT for c in groups.candidates)
    assert groups.regions["surname"][2] > 500


@pytest.mark.parametrize("first", ["ЛИ", "АБДЫКАДЫР", "СУЛТАНБЕК-АБДЫРАХМАН", "АЙ НУР"])
def test_short_long_and_double_first_names(first):
    groups = FullNameExtractor().extract(front(first=first), DocumentSide.FRONT)
    assert any(c.field_name == "first_name" and c.normalized_value == first for c in groups.candidates)


@pytest.mark.parametrize(
    "pieces,expected",
    [
        ([("ЖУМА", 400, 60), ("БЕКОВ", 462, 90)], "ЖУМАБЕКОВ"),
        ([("САРЫ", 400, 60), ("КЫЗ", 480, 60)], "САРЫ КЫЗ"),
        ([("ТЕСТ-", 400, 70), ("СЫН", 472, 60)], "ТЕСТ-СЫН"),
    ],
)
def test_geometric_fragments_are_assembled_on_same_baseline(pieces, expected):
    result = front()
    result.blocks = [b for b in result.blocks if b.bbox[1] != 115]
    result.blocks.extend(block(t, x, 115, w) for t, x, w in pieces)
    groups = FullNameExtractor().extract(result, DocumentSide.FRONT)
    assert any(c.normalized_value == expected for c in groups.candidates)


def test_cyrillic_and_latin_remain_independent_and_mrz_never_supplies_patronymic():
    groups = FullNameExtractor().extract(front(), DocumentSide.FRONT)
    case = ClientCase()
    apply_candidates(
        case,
        groups.candidates
        + [
            FieldCandidate(
                field_name="first_name",
                raw_value="ALEKSANDR",
                normalized_value="ALEKSANDR",
                confidence=0.99,
                source_image_id="back",
                source_side=DocumentSide.BACK,
                source_method="MRZ",
            )
        ],
        "kg_physical_id",
    )
    assert case.data.first_name == "АЛЕКСАНДР"
    assert case.data.first_name_latin == "ALEKSANDR"
    assert case.data.patronymic == "СЕРГЕЕВИЧ"
    assert case.data.full_name == "ТЕСТОВ АЛЕКСАНДР СЕРГЕЕВИЧ"
    assert case.evidence["full_name"].status == FieldStatus.REVIEW
    assert not case.confirmed


def test_missing_patronymic_stays_empty_and_does_not_invent_name():
    groups = FullNameExtractor().extract(front(patronymic=None), DocumentSide.FRONT)
    case = ClientCase()
    apply_candidates(case, groups.candidates, "kg_physical_id")
    assert case.data.patronymic is None
    assert case.data.full_name == "ТЕСТОВ АЛЕКСАНДР"
    assert case.evidence["full_name"].status == FieldStatus.REVIEW


def test_missing_cyrillic_first_name_blocks_automatic_mixed_contract_name():
    groups = FullNameExtractor().extract(front(first="RH"), DocumentSide.FRONT)
    assert not any(c.raw_value == "RH" for c in groups.candidates)
    case = ClientCase()
    apply_candidates(case, groups.candidates, "kg_physical_id")
    assert case.data.first_name_cyrillic is None
    assert case.data.full_name is None
    assert case.evidence["full_name"].status == FieldStatus.NOT_RECOGNIZED


def test_printed_latin_stays_visible_but_mrz_disagreement_requires_review():
    groups = FullNameExtractor().extract(front(), DocumentSide.FRONT)
    visual = next(c for c in groups.candidates if c.field_name == "surname" and c.original_latin)
    mrz = visual.model_copy(
        update={
            "normalized_value": "OTHER",
            "raw_value": "OTHER",
            "source_image_id": "back",
            "source_side": DocumentSide.BACK,
            "source_method": "MRZ",
        }
    )
    case = ClientCase()
    apply_candidates(case, groups.candidates + [mrz], "kg_physical_id")
    assert case.data.surname_latin == visual.normalized_value
    assert case.evidence["surname_latin"].status == FieldStatus.CONFLICT
    assert not case.confirmed


@pytest.mark.parametrize("side", [DocumentSide.BACK, DocumentSide.ADDITIONAL, DocumentSide.UNKNOWN])
def test_wrong_side_does_not_produce_printed_names(side):
    assert not FullNameExtractor().extract(front(), side).candidates


def test_displayed_cyrillic_source_is_front_even_if_mrz_is_first():
    cyr = FullNameExtractor().extract(front(), DocumentSide.FRONT).candidates[0]
    mrz = cyr.model_copy(
        update={
            "normalized_value": "TESTOV",
            "source_side": DocumentSide.BACK,
            "source_method": "MRZ",
            "confidence": 0.99,
        }
    )
    evidence = FieldEvidence(candidates=[mrz, cyr])
    assert displayed_candidate(evidence, "ТЕСТОВ") is cyr
    assert displayed_candidate(evidence, "TESTOV") is mrz
    evidence.status = FieldStatus.MANUALLY_CORRECTED
    assert displayed_candidate(evidence, "ТЕСТОВ") is None


def test_gui_source_label_matches_displayed_value(qtbot):
    from document_excel.presentation.components import ReviewField

    groups = FullNameExtractor().extract(front(), DocumentSide.FRONT)
    cyr = next(c for c in groups.candidates if c.normalized_value == "АЛЕКСАНДР")
    mrz = cyr.model_copy(
        update={
            "normalized_value": "ALEKSANDR",
            "source_method": "MRZ",
            "source_side": DocumentSide.BACK,
            "confidence": 0.99,
        }
    )
    widget = ReviewField("first_name")
    qtbot.addWidget(widget)
    evidence = FieldEvidence(candidates=[mrz, cyr])
    widget.refresh("АЛЕКСАНДР", evidence, True)
    assert "лицевая" in widget.source.text() and "MRZ" not in widget.source.text()
    widget.refresh("ALEKSANDR", evidence, True)
    assert "MRZ" in widget.source.text() and "оборотная" in widget.source.text()
    evidence.status = FieldStatus.MANUALLY_CORRECTED
    widget.refresh("АЛЕКСАНДР", evidence, True)
    assert widget.source.text() == "Введено пользователем"


def test_displayed_date_provenance_normalizes_ui_format():
    candidate = FieldCandidate(
        field_name="birth_date",
        raw_value="01.02.1990",
        normalized_value="1990-02-01",
        confidence=0.95,
        source_image_id="synthetic",
    )
    assert displayed_candidate(FieldEvidence(candidates=[candidate]), "01.02.1990") is candidate


def test_partial_high_confidence_word_recovers_entire_row_only():
    pixels = np.full((100, 600, 3), 255, np.uint8)
    cv2.putText(pixels, "TESTOV", (20, 60), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 0), 2)
    partial = FieldCandidate(
        field_name="surname",
        raw_value="ТЕСТ",
        normalized_value="ТЕСТ",
        confidence=0.95,
        source_image_id="synthetic",
        source_bbox=(15, 30, 52, 35),
    )
    groups = NameGroups(
        candidates=[partial],
        regions={"surname": (10, 25, 580, 45)},
        rows={"surname": [block("ТЕСТ", 15, 30, 52)]},
    )

    class Engine:
        calls = []

        def recognize_physical_value(self, image, image_id, cancel, language, kind):
            self.calls.append(image.shape)
            return OcrResult(
                [block("ТЕСТОВ", 0, 0, 100, 0.98)], image_id, "crop", image.shape[1], image.shape[0]
            )

    engine = Engine()
    FullNameExtractor().recover(
        pixels, OcrResult([], "synthetic", "original", 600, 100), groups, engine, threading.Event()
    )
    assert [c.normalized_value for c in groups.candidates] == ["ТЕСТОВ"]
    assert len(engine.calls) == 1 and engine.calls[0][1] >= 580


def test_weak_retry_cannot_replace_a_complete_high_confidence_name():
    pixels = np.zeros((100, 600, 3), np.uint8)
    existing = FieldCandidate(
        field_name="surname",
        raw_value="ТЕСТОВ",
        normalized_value="ТЕСТОВ",
        confidence=0.98,
        source_image_id="synthetic",
        source_bbox=(10, 25, 90, 45),
    )
    groups = NameGroups(candidates=[existing], regions={"surname": (10, 25, 580, 45)})
    FullNameExtractor().recover(
        pixels, OcrResult([], "synthetic", "original", 600, 100), groups, object(), threading.Event()
    )
    assert groups.candidates == [existing]
