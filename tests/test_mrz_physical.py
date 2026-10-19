"""Synthetic TD1 and plastic layouts; never customer MRZ or personal data."""

import threading
from datetime import date
from pathlib import Path

import numpy as np
import pytest

from document_excel.domain.models import (
    Address,
    ClientCase,
    DocumentSide,
    FieldCandidate,
    FieldEvidence,
    FieldStatus,
    RegistrationStatus,
    SourceImage,
    parse_date,
)
from document_excel.infrastructure.document_profiles import KyrgyzPhysicalIDProfile
from document_excel.infrastructure.extraction import (
    DocumentSideClassifier,
    FieldExtractor,
    apply_candidates,
    classify,
    load_document_profiles,
)
from document_excel.infrastructure.mrz import TD1Parser, check_digit, find_mrz_lines, mrz_candidates
from document_excel.infrastructure.ocr import OcrResult, TextBlock
from document_excel.infrastructure.physical_id import prepare_plastic_card, prepare_region


def td1(number="ID0000001", pin="00123456789012", birth="901107", expiry="341122"):
    a = "I<KGZ" + number + check_digit(number) + pin + "<"
    b = birth + check_digit(birth) + "M" + expiry + check_digit(expiry) + "KGZ" + "<" * 11
    b += check_digit(a[5:] + b[:7] + b[8:15] + b[18:29])
    return [a, b, "SINTETIKOV<<TEST".ljust(30, "<")]


def test_neural_mrz_stops_after_three_valid_original_lines(monkeypatch):
    from document_excel.infrastructure.mrz import MRZRecognizer

    lines = iter(td1())

    class Engine:
        preferred_input = "color"
        calls = 0

        def recognize_mrz_line(self, image, image_id, cancel):
            self.calls += 1
            return OcrResult(
                [TextBlock(next(lines), 0.97, (0, 0, 200, 20), (1, 1, 1))], image_id, "mrz", 200, 20
            )

    engine = Engine()
    recognizer = MRZRecognizer(engine)
    monkeypatch.setattr(
        recognizer, "line_boxes", lambda _: [(0, 200, 200, 20), (0, 230, 200, 20), (0, 260, 200, 20)]
    )
    monkeypatch.setattr(
        recognizer, "recognize_glyph_spans", lambda *_: pytest.fail("Valid originals need no glyph retries")
    )
    report = recognizer.recognize(np.full((300, 200, 3), 255, np.uint8), "synthetic-back", threading.Event())
    assert report.valid and engine.calls == 3


def test_mrz_filler_separator_does_not_erase_printed_hyphenated_name():
    lines = td1()
    lines[2] = "SINTETIKOV<<ANNA<MARIA".ljust(30, "<")
    report = TD1Parser().parse(lines, "synthetic-back")
    assert report.valid
    printed = FieldCandidate(
        field_name="first_name",
        raw_value="ANNA-MARIA",
        normalized_value="ANNA-MARIA",
        confidence=0.98,
        source_image_id="synthetic-front",
        source_side=DocumentSide.FRONT,
    )
    values = mrz_candidates(report, [printed])
    assert not any(c.field_name == "first_name" for c in values)


def ocr(*lines, image_id="front"):
    return OcrResult(
        [TextBlock(s, 0.96, (400, 40 + i * 45, 600, 28), (1, 1, i)) for i, s in enumerate(lines)],
        image_id,
        "original",
        1600,
        1000,
    )


def front():
    return ocr(
        "КЫРГЫЗ РЕСПУБЛИКАСЫ IDENTITY CARD",
        "Фамилия / Surname: СИНТЕТИКОВ / SINTETIKOV",
        "Имя / Name: ТЕСТ / TEST",
        "Отчество / Patronymic: ПРИМЕРОВИЧ",
        "Пол / Sex: Э / М",
        "Гражданство / Nationality: КЫРГЫЗ РЕСПУБЛИКАСЫ",
        "Дата рождения / Date of birth: 07 11 1990",
        "№ документа / Document No: ID0000001",
        "Дата окончания / Date of expiry: 22 11 2034",
    )


def candidate(name, value, **kwargs):
    return FieldCandidate(
        field_name=name,
        raw_value=value,
        normalized_value=value,
        confidence=0.9,
        source_image_id="front",
        **kwargs,
    )


@pytest.mark.parametrize("value,expected", [("000000000", "0"), ("B", "7"), ("AB2134<<<", "5")])
def test_icao_check_digit_weights(value, expected):
    assert check_digit(value) == expected


def test_official_icao_td1_composite_vector():
    # ICAO 9303 Part 3, Appendix A, Example 4. Expected digit is independent
    # of our generator, so a shared generator/parser indexing error is caught.
    report = TD1Parser().parse(
        [
            "I<YTOD231458907<<<<<<<<<<<<<<<",
            "3407127M9507122YTO<<<<<<<<<<<2",
            "SPECIMEN<<TEST".ljust(30, "<"),
        ]
    )
    assert report.valid and report.checks["composite"]
    assert report.values["document_number"] == "D23145890"


def test_semantic_label_bounds_exclude_a_value_joined_in_same_ocr_row():
    from document_excel.infrastructure.physical_id import semantic_labels

    result = OcrResult(
        [
            TextBlock("Date of birth", 0.9, (20, 10, 160, 10), (1, 1, 1)),
            TextBlock("07 11 1990", 0.9, (20, 25, 120, 20), (1, 1, 1)),
        ],
        "synthetic",
        "original",
        400,
        200,
    )
    labels = semantic_labels(result)
    assert len(labels) == 1 and labels[0][1] == "birth_date"
    assert labels[0][0].bbox == (20, 10, 160, 10)


def test_translation_alias_to_the_right_does_not_cut_value_prefix():
    from PIL import Image, ImageDraw, ImageFont

    from document_excel.infrastructure.physical_id import PlasticCard, layout_candidates

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
    image = Image.new("RGB", (1000, 600), "white")
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype(str(font_path), 22)
    draw.text((390, 160), "ТЕСТ", font=font, fill="black")
    draw.text((390, 190), "TEST", font=font, fill="black")
    card = PlasticCard(np.asarray(image).copy(), np.eye(3), (1000, 600), True)
    result = OcrResult(
        [TextBlock("Surname", 0.95, (500, 130, 130, 15), (1, 1, 1))], "synthetic", "original", 1000, 600
    )

    class Engine:
        def recognize_physical_value(self, pixels, image_id, cancel, lang, kind):
            return ocr("ТЕСТ" if lang == "rus" else "TEST") if pixels.min() < 128 else ocr("")

    values = layout_candidates(card, result, Engine(), threading.Event(), DocumentSide.FRONT)
    names = [v for v in values if v.field_name == "surname"]
    assert {v.normalized_value for v in names} == {"ТЕСТ", "TEST"}


def test_value_line_excludes_next_small_label_and_preserves_original():
    from PIL import Image, ImageDraw, ImageFont

    from document_excel.infrastructure.physical_id import clean_value_line

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
    image = Image.new("RGB", (360, 100), "white")
    draw = ImageDraw.Draw(image)
    draw.text((10, 0), "07 11 1990", fill="black", font=ImageFont.truetype(str(font_path), 40))
    draw.text((10, 72), "NEXT FIELD LABEL", fill="black", font=ImageFont.truetype(str(font_path), 13))
    array = np.asarray(image).copy()
    result = clean_value_line(array, first_band=True)
    assert result.shape[0] < 55 and result.shape[1] > 180
    assert np.array_equal(array, np.asarray(image))


def test_cyrillic_name_fallback_uses_utf8_config_and_cleans_it(monkeypatch):
    from document_excel.infrastructure.ocr import TesseractEngine
    from document_excel.infrastructure.storage import Settings

    engine = TesseractEngine(Settings())
    calls = []

    def read(*args, **kwargs):
        path = Path(kwargs["extra_config"].split('"')[1])
        assert "ҮӨҢ" in path.read_text(encoding="utf-8")
        calls.append((args[4], path))
        return ocr("СИНТЕТИК" if len(calls) > 1 else "")

    monkeypatch.setattr(engine, "_recognize", read)
    result = engine.recognize_cyrillic_name(
        np.full((20, 100, 3), 200, np.uint8), "synthetic", threading.Event()
    )
    assert result.text == "СИНТЕТИК"
    assert [psm for psm, _ in calls] == [7, 13]
    assert all(not path.exists() for _, path in calls)
    assert engine.settings.languages == ["rus", "eng"] and engine.settings.psm == 6


def test_td1_complete_checks_and_pin_rule():
    result = TD1Parser().parse(td1())
    assert result.valid and all(result.checks.values())
    assert result.values["document_number"] == "ID0000001"
    assert result.values["personal_identifier"] == "00123456789012"
    assert result.values["surname_latin"] == "SINTETIKOV"
    assert result.values["birth_date"] == "901107"
    assert "passport_issued_date" not in result.values and "passport_issued_by" not in result.values


@pytest.mark.parametrize("line,position", [(0, 14), (1, 6), (1, 14), (1, 29)])
def test_every_check_digit_is_required(line, position):
    lines = td1()
    value = lines[line]
    lines[line] = value[:position] + str((int(value[position]) + 1) % 10) + value[position + 1 :]
    assert not TD1Parser().parse(lines).valid
    assert not mrz_candidates(TD1Parser().parse(lines), [])


@pytest.mark.parametrize("lines", [[], ["A" * 30] * 2, ["A" * 29] * 3, ["A" * 31] * 3, ["?" * 30] * 3])
def test_incomplete_or_illegal_mrz_is_not_padded_or_accepted(lines):
    assert not TD1Parser().parse(lines).valid


@pytest.mark.parametrize("birth", ["901332", "900231", "9O1107", "9I1107"])
def test_invalid_calendar_or_ambiguous_ocr_numeric_characters(birth):
    assert not TD1Parser().parse(td1(birth=birth)).valid


def test_mrz_normalizes_spaces_but_preserves_fillers():
    lines = td1()
    spaced = "\n".join(" ".join(s) for s in lines)
    assert find_mrz_lines(spaced) == lines
    assert TD1Parser().parse(find_mrz_lines(spaced)).valid


def test_unknown_optional_data_is_not_a_universal_pin():
    lines = td1()
    lines[0] = lines[0][:15] + "X" * 15
    lines[1] = lines[1][:29] + check_digit(lines[0][5:] + lines[1][:7] + lines[1][8:15] + lines[1][18:29])
    parsed = TD1Parser().parse(lines)
    assert parsed.valid and "personal_identifier" not in parsed.values


def test_mrz_century_needs_visual_full_date():
    parsed = TD1Parser().parse(td1(birth="081107"))
    result = mrz_candidates(parsed, [])
    assert not any(c.field_name == "birth_date" for c in result)
    assert parsed.warnings
    parsed = TD1Parser().parse(td1(birth="081107"))
    result = mrz_candidates(parsed, [candidate("birth_date", "2008-11-07")])
    assert next(c.normalized_value for c in result if c.field_name == "birth_date") == "2008-11-07"


def test_mrz_date_conflict_is_not_silently_replaced():
    parsed = TD1Parser().parse(td1())
    values = mrz_candidates(parsed, [candidate("birth_date", "1991-11-07")])
    case = ClientCase(mrz_reports={"back": parsed})
    apply_candidates(case, [candidate("birth_date", "1991-11-07")] + values, "kg_physical_id")
    assert case.data.birth_date == date(1991, 11, 7)
    assert case.evidence["birth_date"].status == FieldStatus.CONFLICT


def test_visual_and_mrz_number_conflict_keeps_sources_and_manual_override():
    parsed = TD1Parser().parse(td1(), "back")
    values = mrz_candidates(parsed, [])
    case = ClientCase(mrz_reports={"back": parsed})
    apply_candidates(case, [candidate("passport_number", "0000008")] + values, "kg_physical_id")
    evidence = case.evidence["passport_number"]
    assert evidence.status == FieldStatus.CONFLICT
    assert case.data.passport_number is None
    assert {c.source_method for c in evidence.candidates} == {"VISUAL_OCR", "MRZ"}
    case.data.passport_number = "0000008"
    evidence.source = "ручной ввод"
    evidence.status = FieldStatus.MANUALLY_CORRECTED
    apply_candidates(case, values, "kg_physical_id")
    assert case.data.passport_number == "0000008"


def test_physical_classification_and_independent_sides():
    profiles = load_document_profiles()
    assert classify(front(), profiles) == "kg_physical_id"
    assert DocumentSideClassifier().classify(front(), "kg_physical_id") == DocumentSide.FRONT
    back = ocr("Орган выдачи: ТЕСТ", "Дата выдачи: 22 11 2024", *td1(), image_id="back")
    assert classify(back, profiles) == "kg_physical_id"
    assert DocumentSideClassifier().classify(back, "kg_physical_id") == DocumentSide.BACK
    assert classify(ocr("КЫРГЫЗ", "IDENTITY"), profiles) != "kg_physical_id"


def test_physical_fields_official_scripts_and_spaced_dates():
    case = ClientCase()
    apply_candidates(case, FieldExtractor(KyrgyzPhysicalIDProfile()).extract(front()), "kg_physical_id")
    assert case.data.surname_cyrillic == "СИНТЕТИКОВ"
    assert case.data.surname_latin == "SINTETIKOV"
    assert case.data.first_name_latin == "TEST"
    assert case.data.patronymic_cyrillic == "ПРИМЕРОВИЧ"
    assert case.data.birth_date == date(1990, 11, 7) and case.data.expiration_date == date(2034, 11, 22)
    assert case.data.passport_type == "ID" and case.data.passport_number == "0000001"
    assert not any(e.status == FieldStatus.CONFLICT for e in case.evidence.values())


def test_birthplace_never_becomes_registration_and_absence_policy():
    case = ClientCase(
        images=[
            SourceImage(path=Path("front"), role=DocumentSide.FRONT),
            SourceImage(path=Path("back"), role=DocumentSide.BACK),
        ]
    )
    values = FieldExtractor(KyrgyzPhysicalIDProfile()).extract(ocr("Место рождения: ТЕСТОВАЯ СТРАНА"))
    apply_candidates(case, values, "kg_physical_id")
    assert case.data.place_of_birth == "ТЕСТОВАЯ СТРАНА"
    assert not case.data.registration_address.display()
    assert case.registration_status == RegistrationStatus.NOT_PRESENT
    case.data.registration_address = Address(raw="ВВЕДЁННЫЙ АДРЕС")
    case.evidence["registration_address"] = FieldEvidence(
        source="ручной ввод", status=FieldStatus.MANUALLY_CORRECTED
    )
    case.registration_status = RegistrationStatus.MANUALLY_ENTERED
    apply_candidates(case, values, "kg_physical_id")
    assert case.data.registration_address.raw == "ВВЕДЁННЫЙ АДРЕС"
    assert case.registration_status == RegistrationStatus.MANUALLY_ENTERED


@pytest.mark.parametrize("value", ["07 11 1990", "07.11.1990", "07/11/1990", "1990-11-07"])
def test_date_formats(value):
    assert parse_date(value) == date(1990, 11, 7)


def test_perspective_card_maps_regions_to_original_and_keeps_pixels():
    import cv2

    original = np.full((850, 1300, 3), 45, np.uint8)
    corners = np.array([[130, 130], [1180, 70], [1100, 750], [200, 700]], np.int32)
    cv2.fillConvexPoly(original, corners, (220, 235, 245))
    before = original.copy()
    prepared = prepare_plastic_card(original)
    assert prepared.detected and np.array_equal(original, before)
    box = prepared.original_bbox((0, 0, prepared.image.shape[1], prepared.image.shape[0]))
    assert box[0] < 210 and box[1] < 140 and box[2] > 900 and box[3] > 600
    assert prepare_region(prepared.image, "date").ndim == 2


def test_td1_extended_document_number():
    number = "AB0000001234"
    optional = number[9:] + check_digit(number) + "<"
    a = "I<KGZ" + number[:9] + "<" + optional.ljust(15, "<")
    b = td1()[1][:29]
    b += check_digit(a[5:] + b[:7] + b[8:15] + b[18:29])
    parsed = TD1Parser().parse([a, b, td1()[2]])
    assert parsed.valid and parsed.values["document_number"] == number
    assert "personal_identifier" not in parsed.values


@pytest.mark.parametrize(
    "letter,expected", [("<", True), ("K", False), ("C", False), ("L", False), ("I", False)]
)
def test_chevron_recognition_does_not_replace_real_letters(letter, expected):
    import cv2
    from PIL import Image, ImageDraw, ImageFont

    from document_excel.infrastructure.mrz import MRZRecognizer

    font_path = next(
        (
            p
            for p in [
                Path("C:/Windows/Fonts/cour.ttf"),
                Path("/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"),
            ]
            if p.exists()
        ),
        None,
    )
    if not font_path:
        pytest.skip("No monospace test font")
    image = Image.new("L", (100, 110), 255)
    ImageDraw.Draw(image).text((10, 0), letter, fill=0, font=ImageFont.truetype(str(font_path), 80))
    mask = cv2.threshold(np.asarray(image), 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
    x, y, w, h = cv2.boundingRect(mask)
    assert MRZRecognizer.is_filler(mask[y : y + h, x : x + w]) == expected


def test_card_rotation_maps_back_to_original():
    from document_excel.infrastructure.physical_id import PlasticCard, rotate_plastic_card

    image = np.zeros((300, 500, 3), np.uint8)
    card = PlasticCard(image, np.eye(3), (500, 300), True)
    rotated = rotate_plastic_card(card, 1)
    assert rotated.image.shape[:2] == (500, 300)
    assert rotated.original_bbox((0, 0, 300, 500)) == (0, 0, 500, 300)
    assert rotate_plastic_card(card, 4).image.shape == image.shape


class SyntheticProfileEngine:
    def __init__(self, fail=False):
        self.fail = fail
        self.calls = []

    def recognize(self, image, image_id, variant, cancel):
        self.calls.append(image_id)
        if image_id == "front":
            return front()
        return ocr(
            "Place of birth: TEST COUNTRY",
            "Authority: TEST OFFICE",
            "Date of issue: 22 11 2024",
            "Personal number: 00123456789012",
            image_id="back",
        )

    def recognize_with_psm(self, image, image_id, variant, cancel, psm):
        return self.recognize(image, image_id, variant, cancel)

    def recognize_mrz(self, image, image_id, cancel):
        if self.fail:
            from document_excel.domain.errors import AppError

            raise AppError("Synthetic optional OCR failure")
        return ocr(*td1(), image_id="back")


def test_profile_from_later_front_revisits_earlier_back(tmp_path):
    from PIL import Image

    from document_excel.application.workflows import RecognitionService
    from document_excel.infrastructure.storage import Settings

    paths = []
    for name in ("back", "front"):
        path = tmp_path / f"{name}.png"
        Image.new("RGB", (1600, 1000), "white").save(path)
        paths.append(SourceImage(id=name, path=path))
    engine = SyntheticProfileEngine()
    result = RecognitionService(Settings(), engine).run(paths, threading.Event(), lambda *_: None)
    assert result.document_type == "kg_physical_id"
    assert result.mrz_reports["back"].valid
    assert result.sides == {"front": DocumentSide.FRONT, "back": DocumentSide.BACK}
    assert engine.calls.count("back") >= 2


def test_optional_physical_failure_keeps_visual_data(tmp_path):
    from PIL import Image

    from document_excel.application.workflows import RecognitionService
    from document_excel.infrastructure.storage import Settings

    path = tmp_path / "synthetic.png"
    Image.new("RGB", (1600, 1000), "white").save(path)
    result = RecognitionService(Settings(), SyntheticProfileEngine(fail=True)).run(
        [SourceImage(id="back", path=path)],
        threading.Event(),
        lambda *_: None,
        document_type="kg_physical_id",
    )
    assert any(c.field_name == "personal_identifier" for c in result.candidates)
    assert any("MRZ" in w for w in result.warnings)


def test_additional_document_keeps_role_and_does_not_require_mrz(tmp_path):
    from PIL import Image

    from document_excel.application.workflows import RecognitionService
    from document_excel.infrastructure.storage import Settings

    path = tmp_path / "additional.png"
    Image.new("RGB", (1600, 1000), "white").save(path)
    result = RecognitionService(Settings(), SyntheticProfileEngine()).run(
        [SourceImage(id="back", path=path, role=DocumentSide.ADDITIONAL)],
        threading.Event(),
        lambda *_: None,
        document_type="kg_physical_id",
    )
    assert result.sides["back"] == DocumentSide.ADDITIONAL
    assert not result.mrz_reports
    assert all(c.source_side == DocumentSide.ADDITIONAL for c in result.candidates)


def test_real_ocr_plastic_mrz_to_excel(tmp_path):
    from document_excel.infrastructure.ocr import TesseractEngine
    from document_excel.infrastructure.physical_verification import run_physical_self_test
    from document_excel.infrastructure.storage import Settings

    diagnostic = TesseractEngine(Settings()).diagnose()
    if diagnostic.error or not {"rus", "eng"} <= set(diagnostic.languages):
        pytest.skip("Local Tesseract rus+eng unavailable")
    result = run_physical_self_test(tmp_path, Settings(ocr_engine="tesseract"))
    assert result["ok"], result
    assert result["embedded_document_photos"] == 2
