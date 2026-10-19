"""Independent digital semantics, geometry, staging and failure regressions."""

from importlib.resources import files
from types import SimpleNamespace

import numpy as np
import pytest
from PIL import Image, ImageDraw, ImageFont

from document_excel.application.ocr_state import OcrState
from document_excel.application.workflows import RecognitionOutput
from document_excel.domain.models import ClientCase, DocumentSide, FieldStatus, SourceImage
from document_excel.infrastructure.digital_fields import DigitalDocumentExtractor
from document_excel.infrastructure.digital_region import DigitalIDRegionDetector
from document_excel.infrastructure.digital_verification import synthetic_digital_images
from document_excel.infrastructure.document_profiles import KyrgyzDigitalIDProfile, KyrgyzPhysicalIDProfile
from document_excel.infrastructure.extraction import FieldExtractor, apply_candidates
from document_excel.infrastructure.ocr_contract import OcrResult, TextBlock
from document_excel.infrastructure.preprocessing import load_image


def text_result(lines, width=1000, height=1000):
    blocks = [TextBlock(value, 0.98, box, (1, 1, index)) for index, (value, box) in enumerate(lines)]
    return OcrResult(blocks, "front", "digital-card", width, height, "stub", "PP-OCRv5 synthetic")


def front(*, birth="07.11.1990", patronymic="ТЕСТОВИЧ", split_forms=False):
    rows = [
        ("Аты / Имя / Name", (300, 100, 500, 20)),
        ("ПРИМЕР / PRIMER", (300, 132, 400, 24)),
        ("Фамилиясы / Фамилия / Surname", (300, 200, 600, 20)),
        ("СИНТЕТИКОВ / SINTETIKOV", (300, 232, 560, 24)),
        ("Атасынын аты / Отчество / Patronymic", (300, 300, 600, 20)),
        ("Жынысы / Пол / Sex", (300, 400, 400, 20)),
        ("Э / М", (300, 432, 120, 24)),
        ("Жарандыгы / Гражданство / Citizenship", (300, 500, 600, 20)),
        ("Кыргыз Республикасы", (300, 532, 500, 24)),
        ("Улуту / Национальность / Nationality", (300, 600, 600, 20)),
        ("ОРУС", (300, 632, 150, 24)),
        ("Туулган күнү / Дата рождения", (300, 780, 600, 20)),
        ("/ Date of birth", (300, 805, 400, 20)),
        (birth, (300, 840, 280, 24)),
        ("1", (900, 850, 20, 24)),
    ]
    if patronymic:
        rows.append((patronymic, (300, 332, 500, 24)))
    if split_forms:
        rows[1:2] = [("ПРИМЕР", (300, 132, 180, 24)), ("PRIMER", (300, 165, 180, 24))]
    return text_result(rows)


def back():
    result = text_result(
        [
            ("Берген мекеме / Орган выдачи / Authority", (80, 100, 600, 20)),
            ("МКК 012345", (80, 132, 300, 24)),
            ("Берилген күнү / Дата выдачи", (80, 200, 600, 20)),
            ("/ Date of issue", (80, 225, 300, 20)),
            ("22.11.2024", (80, 260, 300, 24)),
            ("Жеке номери / Персональный номер", (80, 340, 600, 20)),
            ("/ Personal number", (80, 365, 300, 20)),
            ("00123456789012", (80, 400, 400, 24)),
            ("Документтин № / № Документа", (80, 480, 600, 20)),
            ("/ Document №", (80, 505, 300, 20)),
            ("ID0000001", (80, 540, 300, 24)),
            ("Колдонуу мөөнөтү / Срок действия", (80, 620, 600, 20)),
            ("/ Date of expiry", (80, 645, 300, 20)),
            ("22.11.2034", (80, 680, 300, 24)),
            ("Катталган жери / Адрес прописки / Address", (80, 780, 700, 20)),
            ("ТЕСТОВАЯ ОБЛАСТЬ, ТЕСТОВЫЙ РАЙОН,", (80, 820, 700, 24)),
            ("улица Примерная, дом 7", (80, 852, 600, 24)),
            ("2", (900, 900, 20, 24)),
        ]
    )
    result.image_id = "back"
    return result


def resolve(*results):
    profile = KyrgyzDigitalIDProfile()
    candidates = [c for r in results for c in FieldExtractor(profile).extract(r)]
    case = ClientCase()
    apply_candidates(case, candidates, "kg_digital_id", True)
    return case, candidates


def test_digital_profile_is_independent_and_has_no_mrz_or_absence_rule():
    digital, physical = KyrgyzDigitalIDProfile(), KyrgyzPhysicalIDProfile()
    assert digital.mrz_strategy == "none" and physical.mrz_strategy == "TD1"
    assert digital.registration_policy == "detect" and physical.registration_policy == "not_printed"
    assert digital.front_layout.index("first_name") < digital.front_layout.index("surname")
    assert digital.ocr_regions is not physical.ocr_regions


def test_digital_address_keeps_unparsed_fragments_and_maps_explicit_city():
    from document_excel.domain.digital_address import digital_address_from_text
    from document_excel.infrastructure.excel import ExcelDataAssembler

    raw = "Тестовая Республика, г. Пример, неполный р-?, улица Тестовая, дом 007"
    parsed = digital_address_from_text(raw)
    assert parsed.raw == raw and parsed.house == "007"
    assert parsed.city == "г. Пример" and parsed.district == ""
    assert "неполный р-?" in parsed.extra
    case = ClientCase()
    case.data.document_type = "kg_digital_id"
    case.data.registration_address = parsed
    data = ExcelDataAssembler().assemble(case)
    assert data["registration_address.region"] == "г. Пример"
    assert "007" in data["registration_address.street"]


def test_digital_address_explicit_prefixes_work_without_commas():
    from document_excel.domain.digital_address import digital_address_from_text

    raw = "Тестовая Республика г. Пример, Тестовый р-н, улица Тестовая дом 007"
    parsed = digital_address_from_text(raw)
    assert parsed.raw == raw and parsed.country == "Тестовая Республика"
    assert parsed.city == "г. Пример" and parsed.house == "007"
    assert parsed.street == "улица Тестовая" and parsed.district == "Тестовый р-н"


def test_wrapped_same_script_compound_name_is_not_two_alternatives():
    result = front()
    result.blocks = [b for b in result.blocks if b.text != "СИНТЕТИКОВ / SINTETIKOV"]
    result.blocks.extend(
        [
            TextBlock("СИНТЕТИКОВ", 0.99, (300, 232, 250, 24), (1, 1, 80)),
            TextBlock("ТЕСТ", 0.99, (300, 262, 200, 24), (1, 1, 81)),
        ]
    )
    case, _ = resolve(result)
    assert case.data.surname_cyrillic == "СИНТЕТИКОВ ТЕСТ"


def test_manual_digital_reread_uses_only_selected_pixels_and_semantic_field():
    import threading

    from document_excel.infrastructure.digital_reread import reread_digital_field

    calls = []

    class Engine:
        def recognize(self, pixels, image_id, variant, cancel):
            calls.append(pixels.shape)
            return OcrResult(
                [TextBlock("ПРИМЕР / PRIMER", 0.99, (8, 4, 150, 20), (1, 1, 1))],
                image_id,
                variant,
                pixels.shape[1],
                pixels.shape[0],
            )

    source = SimpleNamespace(id="front", role=DocumentSide.FRONT)
    values = reread_digital_field(
        np.zeros((1200, 600, 3), np.uint8),
        source,
        "first_name",
        (100, 400, 200, 35),
        Engine(),
        threading.Event(),
    )
    assert len(calls) == 1 and calls[0][:2] == (70, 400)
    assert {c.normalized_value for c in values} == {"ПРИМЕР", "PRIMER"}
    assert all(c.field_name == "first_name" and c.source_bbox[1] >= 400 for c in values)


def test_front_name_order_scripts_gender_citizenship_and_bottom_birth_are_independent():
    case, _ = resolve(front())
    assert case.data.full_name == "СИНТЕТИКОВ ПРИМЕР ТЕСТОВИЧ"
    assert case.data.surname_cyrillic == "СИНТЕТИКОВ" and case.data.surname_latin == "SINTETIKOV"
    assert case.data.first_name_cyrillic == "ПРИМЕР" and case.data.first_name_latin == "PRIMER"
    assert case.data.gender == "MALE" and case.data.citizenship == "KGZ"
    assert str(case.data.birth_date) == "1990-11-07"
    assert all(c.normalized_value != "ОРУС" for e in case.evidence.values() for c in e.candidates)


@pytest.mark.parametrize("date", ["07.11.1990", "07 11 1990", "07/11/1990", "07.11.2015"])
def test_multiline_birth_caption_and_valid_minor_date(date):
    case, _ = resolve(front(birth=date))
    assert case.data.birth_date.day == 7 and case.data.birth_date.month == 11
    assert not case.evidence["birth_date"].status == FieldStatus.INVALID


def test_missing_patronymic_keeps_two_parts_and_never_guesses_a_third():
    case, _ = resolve(front(patronymic=None))
    assert case.data.full_name == "СИНТЕТИКОВ ПРИМЕР" and case.data.patronymic is None


def test_forms_on_separate_ocr_rows_are_not_lost():
    case, _ = resolve(front(split_forms=True))
    assert case.data.first_name_cyrillic == "ПРИМЕР" and case.data.first_name_latin == "PRIMER"


def test_back_strings_dates_and_multiline_registration_have_their_own_bands():
    case, _ = resolve(back())
    assert case.data.personal_identifier == "00123456789012"
    assert case.data.document_number == "ID0000001"
    assert case.data.passport_issued_by == "МКК" and case.data.passport_department_code == "012345"
    assert (
        str(case.data.passport_issued_date) == "2024-11-22" and str(case.data.expiration_date) == "2034-11-22"
    )
    assert (
        "ТЕСТОВЫЙ РАЙОН" in case.data.registration_address.raw
        and "дом 7" in case.data.registration_address.raw
    )
    assert (
        "2034" not in case.data.registration_address.raw
        and not case.data.registration_address.raw.endswith("2")
    )
    assert str(case.registration_status) != "NOT_PRESENT"


def test_a_misspelled_translated_caption_is_not_an_inline_surname():
    result = front()
    result.blocks[2].text = "Фамилиясы / Фамилия / Sumame"
    case, _ = resolve(result)
    assert case.data.surname_cyrillic == "СИНТЕТИКОВ" and case.data.surname_latin == "SINTETIKOV"
    assert case.evidence["surname"].status != FieldStatus.CONFLICT


def test_one_rejected_field_does_not_clear_other_values():
    case, _ = resolve(front(birth="31.02.1990"), back())
    assert case.data.birth_date is None
    assert case.evidence["birth_date"].candidates and case.data.full_name
    assert case.data.personal_identifier == "00123456789012"


@pytest.mark.parametrize("small,overlay", [(False, False), (True, False), (False, True)])
def test_mobile_detector_retains_card_bottom_and_maps_to_original(tmp_path, small, overlay):
    paths, _ = synthetic_digital_images(tmp_path, small=small, overlay=overlay)
    for path in paths:
        image = load_image(path)
        region = DigitalIDRegionDetector().prepare(image)
        assert region is not None and region.mobile_context
        x, y, w, h = region.bbox
        factor = 0.75 if small else 1
        assert x <= 50 * factor and y <= 305 * factor and x + w >= 670 * factor and y + h >= 995 * factor
        full = region.original_bbox((0, 0, region.image.shape[1], region.image.shape[0]))
        assert full == region.bbox
        assert np.array_equal(image, load_image(path))


def test_weak_bounds_fall_back_to_full_image_without_erasing_dark_text():
    pixels = np.full((1280, 576, 3), 190, np.uint8)
    pixels[500:530, 120:300] = 0
    detection = SimpleNamespace(detected=False, confidence=0.2, bbox=(120, 500, 180, 30))
    region = DigitalIDRegionDetector().prepare(pixels, forced=True, detection=detection)
    assert region.bbox == (0, 0, 576, 1280) and not region.detected
    assert region.image.min() == 0


def test_digital_fields_are_block_only_and_do_not_require_an_engine():
    candidates, layout = DigitalDocumentExtractor().extract(front(), KyrgyzDigitalIDProfile())
    assert candidates and layout.side() == DocumentSide.FRONT


@pytest.mark.parametrize("separator", [" / ", " "])
def test_slash_is_localized_from_glyphs_without_estimated_character_widths(separator):
    from document_excel.infrastructure.digital_recovery import bilingual_split

    font = ImageFont.truetype(
        str(files("document_excel").joinpath("resources/fonts/NotoSans-Regular.ttf")), 32
    )
    image = Image.new("RGB", (500, 70), "white")
    draw = ImageDraw.Draw(image)
    draw.text((6, 2), "ПРИМЕР" + separator + "PRIMER", font=font, fill="black")
    split = bilingual_split(np.array(image))
    if "/" in separator:
        assert split is not None
        assert 6 + font.getlength("ПРИМЕР ") <= split <= 6 + font.getlength("ПРИМЕР / ")
    else:
        assert split is None


def test_generator_stage_timing_does_not_count_consumer_pause(monkeypatch):
    from document_excel.infrastructure import paddle_ocr

    clock = [0.0]
    monkeypatch.setattr(paddle_ocr.time, "perf_counter", lambda: clock[0])

    class Predictor:
        def predict(self):
            clock[0] += 2
            yield "first"
            clock[0] += 3
            yield "second"

    times = {}
    iterator = paddle_ocr._TimedPredictor(Predictor(), times, "recognition").predict()
    assert next(iterator) == "first"
    clock[0] += 100
    assert next(iterator) == "second"
    with pytest.raises(StopIteration):
        next(iterator)
    assert times["recognition"] == 5


def test_pending_identity_has_a_readonly_preview_but_no_live_merge_or_export(tmp_path):
    files = []
    for name in ("front", "back"):
        path = tmp_path / (name + ".png")
        path.write_bytes(name.encode())
        files.append(SourceImage(id=name, path=path))
    case = ClientCase(images=files)
    case.data.document_type = "kg_digital_id"
    case.document_type_manual = True
    _, candidates = resolve(front(), back())
    state = OcrState()
    ticket = state.begin(case)
    assert state.commit(case, ticket.id, RecognitionOutput(candidates, "kg_digital_id", [], True, 1))
    assert not case.data.full_name and case.identity_status == "review"
    preview = state.preview(case)
    assert preview.data.full_name == "СИНТЕТИКОВ ПРИМЕР ТЕСТОВИЧ"
    assert preview.data.personal_identifier == "00123456789012" and not preview.confirmed
    assert not case.data.full_name and not case.evidence
    assert state.accept_identity(case) and case.data.full_name and case.identity_status == "user_confirmed"


@pytest.mark.parametrize("change", ["profile", "bytes", "client_run"])
def test_pending_preview_cannot_publish_stale_input(tmp_path, change):
    paths, _ = synthetic_digital_images(tmp_path)
    case = ClientCase(images=[SourceImage(id="front", path=paths[0]), SourceImage(id="back", path=paths[1])])
    case.data.document_type = "kg_digital_id"
    case.document_type_manual = True
    _, candidates = resolve(front(), back())
    state = OcrState()
    run = state.begin(case)
    state.commit(case, run.id, RecognitionOutput(candidates, "kg_digital_id", [], True, 1))
    if change == "profile":
        case.data.document_type = "kg_physical_id"
    elif change == "bytes":
        paths[0].write_bytes(b"changed")
    else:
        state.begin(case)
    assert state.preview(case) is None and not state.accept_identity(case)
    assert not case.data.full_name
