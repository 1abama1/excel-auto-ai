"""Synthetic service tests; OCR output is never written into reports or logs."""

from __future__ import annotations

import threading

import numpy as np
import pytest

from document_excel.application.workflows import RecognitionOutput, RecognitionService
from document_excel.domain.errors import AppError, Cancelled
from document_excel.domain.models import DocumentSide, FieldStatus, SourceImage
from document_excel.infrastructure.ocr import OcrDiagnostic, OcrResult, TesseractEngine, TextBlock
from document_excel.infrastructure.preprocessing import DocumentRegionDetection, PreparedImage, Quality
from document_excel.infrastructure.storage import Settings


def ocr_result(lines, source_id="source", variant="original", width=800, height=600):
    return OcrResult(
        [
            TextBlock(text, 0.96, (20, 20 + index * 50, 600, 30), (1, 1, index))
            for index, text in enumerate(lines)
        ],
        source_id,
        variant,
        width,
        height,
    )


class FakeEngine:
    def __init__(self, primary, alternative=None, *, cancel_after_primary=False, alternative_error=False):
        self.primary = primary
        self.alternative = alternative
        self.calls = []
        self.cancel_after_primary = cancel_after_primary
        self.alternative_error = alternative_error

    def recognize(self, image, source_id, variant, cancel):
        self.calls.append((6, source_id, variant, image.shape[:2]))
        if self.cancel_after_primary:
            cancel.set()
        lines = self.primary.get(source_id, []) if isinstance(self.primary, dict) else self.primary
        return ocr_result(lines, source_id, variant, image.shape[1], image.shape[0])

    def recognize_with_psm(self, image, source_id, variant, cancel, psm=11):
        self.calls.append((psm, source_id, variant, image.shape[:2]))
        if self.alternative_error:
            raise AppError("Synthetic optional failure")
        lines = self.alternative if self.alternative is not None else self.primary
        return ocr_result(lines, source_id, variant, image.shape[1], image.shape[0])


class FakePreprocessor:
    def __init__(self, *, cropped=True, physical_warning=False):
        self.original = np.full((1000, 1200, 3), 225, np.uint8)
        gray = np.full((600, 800), 225, np.uint8) if cropped else np.full((1000, 1200), 225, np.uint8)
        geometry_warning = "Границы документа не определены уверенно."
        self.prepared = PreparedImage(
            self.original,
            {"original": self.original, "grayscale": gray},
            Quality(150, 225, 80, [geometry_warning] + (["Возможно размытие."] if physical_warning else [])),
            DocumentRegionDetection(
                (100, 200, 800, 600) if cropped else (0, 0, 1200, 1000), 0.9, [geometry_warning], cropped
            ),
            {"original": (0, 0), "grayscale": (100, 200) if cropped else (0, 0)},
            ("grayscale",),
        )

    def prepare(self, path, mode, rotation):
        return self.prepared


def service(engine, *, physical_warning=False, settings=None):
    value = RecognitionService(settings or Settings(), engine)
    value.preprocessor = FakePreprocessor(physical_warning=physical_warning)
    return value


def source(tmp_path, role=DocumentSide.UNKNOWN, source_id="source"):
    return SourceImage(path=tmp_path / "synthetic.png", id=source_id, role=role)


FRONT = [
    "KYRGYZ REPUBLIC IDENTITY",
    "SURNAME: SYNTHETIC",
    "GIVEN NAME: TEST",
    "DATE OF BIRTH: 01.02.1990",
]
BACK = [
    "KYRGYZ REPUBLIC IDENTITY",
    "PERSONAL NUMBER: 00123456789012",
    "DOCUMENT NUMBER: AA0000001",
    "AUTHORITY: SYNTHETIC OFFICE",
    "DATE OF ISSUE: 03.04.2020",
]


def test_crop_boxes_map_to_original_and_geometry_warning_does_not_demote(tmp_path):
    engine = FakeEngine(FRONT)
    progress = []
    output = service(engine).run([source(tmp_path)], threading.Event(), lambda *args: progress.append(args))
    assert len(engine.calls) == 1  # Good labelled fields need no alternative.
    assert engine.calls[0][2:] == ("grayscale", (600, 800))
    candidate = next(c for c in output.candidates if c.field_name == "surname")
    assert candidate.source_bbox == (120, 270, 600, 30)
    assert candidate.validation_status == FieldStatus.VALID
    assert candidate.source_side == DocumentSide.FRONT
    assert output.sides == {"source": DocumentSide.FRONT}
    assert output.ocr_texts["source"].startswith("KYRGYZ")
    assert progress[-1][0] == 100 and all(p[0] < 100 for p in progress[:-1])
    assert all(
        any(stage in message for _, message in progress)
        for stage in (
            "Подготовка изображения",
            "Поиск документа",
            "Распознавание текста",
            "Извлечение данных",
            "Проверка",
        )
    )


def test_physical_quality_warning_still_lowers_confidence(tmp_path):
    output = service(FakeEngine(FRONT), physical_warning=True).run(
        [source(tmp_path)],
        threading.Event(),
        lambda *_: None,
    )
    assert all(c.validation_status == FieldStatus.REVIEW for c in output.candidates)
    assert all(c.confidence < 0.82 for c in output.candidates)


def test_roi_maps_boxes_to_original_and_only_recognizes_selection(tmp_path):
    engine = FakeEngine(["PERSONAL NUMBER: 00123456789012"])
    output = service(engine).run(
        [source(tmp_path)], threading.Event(), lambda *_: None, roi=(50, 500, 1000, 90)
    )
    assert engine.calls == [(6, "source", "roi", (90, 1000))]
    assert output.candidates[0].source_bbox == (70, 520, 600, 30)
    assert output.candidates[0].variant == "roi"


def test_roi_clips_negative_origin_instead_of_wrapping_numpy_slice(tmp_path):
    engine = FakeEngine(["PERSONAL NUMBER: 00123456789012"])
    output = service(engine).run(
        [source(tmp_path)], threading.Event(), lambda *_: None, roi=(-15, -20, 800, 200)
    )
    assert engine.calls[0][-1] == (180, 785)
    assert output.candidates[0].source_bbox == (20, 20, 600, 30)


def test_outside_roi_does_not_call_engine(tmp_path):
    engine = FakeEngine(FRONT)
    with pytest.raises(AppError, match="область"):
        service(engine).run([source(tmp_path)], threading.Event(), lambda *_: None, roi=(1500, 1400, 50, 50))
    assert not engine.calls


def test_user_assigned_role_overrides_classifier_in_candidates_and_output(tmp_path):
    output = service(FakeEngine(FRONT)).run(
        [source(tmp_path, DocumentSide.ADDITIONAL)],
        threading.Event(),
        lambda *_: None,
    )
    assert output.sides["source"] == DocumentSide.ADDITIONAL
    assert all(candidate.source_side == DocumentSide.ADDITIONAL for candidate in output.candidates)


def test_front_and_back_remain_associated_with_their_source(tmp_path):
    engine = FakeEngine({"front": FRONT, "back": BACK})
    output = service(engine).run(
        [source(tmp_path, source_id="front"), source(tmp_path, source_id="back")],
        threading.Event(),
        lambda *_: None,
    )
    assert output.sides == {"front": DocumentSide.FRONT, "back": DocumentSide.BACK}
    assert set(output.ocr_texts) == {"front", "back"}
    assert all(c.source_image_id == "front" for c in output.candidates if c.field_name == "surname")
    assert all(c.source_image_id == "back" for c in output.candidates if c.field_name == "passport_number")
    assert len(engine.calls) == 2


def test_one_sparse_alternative_wins_by_valid_field_coverage(tmp_path):
    engine = FakeEngine(["PERSONAL NUMBER: 00123456789012"] + ["UNLABELLED NOISE" * 20] * 20, BACK)
    output = service(engine).run([source(tmp_path)], threading.Event(), lambda *_: None)
    assert [call[0] for call in engine.calls] == [6, 11]
    assert {c.field_name for c in output.candidates} == {
        "personal_identifier",
        "passport_number",
        "passport_type",
        "passport_issued_by",
        "passport_issued_date",
    }
    assert output.sides["source"] == DocumentSide.BACK
    assert "UNLABELLED NOISE" not in output.ocr_texts["source"]


def test_longer_but_invalid_alternative_does_not_replace_usable_primary(tmp_path):
    engine = FakeEngine(
        ["PERSONAL NUMBER: 00123456789012"],
        [
            "PERSONAL NUMBER: INVALID",
            "DATE OF BIRTH: 31.02.2000",
            "DOCUMENT NUMBER: @@@@",
        ]
        + ["RANDOM LONG LINE" * 20] * 20,
    )
    output = service(engine).run([source(tmp_path)], threading.Event(), lambda *_: None)
    assert [call[0] for call in engine.calls] == [6, 11]
    assert len(output.candidates) == 1
    assert output.candidates[0].field_name == "personal_identifier"
    assert output.candidates[0].validation_status == FieldStatus.VALID


def test_alternative_failure_keeps_primary_result(tmp_path):
    engine = FakeEngine(["PERSONAL NUMBER: 00123456789012"], alternative_error=True)
    output = service(engine).run([source(tmp_path)], threading.Event(), lambda *_: None)
    assert len(output.candidates) == 1
    assert any("дополнительная проверка" in warning for warning in output.warnings)


def test_empty_primary_can_recover_once(tmp_path):
    engine = FakeEngine([], BACK)
    output = service(engine).run([source(tmp_path)], threading.Event(), lambda *_: None)
    assert len(output.candidates) == 5
    assert len(engine.calls) == 2


def test_cancellation_before_alternative_prevents_result_and_success_progress(tmp_path):
    engine = FakeEngine(["PERSONAL NUMBER: 00123456789012"], BACK, cancel_after_primary=True)
    progress = []
    with pytest.raises(Cancelled):
        service(engine).run([source(tmp_path)], threading.Event(), lambda *args: progress.append(args))
    assert len(engine.calls) == 1
    assert not any(percent == 100 for percent, _ in progress)


def test_preconfigured_sparse_mode_does_not_run_same_mode_twice(tmp_path):
    engine = FakeEngine(["PERSONAL NUMBER: 00123456789012"])
    service(engine, settings=Settings(psm=11)).run([source(tmp_path)], threading.Event(), lambda *_: None)
    assert len(engine.calls) == 1


def test_legacy_output_constructor_keeps_independent_defaults():
    first = RecognitionOutput([], "unknown", [], False, 0)
    second = RecognitionOutput([], "unknown", [], False, 0)
    first.sides["source"] = DocumentSide.FRONT
    first.ocr_texts["source"] = "SYNTHETIC"
    assert not second.sides and not second.ocr_texts


def test_optional_psm_does_not_mutate_settings(monkeypatch):
    engine = TesseractEngine(Settings(languages=["eng"], psm=6))
    monkeypatch.setattr(engine, "diagnose", lambda: OcrDiagnostic("synthetic.exe", "5.0", ["eng"]))
    calls = []

    def fake_data(image, **kwargs):
        calls.append(kwargs["config"])
        return {
            "text": ["SYNTHETIC"],
            "conf": [96],
            "left": [1],
            "top": [2],
            "width": [100],
            "height": [20],
            "block_num": [1],
            "par_num": [1],
            "line_num": [1],
        }

    monkeypatch.setattr("document_excel.infrastructure.ocr.pytesseract.image_to_data", fake_data)
    image = np.full((200, 300), 225, np.uint8)
    engine.recognize_with_psm(image, "source", "original", threading.Event(), 11)
    engine.recognize(image, "source", "original", threading.Event())
    assert calls == ["--oem 1 --psm 11", "--oem 1 --psm 6"]
    assert engine.settings.psm == 6
    with pytest.raises(AppError, match="сегментации"):
        engine.recognize_with_psm(image, "source", "original", threading.Event(), 99)


def test_tiny_crop_upscale_keeps_input_coordinate_system(monkeypatch):
    engine = TesseractEngine(Settings(languages=["eng"]))
    monkeypatch.setattr(engine, "diagnose", lambda: OcrDiagnostic("synthetic.exe", "5.0", ["eng"]))
    observed_shapes = []

    def fake_data(image, **kwargs):
        observed_shapes.append(image.shape[:2])
        return {
            "text": ["SYNTHETIC"],
            "conf": [96],
            "left": [30],
            "top": [60],
            "width": [120],
            "height": [30],
            "block_num": [1],
            "par_num": [1],
            "line_num": [1],
        }

    monkeypatch.setattr("document_excel.infrastructure.ocr.pytesseract.image_to_data", fake_data)
    result = engine.recognize(np.full((200, 300), 225, np.uint8), "source", "original", threading.Event())
    assert observed_shapes == [(600, 900)]
    assert (result.width, result.height) == (300, 200)
    assert result.blocks[0].bbox == (10, 20, 40, 10)


def test_legacy_engine_without_optional_method_remains_compatible(tmp_path):
    class LegacyEngine:
        def recognize(self, image, source_id, variant, cancel):
            return ocr_result(["PERSONAL NUMBER: 00123456789012"], source_id, variant)

    output = service(LegacyEngine()).run([source(tmp_path)], threading.Event(), lambda *_: None)
    assert len(output.candidates) == 1
    assert output.candidates[0].field_name == "personal_identifier"


def test_cancelled_run_never_loads_a_source(tmp_path):
    class NeverPrepare:
        def prepare(self, *args):
            pytest.fail("Cancelled run read an image")

    instance = service(FakeEngine(FRONT))
    instance.preprocessor = NeverPrepare()
    cancel = threading.Event()
    cancel.set()
    with pytest.raises(Cancelled):
        instance.run([source(tmp_path)], cancel, lambda *_: None)


def test_noninteger_upscale_rounding_keeps_boxes_inside_input(monkeypatch):
    engine = TesseractEngine(Settings(languages=["eng"]))
    monkeypatch.setattr(engine, "diagnose", lambda: OcrDiagnostic("synthetic.exe", "5.0", ["eng"]))

    def fake_data(image, **kwargs):
        return {
            "text": ["SYNTHETIC"],
            "conf": [96],
            "left": [0],
            "top": [0],
            "width": [image.shape[1]],
            "height": [image.shape[0]],
            "block_num": [1],
            "par_num": [1],
            "line_num": [1],
        }

    monkeypatch.setattr("document_excel.infrastructure.ocr.pytesseract.image_to_data", fake_data)
    result = engine.recognize(np.full((200, 700), 225, np.uint8), "source", "original", threading.Event())
    assert result.blocks[0].bbox == (0, 0, 700, 200)
