"""Actual Qt integration for plastic profiles, MRZ and client isolation."""

from pathlib import Path

import pytest
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication, QPlainTextEdit, QPushButton

from document_excel.domain.models import ClientCase, DocumentSide, RegistrationStatus, SourceImage
from document_excel.infrastructure.extraction import apply_candidates
from document_excel.infrastructure.mrz import TD1Parser
from document_excel.infrastructure.physical_verification import synthetic_plastic_images
from document_excel.infrastructure.storage import SettingsRepository
from document_excel.presentation.main_window import MainWindow


@pytest.fixture
def physical_window(qtbot, tmp_path):
    window = MainWindow(SettingsRepository(tmp_path / "app-data"))
    qtbot.addWidget(window)
    window.show()
    return window


def test_manual_profile_choice_and_undo(physical_window):
    window = physical_window
    window.upload_document_type.setCurrentIndex(window.upload_document_type.findData("kg_physical_id"))
    assert window.current_case.data.document_type == "kg_physical_id"
    assert window.current_case.document_type_manual
    window.undo_stack.undo()
    assert (
        window.current_case.data.document_type == "unknown" and not window.current_case.document_type_manual
    )
    window.undo_stack.redo()
    assert window.current_case.document_type_manual


def test_registration_prompt_and_manual_official_script_edits(physical_window):
    window = physical_window
    case = window.current_case
    case.data.document_type = "kg_physical_id"
    case.registration_status = RegistrationStatus.NOT_PRESENT
    window.go_step(2)
    window.refresh_form()
    assert not window.registration_notice.isHidden() and not window.registration_actions.isHidden()
    window.apply_field_edit("surname_cyrillic", "СИНТЕТИКОВ")
    window.apply_field_edit("first_name_cyrillic", "ТЕСТ")
    window.apply_field_edit("surname_latin", "SINTETIKOV")
    assert case.data.surname == "СИНТЕТИКОВ" and case.data.full_name == "СИНТЕТИКОВ ТЕСТ"
    assert case.data.surname_latin == "SINTETIKOV"
    assert case.evidence["surname_cyrillic"].source_method == "MANUAL_INPUT"
    window.undo_stack.undo()
    assert not case.data.surname_latin


def test_side_confirmation_applies_only_current_client(physical_window, qtbot, tmp_path):
    window = physical_window
    paths = synthetic_plastic_images(tmp_path / "synthetic")
    case = window.current_case
    case.images = [
        SourceImage(id="front", path=paths[0], role=DocumentSide.BACK),
        SourceImage(id="back", path=paths[1], role=DocumentSide.FRONT),
    ]
    case.side_proposals = {"front": DocumentSide.FRONT, "back": DocumentSide.BACK}
    other = ClientCase(images=[SourceImage(path=paths[0], role=DocumentSide.BACK)])
    window.batch.cases.append(other)

    def apply():
        dialog = QApplication.activeModalWidget()
        next(b for b in dialog.findChildren(QPushButton) if b.text().startswith("Применить")).click()

    QTimer.singleShot(0, apply)
    window.case_actions.review_sides()
    assert [image.role for image in case.images] == [DocumentSide.FRONT, DocumentSide.BACK]
    assert not case.side_proposals and other.images[0].role == DocumentSide.BACK


def test_mrz_is_on_demand_and_isolated_per_client(physical_window):
    window = physical_window
    case = window.current_case
    case.data.document_type = "kg_physical_id"
    # An intentionally invalid synthetic report must not get a green check.
    report = TD1Parser().parse(["I<KGZ" + "<" * 25] * 3, "back")
    case.mrz_reports = {"back": report}
    window.refresh_form()
    assert "не прошла" in window.mrz_status.text()

    def inspect():
        dialog = QApplication.activeModalWidget()
        assert report.lines[0] in dialog.findChild(QPlainTextEdit).toPlainText()
        next(b for b in dialog.findChildren(QPushButton) if b.text() == "Закрыть").click()

    QTimer.singleShot(0, inspect)
    window.case_actions.show_mrz()
    window.batch.cases.append(ClientCase())
    window.select_case(1)
    assert not window.current_case.mrz_reports
    assert window.mrz_status.isHidden()


def test_removing_image_clears_mrz_and_derived_names(tmp_path):
    from document_excel.application.workflows import CaseService
    from document_excel.domain.models import FieldCandidate

    case = ClientCase(images=[SourceImage(id="old", path=Path("unused"))])
    apply_candidates(
        case,
        [
            FieldCandidate(
                field_name="surname",
                raw_value="СИНТЕТИКОВ",
                normalized_value="СИНТЕТИКОВ",
                confidence=0.9,
                source_image_id="old",
                original_cyrillic="СИНТЕТИКОВ",
            )
        ],
        "kg_physical_id",
    )
    case.mrz_reports["old"] = TD1Parser().parse([], "old")
    assert case.data.surname_cyrillic
    CaseService.remove_evidence(case, "old")
    assert not case.data.surname_cyrillic and not case.mrz_reports


def test_field_details_highlights_sources_and_repeats_selected_region(physical_window, monkeypatch):
    from PySide6.QtWidgets import QComboBox

    from document_excel.domain.models import FieldCandidate, FieldEvidence

    window = physical_window
    values = [
        FieldCandidate(
            field_name="passport_number",
            raw_value="ID1234567",
            normalized_value="1234567",
            confidence=0.9,
            source_image_id=side.value,
            source_side=side,
            source_method=method,
            source_bbox=(10, 20, 100, 25),
        )
        for side, method in (
            (DocumentSide.FRONT, "VISUAL_OCR"),
            (DocumentSide.BACK, "VISUAL_OCR"),
            (DocumentSide.BACK, "MRZ"),
        )
    ]
    window.current_case.evidence["passport_number"] = FieldEvidence(candidates=values)
    focused, repeated = [], []
    monkeypatch.setattr(window, "focus_candidate", lambda c: focused.append(c))
    monkeypatch.setattr(window.case_actions, "repeat_field", lambda name, c: repeated.append((name, c)))

    def inspect():
        dialog = QApplication.activeModalWidget()
        choices = dialog.findChild(QComboBox)
        assert (
            "Лицевая" in choices.itemText(0)
            and "Оборотная" in choices.itemText(1)
            and "MRZ" in choices.itemText(2)
        )
        choices.setCurrentIndex(1)
        dialog.findChild(QPushButton, "repeatFieldOcr").click()

    QTimer.singleShot(0, inspect)
    window.show_field_details("passport_number")
    assert focused[0] is values[0] and focused[-1] is values[1]
    assert repeated == [("passport_number", values[1])]


def test_repeat_field_reads_original_crop_and_preserves_manual_and_other_client(
    physical_window, image_factory, monkeypatch
):
    import threading

    from document_excel.domain.models import FieldCandidate, FieldEvidence, FieldStatus
    from document_excel.infrastructure.ocr import OcrResult, TesseractEngine, TextBlock

    window = physical_window
    case = window.current_case
    case.data.document_type = "kg_physical_id"
    path = image_factory()
    source = SourceImage(
        path=path,
        role=DocumentSide.BACK,
        processed_path=Path("unavailable-derived.png"),
        processed_confirmed=True,
    )
    case.images = [source]
    c = FieldCandidate(
        field_name="passport_issued_by",
        raw_value="СОМ 123456",
        normalized_value="СОМ 123456",
        confidence=0.9,
        source_image_id=source.id,
        source_bbox=(10, 20, 300, 70),
        source_side=DocumentSide.BACK,
    )
    case.data.passport_issued_by = "РУЧНОЙ ОРГАН"
    case.evidence["passport_issued_by"] = FieldEvidence(
        source="ручной ввод", status=FieldStatus.MANUALLY_CORRECTED, candidates=[c]
    )
    other = ClientCase()
    window.batch.cases.append(other)
    shapes = []

    def recognize(self, pixels, image_id, cancel, field, variant):
        shapes.append(pixels.shape[:2])
        return OcrResult(
            [TextBlock("СЕМ 123456", 0.9, (0, 0, 200, 20), (1, 1, 1))], image_id, variant, 300, 70
        )

    monkeypatch.setattr(TesseractEngine, "recognize_official_region", recognize)
    window.settings.ocr_engine = "tesseract"
    monkeypatch.setattr(
        window, "run_worker", lambda run, done, **_kwargs: done(run(threading.Event(), lambda *_: None))
    )
    monkeypatch.setattr(window, "focus_candidate", lambda *_: None)
    window.case_actions.repeat_field("passport_issued_by", c)
    assert shapes == [(70, 300)] * 5
    assert case.data.passport_issued_by == "РУЧНОЙ ОРГАН" and not other.data.passport_issued_by
    assert all(v.source_bbox == (10, 20, 300, 70) for v in case.evidence["passport_issued_by"].candidates)
    assert source.processed_confirmed


def test_missing_date_has_area_hint_and_repeat_action(physical_window, monkeypatch):
    from PySide6.QtWidgets import QComboBox

    from document_excel.domain.models import FieldEvidence, FieldStatus

    window = physical_window
    e = FieldEvidence(
        status=FieldStatus.NOT_RECOGNIZED, expected_image_id="front", expected_bbox=(30, 40, 200, 50)
    )
    window.current_case.evidence["birth_date"] = e
    repeated = []
    monkeypatch.setattr(window, "focus_candidate", lambda *_: None)
    monkeypatch.setattr(
        window.case_actions, "repeat_field", lambda name, c: repeated.append((name, c.source_bbox))
    )

    def inspect():
        dialog = QApplication.activeModalWidget()
        assert dialog.findChild(QComboBox).count() == 1
        button = dialog.findChild(QPushButton, "repeatFieldOcr")
        assert button.isEnabled()
        button.click()

    QTimer.singleShot(0, inspect)
    window.show_field_details("birth_date")
    assert repeated == [("birth_date", (30, 40, 200, 50))]


def test_diagnostics_mask_values_until_explicit_reveal(physical_window):
    from PySide6.QtWidgets import QTextEdit

    from document_excel.domain.models import FieldCandidate, FieldEvidence

    window = physical_window
    c = FieldCandidate(
        field_name="surname",
        raw_value="PRIVATE_TEST_VALUE",
        normalized_value="PRIVATE_TEST_VALUE",
        confidence=0.9,
        source_image_id="front",
    )
    window.current_case.evidence["surname"] = FieldEvidence(candidates=[c])
    window.case_ocr_texts[window.current_case.id] = {"front": "PRIVATE_TEST_VALUE"}

    def inspect():
        dialog = QApplication.activeModalWidget()
        text = dialog.findChild(QTextEdit).toPlainText()
        assert "PRIVATE_TEST_VALUE" not in text and "candidates" in text
        dialog.reject()

    QTimer.singleShot(0, inspect)
    window.show_raw_ocr()


def test_photo_selection_changes_view_only(physical_window, qtbot, tmp_path):
    window = physical_window
    paths = synthetic_plastic_images(tmp_path / "selection")
    case = window.current_case
    case.images = [
        SourceImage(path=p, role=role)
        for p, role in zip(paths, (DocumentSide.FRONT, DocumentSide.BACK), strict=True)
    ]
    case.data.surname = "СИНТЕТИКОВ"
    case.data.first_name = "ТЕСТ"
    case.data.full_name = "СИНТЕТИКОВ ТЕСТ"
    window.select_case(0)
    before = case.digest()
    for i in (1, 0, 1, 0):
        window.image_list.setCurrentIndex(i)
        qtbot.waitUntil(lambda: not window.background_workers, timeout=10000)
        assert case.digest() == before


def test_late_gui_result_is_ignored_after_photo_change(physical_window, image_factory):
    from document_excel.application.workflows import RecognitionOutput
    from document_excel.domain.models import FieldCandidate

    window = physical_window
    case = window.current_case
    case.images = [SourceImage(id="front", path=image_factory(), role=DocumentSide.FRONT)]
    run = window.case_actions.ocr_state.begin(case)
    case.images[0].rotation = 90
    c = FieldCandidate(
        field_name="surname",
        raw_value="СТАРЫЙ",
        normalized_value="СТАРЫЙ",
        confidence=0.9,
        source_image_id="front",
        source_side=DocumentSide.FRONT,
    )
    window.case_actions.ocr_completed(
        [(case.id, RecognitionOutput([c], "kg_physical_id", [], False, 0), "")], {case.id: run}
    )
    assert not case.data.surname and not case.evidence
    assert case.diagnostic_events[-1]["event"] == "STALE_RESULT_IGNORED"
