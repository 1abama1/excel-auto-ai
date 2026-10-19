from __future__ import annotations

import pytest
from PySide6.QtCore import QPoint, Qt
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import QLabel, QMessageBox, QPushButton

from document_excel.application.workflows import RecognitionOutput
from document_excel.domain.models import (
    ClientCase,
    FieldCandidate,
    FieldEvidence,
    FieldStatus,
    RegistrationStatus,
    SourceImage,
)
from document_excel.infrastructure.storage import SettingsRepository
from document_excel.presentation.components import StepIndicator
from document_excel.presentation.main_window import MainWindow
from document_excel.presentation.theme import apply_theme


@pytest.fixture
def redesigned_window(qtbot, qapp, tmp_path, monkeypatch):
    monkeypatch.setattr(MainWindow, "diagnose", lambda self: None)
    window = MainWindow(SettingsRepository(tmp_path / "redesign-data"))
    window.settings.auto_ocr = False
    apply_theme(qapp, window.settings)
    errors = []
    monkeypatch.setattr(window, "error", errors.append)
    window.test_errors = errors
    qtbot.addWidget(
        window,
        before_close_func=lambda _: qtbot.waitUntil(lambda: not window.background_workers, timeout=10000),
    )
    window.show()
    qtbot.waitUntil(window.isVisible)
    return window


def test_empty_workflow_has_one_primary_and_hides_technical_table(redesigned_window):
    window = redesigned_window
    assert window.current_step == 0
    assert window.step_stack.count() == 5
    assert window.step_stack.currentIndex() == 0
    assert not window.fields.isVisible()
    visible_primary = [
        button
        for button in window.findChildren(QPushButton)
        if button.isVisible() and button.objectName() == "primary"
    ]
    assert visible_primary == [window.primary_button]
    technical = ("tesseract", "timeout", "кандидатов:", "модели:")
    assert not any(
        token in label.text().lower()
        for label in window.findChildren(QLabel)
        if label.isVisible()
        for token in technical
    )


def test_two_sides_stay_in_one_client_and_can_be_reviewed(redesigned_window, image_factory):
    window = redesigned_window
    window.case_actions.add_images([image_factory(), image_factory(name="synthetic-back.png")])
    assert len(window.batch.cases) == 1
    assert len(window.current_case.images) == 2
    assert window.image_list.count() == 2
    assert window.primary_button.isEnabled()
    window.go_step(2)
    assert window.step_stack.currentIndex() == 2
    assert window.viewer.isVisible()
    assert window.review_form.isVisible()


def test_manual_review_edit_supports_undo_redo_and_invalidates_confirmation(redesigned_window, case):
    window = redesigned_window
    window.batch.cases = [case]
    window.select_case(0)
    window.go_step(2)
    original = case.data.personal_identifier
    window.apply_field_edit("personal_identifier", "00987654321098")
    assert case.data.personal_identifier == "00987654321098"
    assert case.evidence["personal_identifier"].status == FieldStatus.MANUALLY_CORRECTED
    assert not case.confirmed
    assert window.undo_stack.canUndo()
    window.undo_stack.undo()
    assert case.data.personal_identifier == original
    assert window.review_form.fields["personal_identifier"].editor.text() == original
    window.undo_stack.redo()
    assert case.data.personal_identifier == "00987654321098"


def test_invalid_date_stays_visible_for_correction(redesigned_window, case):
    window = redesigned_window
    window.batch.cases = [case]
    window.select_case(0)
    window.go_step(2)
    window.apply_field_edit("birth_date", "31.02.2000")
    field = window.review_form.fields["birth_date"]
    assert case.data.birth_date is None
    assert case.evidence["birth_date"].status == FieldStatus.INVALID
    assert field.editor.text() == "31.02.2000"
    assert field.status_label.text() == "Ошибка"
    assert field.validation.text()
    window.go_step(0)
    window.go_step(2)
    assert field.editor.text() == "31.02.2000"


def test_back_navigation_commits_pending_visible_editor(redesigned_window, case, qtbot):
    window = redesigned_window
    window.batch.cases = [case]
    window.select_case(0)
    window.go_step(2)
    editor = window.review_form.fields["full_name"].editor
    editor.setFocus()
    qtbot.waitUntil(editor.hasFocus)
    editor.setText("SYNTHETIC NAVIGATION EDIT")
    qtbot.mouseClick(window.back_button, Qt.MouseButton.LeftButton)
    assert window.current_step == 0
    assert case.data.full_name == "SYNTHETIC NAVIGATION EDIT"
    window.go_step(2)
    assert editor.text() == "SYNTHETIC NAVIGATION EDIT"


def test_returning_to_upload_keeps_review_reachable(redesigned_window, case):
    window = redesigned_window
    window.batch.cases = [case]
    window.select_case(0)
    window.go_step(2)
    window.go_step(0)
    steps = window.findChild(StepIndicator)
    assert steps.buttons[2].isEnabled()
    steps.buttons[2].click()
    assert window.current_step == 2


def test_confirmation_commits_visible_editor(redesigned_window, case, qtbot, monkeypatch):
    window = redesigned_window
    window.batch.cases = [case]
    window.select_case(0)
    window.go_step(2)
    editor = window.review_form.fields["full_name"].editor
    editor.setFocus()
    qtbot.waitUntil(editor.hasFocus)
    editor.setText("SYNTHETIC CONFIRMED EDIT")
    monkeypatch.setattr(QMessageBox, "question", lambda *args: QMessageBox.StandardButton.Yes)
    window.case_actions.confirm_data()
    assert case.data.full_name == "SYNTHETIC CONFIRMED EDIT"
    assert case.confirmed
    assert not window.test_errors


def test_switching_clients_commits_edit_without_copying_it(redesigned_window, case, qtbot):
    window = redesigned_window
    second = ClientCase(name="Synthetic second client")
    window.batch.cases = [case, second]
    window.select_case(0)
    window.go_step(2)
    editor = window.review_form.fields["full_name"].editor
    editor.setFocus()
    qtbot.waitUntil(editor.hasFocus)
    editor.setText("SYNTHETIC FIRST CLIENT EDIT")
    window.select_case(1)
    assert case.data.full_name == "SYNTHETIC FIRST CLIENT EDIT"
    assert second.data.full_name is None
    assert window.review_form.fields["full_name"].editor.text() == ""
    window.select_case(0)
    assert window.review_form.fields["full_name"].editor.text() == "SYNTHETIC FIRST CLIENT EDIT"


def test_field_location_selects_its_source_image(redesigned_window, image_factory, qtbot, monkeypatch):
    window = redesigned_window
    window.case_actions.add_images([image_factory(), image_factory(name="evidence-back.png")])
    case = window.current_case
    candidate = FieldCandidate(
        field_name="passport_number",
        raw_value="AA0000001",
        normalized_value="AA0000001",
        confidence=0.95,
        source_image_id=case.images[1].id,
        source_bbox=(100, 100, 300, 45),
    )
    case.data.passport_number = candidate.normalized_value
    case.evidence["passport_number"] = FieldEvidence(
        source="OCR",
        status=FieldStatus.VALID,
        candidates=[candidate],
    )
    window.refresh_form()
    window.go_step(2)
    highlighted = []
    monkeypatch.setattr(window.viewer, "highlight", highlighted.append)
    field = window.review_form.fields["passport_number"]
    window.review_form.reveal("passport_number")
    qtbot.mouseClick(field.locate, Qt.MouseButton.LeftButton)
    assert window.image_list.currentIndex() == 1
    assert highlighted and highlighted[-1] == candidate.source_bbox


def test_escape_cancels_region_selection_and_restores_panning(redesigned_window, image_factory, qtbot):
    window = redesigned_window
    window.case_actions.add_images([image_factory()])
    qtbot.waitUntil(lambda: not window.background_workers, timeout=5000)
    window.go_step(2)
    window.viewer.select_region()
    assert window.viewer.selecting
    assert window.viewer.dragMode() == window.viewer.DragMode.NoDrag
    window.activateWindow()
    window.viewer.setFocus()
    qtbot.keyClick(window.viewer, Qt.Key.Key_Escape)
    qtbot.waitUntil(lambda: not window.viewer.selecting)
    assert window.viewer.dragMode() == window.viewer.DragMode.ScrollHandDrag
    assert window.viewer.roi_start is None


def test_other_clients_ocr_never_opens_current_clients_registration_dialog(
    redesigned_window, qtbot, monkeypatch
):
    window = redesigned_window
    second = ClientCase(name="Synthetic background client")
    window.batch.cases.append(second)
    called = []
    monkeypatch.setattr(window.case_actions, "address_dialog", lambda *a, **kw: called.append(True))
    output = RecognitionOutput([], "kg_id", [], False, 0.1)
    window.case_actions.ocr_completed([(second.id, output, "")])
    qtbot.wait(30)
    assert not called
    assert window.current_case is not second
    assert window.current_step == 0


def test_explicit_absent_registration_is_preserved_after_ocr(redesigned_window, qtbot, monkeypatch):
    window = redesigned_window
    case = window.current_case
    case.registration_status = RegistrationStatus.NOT_PRESENT
    case.evidence["registration_address"] = FieldEvidence(
        source="ручной ввод",
        status=FieldStatus.MANUALLY_CORRECTED,
        message="Отсутствует в документе",
    )
    called = []
    monkeypatch.setattr(window.case_actions, "address_dialog", lambda *a, **kw: called.append(True))
    window.case_actions.ocr_completed([(case.id, RecognitionOutput([], "kg_id", [], False, 0.1), "")])
    qtbot.wait(30)
    assert not called
    assert case.registration_status == RegistrationStatus.NOT_PRESENT


def test_batch_ocr_of_another_client_restores_current_step(
    redesigned_window, image_factory, qtbot, monkeypatch
):
    window = redesigned_window
    second = ClientCase(name="Synthetic batch client", images=[SourceImage(path=image_factory())])
    window.batch.cases.append(second)

    class SyntheticRecognition:
        def __init__(self, settings):
            pass

        def run(self, *args):
            return RecognitionOutput([], "kg_id", [], False, 0.1)

    monkeypatch.setattr("document_excel.presentation.case_actions.RecognitionService", SyntheticRecognition)
    window.case_actions.start_ocr(all_cases=True)
    qtbot.waitUntil(lambda: window.active_worker is None, timeout=3000)
    qtbot.wait(30)
    assert window.current_case is not second
    assert window.current_step == 0
    assert second.status == "На проверке"


def test_queue_single_click_stays_in_queue_and_double_click_opens_client(redesigned_window, case, qtbot):
    window = redesigned_window
    second = ClientCase(name="Synthetic queue client")
    window.batch.cases = [case, second]
    window.select_case(0)
    window.refresh_cases()
    window.navigate(3)
    table = window.queue_table
    position = table.visualItemRect(table.item(1, 0)).center()
    qtbot.mouseClick(table.viewport(), Qt.MouseButton.LeftButton, pos=position)
    assert window.current_case is second
    assert window.pages.currentIndex() == 3
    position = table.visualItemRect(table.item(1, 0)).center()
    qtbot.mouseDClick(table.viewport(), Qt.MouseButton.LeftButton, pos=position)
    assert window.pages.currentIndex() == 0
    assert window.current_case is second


def test_registration_dialog_edit_can_be_undone(redesigned_window, monkeypatch):
    from document_excel.presentation.dialogs import AddressDialog

    window = redesigned_window

    def fill(dialog):
        dialog.entries["city"].setText("Тестовый город")
        dialog.entries["street"].setText("ул. Пример")
        dialog.accept()
        return dialog.result()

    monkeypatch.setattr(AddressDialog, "exec", fill)
    window.case_actions.address_dialog()
    assert window.current_case.data.registration_address.city == "Тестовый город"
    assert window.undo_stack.canUndo()
    window.undo_stack.undo()
    assert not window.current_case.data.registration_address.display()
    window.undo_stack.redo()
    assert window.current_case.data.registration_address.city == "Тестовый город"


def test_clear_workspace_removes_raw_edits_and_ocr_text_from_memory(redesigned_window, monkeypatch):
    window = redesigned_window
    previous_id = window.current_case.id
    window.apply_field_edit("birth_date", "31.02.2000")
    window.case_ocr_texts[previous_id] = {"synthetic-image": "SYNTHETIC OCR TEXT"}
    assert window.case_raw_edits[previous_id]
    monkeypatch.setattr(QMessageBox, "question", lambda *a: QMessageBox.StandardButton.Yes)
    window.case_actions.clear_workspace()
    assert len(window.batch.cases) == 1
    assert window.current_case.id != previous_id
    assert previous_id not in window.case_raw_edits
    assert not window.case_ocr_texts
    assert not window.undo_stack.canUndo()
    assert window.review_form.fields["birth_date"].editor.text() == ""
    assert window.current_step == 0


@pytest.mark.parametrize("action", ["open_result", "open_result_folder"])
def test_unsaved_client_never_opens_another_clients_result(redesigned_window, tmp_path, monkeypatch, action):
    window = redesigned_window
    previous_result = tmp_path / "other-synthetic-client.xlsx"
    previous_result.write_bytes(b"synthetic result")
    window.last_result = previous_result
    messages = []
    opened = []
    monkeypatch.setattr(window, "message", messages.append)
    monkeypatch.setattr(QDesktopServices, "openUrl", lambda url: opened.append(url) or True)
    getattr(window.excel_actions, action)()
    assert not opened
    assert messages


def test_saved_result_is_scoped_to_client_and_new_case_returns_to_upload(redesigned_window, case, tmp_path):
    window = redesigned_window
    result = tmp_path / "synthetic-contract.xlsx"
    result.write_bytes(b"synthetic result")
    case.output_path = result
    window.batch.cases = [case]
    window.select_case(0)
    window.go_step(4)
    assert window.current_step == 4
    assert any(result.name in label.text() for label in window.findChildren(QLabel) if label.isVisible())
    window.case_actions.new_case()
    assert window.current_case.output_path is None
    assert window.current_step == 0
    assert not any(result.name in label.text() for label in window.findChildren(QLabel) if label.isVisible())


@pytest.mark.parametrize("size", [(1366, 768), (1600, 900), (1920, 1080)])
@pytest.mark.parametrize("step", [0, 2, 3])
def test_primary_actions_remain_inside_supported_desktop_sizes(redesigned_window, qtbot, size, step):
    window = redesigned_window
    window.resize(*size)
    window.go_step(step)
    qtbot.wait(30)
    assert window.width() == size[0]
    assert window.height() == size[1]
    for button in [window.primary_button, window.back_button]:
        if not button.isVisible():
            continue
        top_left = button.mapTo(window, QPoint(0, 0))
        bottom_right = button.mapTo(window, button.rect().bottomRight())
        assert window.rect().contains(top_left), (size, step, button.text(), top_left)
        assert window.rect().contains(bottom_right), (size, step, button.text(), bottom_right)
        assert button.height() >= 36
    assert window.primary_button.isVisible()
