from __future__ import annotations

import pytest
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QDialog, QFileDialog, QWidget

from document_excel.domain.models import FieldEvidence, FieldStatus
from document_excel.infrastructure.excel import WorkbookService
from document_excel.infrastructure.storage import SettingsRepository
from document_excel.presentation.dialogs import ExportDialog
from document_excel.presentation.main_window import MainWindow


@pytest.fixture
def window(qtbot, tmp_path, monkeypatch):
    monkeypatch.setattr(MainWindow, "diagnose", lambda self: None)
    value = MainWindow(SettingsRepository(tmp_path / "data"))
    value.settings.auto_ocr = False
    errors = []
    monkeypatch.setattr(value, "error", errors.append)
    value.test_errors = errors
    qtbot.addWidget(value)
    value.show()
    yield value
    # Cleanup may follow queued preview I/O on a host also building the EXE.
    # Keep a bounded wait without treating the old ten seconds as an app SLA.
    qtbot.waitUntil(lambda: not value.background_workers, timeout=30000)


def test_controller_dialogs_receive_a_widget_parent(window, monkeypatch, template, case, profile):
    monkeypatch.setattr(QDialog, "exec", lambda dialog: QDialog.DialogCode.Rejected)
    window.case_actions.address_dialog()
    window.case_actions.contract_dialog()
    plan = WorkbookService().plan(template, case, profile)
    window.excel_actions.preview_and_save(plan, case, profile)

    def file_picker(parent, *args):
        assert isinstance(parent, QWidget)
        return ("", "")

    monkeypatch.setattr(
        QFileDialog,
        "getOpenFileNames",
        lambda parent, *args: ([], "") if isinstance(parent, QWidget) else pytest.fail("bad parent"),
    )
    monkeypatch.setattr(QFileDialog, "getOpenFileName", file_picker)
    window.case_actions.choose_images()
    window.excel_actions.choose_template()


def test_gui_background_export_saves_and_reports_result(window, qtbot, monkeypatch, template, case, tmp_path):
    case.template_path = template
    window.batch.cases = [case]
    window.select_case(0)
    target = tmp_path / "gui-result.xlsx"
    monkeypatch.setattr(ExportDialog, "exec", lambda dialog: QDialog.DialogCode.Accepted)
    monkeypatch.setattr(QFileDialog, "getSaveFileName", lambda *args: (str(target), ""))
    window.excel_actions.export_current()
    qtbot.waitUntil(lambda: case.output_path is not None or bool(window.test_errors), timeout=10000)
    assert not window.test_errors
    assert case.output_path == target
    assert target.exists()
    assert window.last_result == target


def test_gui_batch_export_two_clients_separate_files(window, qtbot, monkeypatch, template, case, tmp_path):
    case.template_path = template
    second = case.model_copy(deep=True)
    second.id = "different-case-id"
    second.data.full_name = "SECOND SYNTHETIC CLIENT"
    second.confirmed_digest = second.digest()
    window.batch.cases = [case, second]
    window.select_case(0)
    monkeypatch.setattr(ExportDialog, "exec", lambda dialog: QDialog.DialogCode.Accepted)
    monkeypatch.setattr(QFileDialog, "getExistingDirectory", lambda *args: str(tmp_path))
    window.excel_actions.export_batch()
    qtbot.waitUntil(lambda: second.output_path is not None or bool(window.test_errors), timeout=10000)
    assert not window.test_errors
    assert case.output_path != second.output_path
    assert case.output_path.exists() and second.output_path.exists()


def test_missing_registration_opens_manual_dialog_after_ocr(window, qtbot, monkeypatch):
    from document_excel.application.workflows import RecognitionOutput

    called = []
    monkeypatch.setattr(window.case_actions, "address_dialog", lambda *args, **kwargs: called.append(True))
    result = RecognitionOutput([], "kg_id", [], False, 0.1)
    window.case_actions.ocr_completed([(window.current_case.id, result, "")])
    qtbot.waitUntil(lambda: bool(called), timeout=1000)
    assert not window.current_case.data.registration_address.display()


def test_digital_pending_result_keeps_review_visible_after_finished_signal(
    window, qtbot, tmp_path, monkeypatch
):
    from types import SimpleNamespace

    from test_digital_documents import back, front

    from document_excel.application.workflows import RecognitionOutput
    from document_excel.infrastructure.digital_fields import DigitalDocumentExtractor
    from document_excel.infrastructure.digital_verification import synthetic_digital_images
    from document_excel.infrastructure.document_profiles import KyrgyzDigitalIDProfile
    from document_excel.presentation import case_actions

    paths, _ = synthetic_digital_images(tmp_path / "phones")
    window.case_actions.add_images(paths)
    qtbot.waitUntil(lambda: not window.background_workers, timeout=10000)
    candidates = []
    for image, result in zip(window.current_case.images, [front(), back()], strict=True):
        result.image_id = image.id
        candidates.extend(DigitalDocumentExtractor().extract(result, KyrgyzDigitalIDProfile())[0])
    output = RecognitionOutput(candidates, "kg_digital_id", [], True, 0.1)
    monkeypatch.setattr(
        case_actions, "RecognitionService", lambda *_: SimpleNamespace(run=lambda *_, **__: output)
    )
    window.case_actions.start_ocr()
    qtbot.waitUntil(lambda: window.active_worker is None and not window.background_workers, timeout=10000)
    qtbot.wait(30)
    assert window.current_case.pending_identity and not window.current_case.data.full_name
    assert window.current_step == 2 and window.review_form.isVisible()
    assert window.review_form.fields["full_name"].editor.text() == "Синтетиков Пример Тестович"
    assert window.review_form.fields["surname"].editor.isReadOnly()


@pytest.mark.parametrize("legacy_auto_ocr", [False, True])
def test_import_waits_for_recognize_button_even_with_legacy_auto_ocr(
    window, qtbot, image_factory, monkeypatch, legacy_auto_ocr, template
):
    window.settings.auto_ocr = legacy_auto_ocr
    monkeypatch.setattr(QFileDialog, "getOpenFileName", lambda *args: (str(template), ""))
    window.excel_actions.choose_template()
    qtbot.waitUntil(lambda: not window.background_workers, timeout=10000)
    calls = []
    monkeypatch.setattr(window.case_actions, "start_ocr", lambda *args: calls.append(True))
    for path in (image_factory(name="front.png"), image_factory(lines=["SYNTHETIC BACK"], name="back.png")):
        window.case_actions.add_images([path])
        qtbot.waitUntil(lambda: not window.background_workers, timeout=10000)
        qtbot.wait(250)
        assert not calls and window.active_worker is None
    assert len(window.current_case.images) == 2
    window.primary_button.click()
    assert calls == [True]


def test_legacy_auto_ocr_setting_is_disabled_without_resetting_other_preferences(tmp_path):
    repo = SettingsRepository(tmp_path)
    repo.path.write_text('{"auto_ocr":true,"theme":"dark","paddle_cpu_threads":1}', encoding="utf8")
    settings = repo.load()
    assert settings.auto_ocr is False and settings.theme == "dark" and settings.paddle_cpu_threads == 1


@pytest.mark.ocr
def test_real_ocr_updates_gui_without_losing_manual_input(window, qtbot, image_factory):
    from document_excel.infrastructure.ocr import TesseractEngine

    if TesseractEngine(window.settings).diagnose().error:
        pytest.skip("Tesseract and rus+eng are required")
    window.settings.ocr_engine = "tesseract"
    window.settings.languages = ["eng"]
    window.settings.preprocessing = "original"
    window.case_actions.add_images([image_factory()])
    window.current_case.data.notes = "MANUAL NOTE"
    window.current_case.evidence["notes"] = FieldEvidence(
        source="ручной ввод", status=FieldStatus.MANUALLY_CORRECTED
    )
    window.case_actions.start_ocr()
    ticks = []
    timer = QTimer()
    timer.timeout.connect(lambda: ticks.append(True))
    timer.start(10)
    qtbot.waitUntil(lambda: window.active_worker is None or bool(window.test_errors), timeout=15000)
    timer.stop()
    assert not window.test_errors
    assert ticks
    assert window.current_case.data.personal_identifier == "00123456789012"
    assert window.current_case.data.notes == "MANUAL NOTE"
    assert not window.current_case.confirmed


def test_cancel_new_profile_preserves_current_profile(window, qtbot, monkeypatch, template):
    window.current_case.template_path = template
    original = window.profile
    monkeypatch.setattr(QDialog, "exec", lambda dialog: QDialog.DialogCode.Rejected)
    window.excel_actions.new_mapping()
    qtbot.waitUntil(lambda: window.active_worker is None, timeout=5000)
    assert window.profile is original


def test_confirm_shortcut_commits_active_table_editor(window, qtbot, case, monkeypatch):
    from PySide6.QtWidgets import QLineEdit, QMessageBox

    window.batch.cases = [case]
    window.select_case(0)
    row = window.field_model.fields.index("full_name")
    index = window.field_model.index(row, 1)
    window.fields.setCurrentIndex(index)
    window.fields.edit(index)
    editor = window.fields.findChild(QLineEdit)
    assert editor is not None
    editor.setFocus()
    editor.setText("EDITED SYNTHETIC NAME")
    monkeypatch.setattr(QMessageBox, "question", lambda *args: QMessageBox.StandardButton.Yes)
    window.case_actions.confirm_data()
    assert case.data.full_name == "EDITED SYNTHETIC NAME"
    assert case.confirmed


def test_closing_does_not_open_cancellation_error_dialog(window, monkeypatch):
    from PySide6.QtWidgets import QMessageBox

    monkeypatch.setattr(QMessageBox, "warning", lambda *args: pytest.fail("Modal dialog during shutdown"))
    window.closing = True
    MainWindow.error(window, "Отмена при закрытии")
