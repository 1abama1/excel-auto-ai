from datetime import date
from pathlib import Path
from types import SimpleNamespace

import pytest
from PySide6.QtWidgets import QDialog, QFileDialog

from document_excel.application.workflows import RecognitionOutput
from document_excel.domain.models import DocumentSide, FieldCandidate, FieldStatus, WorkflowMode
from document_excel.infrastructure.excel import create_demo_template, file_hash
from document_excel.infrastructure.storage import SettingsRepository
from document_excel.presentation import case_actions
from document_excel.presentation.dialogs import CopySaveDialog
from document_excel.presentation.main_window import MainWindow


@pytest.fixture
def simple_window(qtbot, tmp_path, monkeypatch):
    monkeypatch.setattr(MainWindow, "diagnose", lambda self: None)
    window = MainWindow(SettingsRepository(tmp_path / "settings"))
    qtbot.addWidget(
        window,
        before_close_func=lambda _: qtbot.waitUntil(lambda: not window.background_workers, timeout=15000),
    )
    window.show()
    window.excel_actions.filenames.clock = lambda: date(2026, 10, 20)
    source = tmp_path / "contract.xlsx"
    create_demo_template(source)
    monkeypatch.setattr(QFileDialog, "getOpenFileName", lambda *args: (str(source), ""))
    errors = []
    monkeypatch.setattr(window, "error", errors.append)
    window.test_errors = errors
    window.contract_button.click()
    qtbot.waitUntil(lambda: not window.background_workers, timeout=15000)
    return window


def test_one_photo_ocr_review_and_single_summary_save(
    simple_window, image_factory, qtbot, monkeypatch, tmp_path
):
    window = simple_window
    assert window.current_case.workflow_mode == WorkflowMode.BEST_EFFORT_UPDATE
    assert window.sidebar.isHidden()
    window.case_actions.add_images([image_factory()])
    qtbot.waitUntil(lambda: not window.background_workers, timeout=15000)
    image = window.current_case.images[0]
    image.role = DocumentSide.FRONT
    candidate = FieldCandidate(
        field_name="full_name",
        raw_value="ИВАНОВ ИВАН",
        normalized_value="ИВАНОВ ИВАН",
        source_image_id=image.id,
        confidence=0.9,
        validation_status=FieldStatus.VALID,
    )
    output = RecognitionOutput([candidate], "kg_id", [], False, 0.1, sides={image.id: DocumentSide.FRONT})
    calls = []
    monkeypatch.setattr(
        case_actions,
        "RecognitionService",
        lambda *_: SimpleNamespace(run=lambda *a, **kw: (calls.append(1), output)[1]),
    )
    monkeypatch.setattr(
        window.case_actions, "address_dialog", lambda *a, **kw: pytest.fail("mandatory address dialog")
    )
    window.refresh_form()
    assert window.can_recognize()
    window.primary_button.click()
    qtbot.waitUntil(lambda: not window.background_workers, timeout=15000)
    assert window.current_step == 2 and window.primary_button.text() == "Сохранить договор"
    assert "Не найдено" in window.review_summary.text()
    editor = window.review_form.fields["registration_address"].editor
    assert not editor.isReadOnly()
    editor.setFocus()
    qtbot.keyClicks(editor, "Test address")
    assert window.current_case.data.registration_address.display() == "Test address"
    assert calls == [1]  # switching photos and ordinary typing cannot call OCR
    dialogs = []

    def accept(dialog):
        dialogs.append(dialog)
        assert dialog.in_place and dialog.filename.isHidden() and dialog.folder.isHidden()
        assert dialog.save_button.isEnabled()
        dialog.accept()
        return QDialog.DialogCode.Accepted

    monkeypatch.setattr(CopySaveDialog, "exec", accept)
    monkeypatch.setattr(QFileDialog, "getSaveFileName", lambda *a: pytest.fail("extra save dialog"))
    window.primary_button.click()
    qtbot.waitUntil(
        lambda: window.current_case.output_path is not None or bool(window.test_errors), timeout=15000
    )
    qtbot.waitUntil(lambda: not window.background_workers, timeout=15000)
    qtbot.waitUntil(lambda: window.current_step == 4, timeout=5000)
    assert not window.test_errors and len(dialogs) == 1 and calls == [1]
    assert window.current_case.output_path == tmp_path / "Иванов Иван 20.10.2026.xlsx"
    assert window.current_case.template_path == window.current_case.output_path
    assert sorted(p.name for p in tmp_path.glob("*.xlsx")) == ["Иванов Иван 20.10.2026.xlsx"]
    assert window.result_label.text() == "Иванов Иван 20.10.2026.xlsx"
    assert "Обновлено" in window.result_summary.text()
    window.primary_button.click()
    assert window.current_step == 0 and not window.current_case.images
    assert not window.current_case.template_path and not window.current_case.data.full_name


def test_empty_update_keeps_selected_name_and_three_visible_steps(
    simple_window, qtbot, tmp_path, monkeypatch
):
    window = simple_window
    assert window.skip_ocr_button.isEnabled()
    window.skip_ocr_button.click()
    assert window.current_step == 2 and window.primary_button.isEnabled()
    assert sum(b.isVisible() for b in window.step_indicator.buttons) == 3
    (tmp_path / "Custom.xlsx").write_bytes(b"keep existing")
    before = file_hash(window.current_case.template_path)

    def accept(dialog):
        assert dialog.in_place and dialog.filename.text() == "Договор 20.10.2026.xlsx"
        assert dialog.save_button.isEnabled()
        dialog.accept()
        return QDialog.DialogCode.Accepted

    monkeypatch.setattr(CopySaveDialog, "exec", accept)
    window.primary_button.click()
    qtbot.waitUntil(
        lambda: window.current_case.output_path is not None or bool(window.test_errors), timeout=15000
    )
    qtbot.waitUntil(lambda: not window.background_workers, timeout=15000)
    assert not window.test_errors
    assert window.current_case.output_path == tmp_path / "Договор 20.10.2026.xlsx"
    assert file_hash(window.current_case.output_path) == before
    assert len(list(tmp_path.glob("*.xlsx"))) == 2
    assert (tmp_path / "Custom.xlsx").read_bytes() == b"keep existing"


def test_xlsx_drop_uses_safe_loader(simple_window, qtbot, tmp_path):
    source = tmp_path / "drop.xlsx"
    create_demo_template(source)
    simple_window.handle_dropped_files([source])
    qtbot.waitUntil(lambda: not simple_window.background_workers, timeout=15000)
    assert simple_window.current_case.template_path == Path(source)
    assert simple_window.current_case.workbook_profile["field_mappings"]


def test_component_edit_updates_name_after_manual_full_name(simple_window):
    window = simple_window
    window.apply_field_edit("full_name", "Старое Имя")
    window.apply_field_edit("surname", "Новиков")
    window.apply_field_edit("first_name", "Иван")
    assert window.current_case.formatted_name == "Новиков Иван"
