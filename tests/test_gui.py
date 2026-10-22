from __future__ import annotations

import threading
import time

import pytest
from PySide6.QtCore import QMimeData, QPointF, Qt, QThreadPool, QUrl
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import QDialog, QMessageBox

from document_excel.domain.models import Address, FieldStatus
from document_excel.infrastructure.excel import WorkbookService
from document_excel.infrastructure.storage import SettingsRepository
from document_excel.presentation.dialogs import AddressDialog, ExportDialog, MappingDialog
from document_excel.presentation.main_window import MainWindow
from document_excel.presentation.widgets import FieldTableModel
from document_excel.presentation.workers import Worker


@pytest.fixture
def window(qtbot, tmp_path, monkeypatch):
    monkeypatch.setattr(MainWindow, "diagnose", lambda self: None)
    result = MainWindow(SettingsRepository(tmp_path / "app-data"))
    result.settings.auto_ocr = False
    qtbot.addWidget(result)
    result.show()
    return result


def test_gui_launch_multiple_images_and_editing(window, image_factory):
    first = image_factory()
    second = image_factory(name="second.png")
    window.case_actions.add_images([first, second])
    assert len(window.current_case.images) == 2
    assert window.image_list.count() == 2
    model = window.field_model
    row = model.fields.index("personal_identifier")
    model.setData(model.index(row, 1), "00123456789012")
    assert window.current_case.data.personal_identifier == "00123456789012"
    assert window.current_case.evidence["personal_identifier"].status == FieldStatus.MANUALLY_CORRECTED


def test_rotate_confirmed_crop_preserves_original_and_other_side(window, qtbot):
    import numpy as np

    from document_excel.domain.models import DocumentSide
    from document_excel.infrastructure.document_images import selected_image

    raw = np.full((800, 1100, 3), 220, np.uint8)
    front = window.capture_controller.photo(DocumentSide.FRONT, raw[100:600, 100:1000], original=raw)
    back = window.capture_controller.photo(DocumentSide.BACK, raw, original=raw)
    window.current_case.images = [front, back]
    window.select_case(0)
    original = front.path.read_bytes()
    window.case_actions.rotate_image()
    current = window.current_case.images[0]
    assert current.processed_confirmed and current.id == front.id
    assert selected_image(current).shape[:2] == (900, 500)
    assert front.path.read_bytes() == original
    assert window.current_case.images[1].id == back.id
    qtbot.waitUntil(lambda: not window.background_workers, timeout=10000)


def test_drag_and_drop_file(window, image_factory, qtbot):
    path = image_factory()
    data = QMimeData()
    data.setUrls([QUrl.fromLocalFile(str(path))])
    enter = QDragEnterEvent(
        window.viewer.rect().center(),
        Qt.DropAction.CopyAction,
        data,
        Qt.MouseButton.LeftButton,
        Qt.KeyboardModifier.NoModifier,
    )
    window.viewer.dragEnterEvent(enter)
    assert enter.isAccepted()
    event = QDropEvent(
        QPointF(50, 50),
        Qt.DropAction.CopyAction,
        data,
        Qt.MouseButton.LeftButton,
        Qt.KeyboardModifier.NoModifier,
    )
    window.viewer.dropEvent(event)
    assert len(window.current_case.images) == 1
    assert event.isAccepted()


def test_folder_drop_does_not_assign_each_photo_new_person(window, image_factory):
    image_factory()
    second = image_factory(name="side2.png")
    window.case_actions.add_images([second.parent])
    assert len(window.batch.cases) == 1
    assert len(window.current_case.images) == 2


def test_registration_required_fields(qtbot):
    dialog = AddressDialog(Address())
    qtbot.addWidget(dialog)
    dialog.accept()
    assert dialog.result() != QDialog.DialogCode.Accepted
    assert dialog.error.text()
    dialog.entries["city"].setText("Тестовый город")
    dialog.entries["street"].setText("ул. Пример")
    dialog.accept()
    assert dialog.result() == QDialog.DialogCode.Accepted
    assert dialog.address().city == "Тестовый город"


def test_absent_address_explicit_action(qtbot):
    dialog = AddressDialog(Address())
    qtbot.addWidget(dialog)
    dialog.absent.setChecked(True)
    dialog.accept()
    assert dialog.result() == QDialog.DialogCode.Accepted
    assert not dialog.address().display()


def test_export_requires_explicit_checkbox(qtbot, template, case, profile):
    dialog = ExportDialog(WorkbookService().plan(template, case, profile))
    qtbot.addWidget(dialog)
    assert not dialog.save_button.isEnabled()
    dialog.accept()
    assert dialog.result() != QDialog.DialogCode.Accepted
    dialog.confirm.setChecked(True)
    assert dialog.save_button.isEnabled()
    dialog.accept()
    assert dialog.result() == QDialog.DialogCode.Accepted


def test_incomplete_export_and_confirmation_blocked(window, monkeypatch):
    errors = []
    monkeypatch.setattr(window, "error", errors.append)
    window.case_actions.confirm_data()
    assert errors
    assert not window.current_case.confirmed
    window.excel_actions.export_current()
    assert errors[-1] == "Выберите Excel-договор."


def test_confirmation_invalidated_by_manual_edit(window, case, monkeypatch):
    window.batch.cases = [case]
    window.select_case(0)
    monkeypatch.setattr(QMessageBox, "question", lambda *args: QMessageBox.StandardButton.Yes)
    window.case_actions.confirm_data()
    assert case.confirmed
    row = window.field_model.fields.index("full_name")
    window.field_model.setData(window.field_model.index(row, 1), "OTHER SYNTHETIC TEST")
    assert not case.confirmed


def test_invalid_date_visible_and_blocks_export(case):
    model = FieldTableModel(case)
    row = model.fields.index("birth_date")
    model.setData(model.index(row, 1), "31.02.2000")
    assert case.data.birth_date is None
    assert case.evidence["birth_date"].status == FieldStatus.INVALID
    assert "31.02.2000" in model.data(model.index(row, 1), Qt.ItemDataRole.ToolTipRole)


def test_mapping_dialog_roundtrip(qtbot, profile):
    dialog = MappingDialog(profile, ["Дог.", "Акт расч"])
    qtbot.addWidget(dialog)
    dialog.accept()
    assert dialog.result() == QDialog.DialogCode.Accepted
    assert dialog.result_profile.field_mappings == profile.field_mappings


def test_worker_does_not_block_gui_and_can_cancel(qtbot):
    started = threading.Event()

    def run(cancel, progress):
        started.set()
        while not cancel.is_set():
            time.sleep(0.01)
        return "should be discarded"

    worker = Worker(run)
    pool = QThreadPool()
    results = []
    errors = []
    worker.signals.result.connect(results.append)
    worker.signals.error.connect(errors.append)
    with qtbot.waitSignal(worker.signals.finished, timeout=3000):
        pool.start(worker)
        qtbot.waitUntil(started.is_set, timeout=1000)
        worker.cancel()
    assert not results
    assert errors and "отменена" in errors[0]
    pool.waitForDone(1000)


def test_worker_error_redacts_raw_exception(qtbot):
    def run(_cancel, _progress):
        raise RuntimeError("SENSITIVE PASSPORT 00123456789012")

    worker = Worker(run)
    errors = []
    worker.signals.error.connect(errors.append)
    worker.run()
    assert errors and "00123456789012" not in errors[0]


def test_error_ui_is_readable(window, monkeypatch):
    messages = []
    monkeypatch.setattr(QMessageBox, "warning", lambda *args: messages.append(args[-1]))
    window.error("Проверьте фотографию")
    assert messages == ["Проверьте фотографию"]
