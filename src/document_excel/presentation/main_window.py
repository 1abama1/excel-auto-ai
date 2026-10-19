from __future__ import annotations

import logging
import tempfile
from importlib.resources import files
from pathlib import Path

import cv2
from PySide6.QtCore import QObject, Qt, QThreadPool, QTimer
from PySide6.QtGui import QAction, QIcon, QKeySequence, QUndoCommand, QUndoStack
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QLineEdit,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
)

from document_excel.application.workflows import CaseService
from document_excel.domain.errors import AppError
from document_excel.domain.models import (
    FIELD_LABELS,
    ClientCase,
    DocumentBatch,
    DocumentSide,
    FieldEvidence,
    FieldStatus,
    RegistrationStatus,
    data_value,
)
from document_excel.infrastructure.document_images import DocumentCaptureController
from document_excel.infrastructure.excel import TemplateRegistry, WorkbookService, legacy_profile, readiness
from document_excel.infrastructure.ocr_factory import create_engine
from document_excel.infrastructure.preprocessing import detect_document_region, load_image
from document_excel.infrastructure.storage import HistoryRepository, SettingsRepository
from document_excel.presentation.case_actions import CaseActions
from document_excel.presentation.components import PhotoThumbnail
from document_excel.presentation.dialogs import SettingsDialog
from document_excel.presentation.excel_actions import ExcelActions
from document_excel.presentation.theme import apply_theme
from document_excel.presentation.widgets import FieldTableModel
from document_excel.presentation.workers import Worker


class EditCommand(QUndoCommand):
    def __init__(self, window, before, after, label):
        super().__init__(label)
        self.window, self.before, self.after = window, before, after

    def undo(self):
        self.window.restore_edit(self.before)

    def redo(self):
        self.window.restore_edit(self.after)


class WorkerCallbacks(QObject):
    """Qt receiver lifetime follows the window, including queued callbacks."""

    def __init__(self, window, worker, callback):
        super().__init__(window)
        self.window, self.worker, self.callback = window, worker, callback

    def receive(self, value):
        self.callback(value)

    def finish(self):
        self.window.worker_finished(self.worker)
        self.deleteLater()


class MainWindow(QMainWindow):
    def __init__(self, repository: SettingsRepository | None = None) -> None:
        super().__init__()
        self.repository = repository or SettingsRepository()
        self.settings = self.repository.load()
        self.history = HistoryRepository(self.repository.root)
        self.registry = TemplateRegistry(self.repository.root / "profiles")
        self.workbooks, self.case_service = WorkbookService(), CaseService()
        self.capture_controller = DocumentCaptureController(self.case_service)
        self.profile = legacy_profile()
        self.batch = DocumentBatch(cases=[ClientCase(name="Клиент 1")])
        self.current_case = self.batch.cases[0]
        self.active_worker: Worker | None = None
        self.background_workers: list[Worker] = []
        self.pool = QThreadPool(self)
        self.pool.setMaxThreadCount(self.settings.max_parallel)
        self.closing = False
        self.last_result: Path | None = None
        self.current_step = 0
        self.case_steps: dict[str, int] = {}
        self.case_ocr_texts: dict[str, dict[str, str]] = {}
        self.template_sheets: dict[Path, list[str]] = {}
        self.case_raw_edits: dict[str, dict[str, str]] = {}
        self.preview_token = 0
        self.pending_highlight = None
        self.ocr_diagnostic = None
        self._committing = False
        self.undo_stack = QUndoStack(self)
        self.setWindowTitle("Документ → Excel")
        self.setWindowIcon(QIcon(str(files("document_excel").joinpath("resources/app.ico"))))
        self.resize(self.settings.window_width, self.settings.window_height)
        self.setMinimumSize(800, 560)
        self.case_actions, self.excel_actions = CaseActions(self), ExcelActions(self)
        from document_excel.presentation.workflow_view import build_ui

        build_ui(self)
        self.load_defaults()
        self.select_case(0)
        self.refresh_cases()
        self.install_shortcuts()
        if self.repository.warning:
            self.notice.setText(self.repository.warning)
        QTimer.singleShot(100, self, self.diagnose)

    def make_button(self, title, callback, primary=False):
        button = QPushButton(title)
        button.setMinimumHeight(40)
        button.setObjectName("primary" if primary else "secondary")
        button.clicked.connect(callback)
        return button

    def install_shortcuts(self):
        for shortcut, callback in [
            ("Ctrl+O", self.case_actions.choose_images),
            ("Ctrl+N", self.case_actions.new_case),
            ("Ctrl+Return", self.case_actions.confirm_data),
            ("Ctrl+E", self.excel_actions.export_current),
            ("Ctrl+S", self.excel_actions.export_current),
            ("Ctrl+Z", lambda: self.undo_stack.undo() if not self.active_worker else None),
            ("Ctrl+Shift+Z", lambda: self.undo_stack.redo() if not self.active_worker else None),
            ("Ctrl++", lambda: self.viewer.zoom(1.2)),
            ("Ctrl+=", lambda: self.viewer.zoom(1.2)),
            ("Ctrl+-", lambda: self.viewer.zoom(1 / 1.2)),
            ("Ctrl+0", self.viewer.fit),
            ("Escape", self.cancel_work),
        ]:
            action = QAction(self)
            action.setShortcut(QKeySequence(shortcut))
            action.triggered.connect(callback)
            self.addAction(action)

    def load_defaults(self):
        if self.settings.default_profile and Path(self.settings.default_profile).exists():
            try:
                self.profile = self.registry.load(Path(self.settings.default_profile))
            except AppError as exc:
                self.notice.setText(str(exc))
        if self.settings.default_template and Path(self.settings.default_template).exists():
            path = Path(self.settings.default_template)
            self.current_case.template_path = path

            def loaded(description):
                sheets, profile = description
                self.template_sheets[path] = sheets
                if profile != self.profile:
                    self.excel_actions.use_profile(profile)
                self.refresh_form()

            self.run_worker(lambda _c, _p: self.workbooks.describe(path, self.profile), loaded, busy=False)

    def navigate(self, index):
        self.commit_edits()
        if index == 2:
            self.show_history()
        else:
            self.pages.setCurrentIndex(index)
        if index == 3:
            self.refresh_cases()
        for key, button in self.nav_buttons.items():
            button.setChecked(key == index)

    def go_step(self, index):
        if self.active_worker and index != 1:
            return
        self.commit_edits()
        self.current_step = index
        if index != 0 and self.notice.text().startswith(("Фото добавлены", "Фотографии добавлены")):
            self.notice.setText("")
        if index >= 3 and self.notice.text().startswith("Данные извлечены"):
            self.notice.setText("")
        self.case_steps[self.current_case.id] = max(index, self.case_steps.get(self.current_case.id, 0))
        self.pages.setCurrentIndex(0)
        self.step_stack.setCurrentIndex(index)
        self.step_indicator.update_step(index)
        reached = self.case_steps[self.current_case.id]
        for i, button in enumerate(self.step_indicator.buttons):
            button.setEnabled(not self.active_worker and i <= reached and i != 1)
        self.title_label.setText(
            [
                "Добавьте документы",
                "Распознаём документ…",
                "Проверьте данные",
                "Подготовьте Excel",
                "Договор готов",
            ][index]
        )
        self.subtitle_label.setText(
            [
                "Две стороны удостоверения — одна запись клиента",
                "Обработка выполняется на этом компьютере",
                "Сверьте значения с фотографиями и исправьте выделенные поля",
                "Выберите шаблон и проверьте, куда попадут данные",
                "Файл проверен после записи и сохранён в новой копии",
            ][index]
        )
        self.back_button.setVisible(index in (2, 3))
        self.primary_button.setText(
            [
                "Распознать документы",
                "Отменить распознавание",
                "Проверить и продолжить",
                "Сохранить Excel",
                "Новый клиент",
            ][index]
        )
        self.primary_button.setEnabled(
            not self.active_worker and (index != 0 or bool(self.current_case.images))
        )
        if index == 1:
            self.primary_button.setEnabled(bool(self.active_worker))
        self.footer_note.setText(f"{self.current_case.name} · {len(self.current_case.images)} фото")
        if index == 4:
            result = self.current_case.output_path
            self.result_label.setText(str(result) if result else "Excel ещё не сохранён")
            self.open_result_button.setEnabled(bool(result))
            self.open_folder_button.setEnabled(bool(result))
        QTimer.singleShot(0, self.viewer, self.viewer.fit)

    def primary_action(self):
        self.commit_edits()
        actions = [
            self.case_actions.start_ocr,
            self.cancel_work,
            self.case_actions.confirm_data,
            self.excel_actions.export_current,
            self.case_actions.new_case,
        ]
        actions[self.current_step]()

    def snapshot_edit(self):
        c = self.current_case
        return (
            c.id,
            c.data.model_copy(deep=True),
            {k: v.model_copy(deep=True) for k, v in c.evidence.items()},
            c.registration_status,
            c.contract.model_copy(deep=True),
            dict(self.field_model.raw_edits),
            c.document_type_manual,
        )

    def push_snapshot(self, before, label="Изменить данные"):
        after = self.snapshot_edit()
        if before != after:
            self.undo_stack.push(EditCommand(self, before, after, label))

    def restore_edit(self, state):
        case_id, data, evidence, registration, contract, raw, manual_type = state
        if self.current_case.id != case_id:
            return
        self.current_case.data = data.model_copy(deep=True)
        self.current_case.evidence = {k: v.model_copy(deep=True) for k, v in evidence.items()}
        self.current_case.registration_status = registration
        self.current_case.contract = contract.model_copy(deep=True)
        self.current_case.document_type_manual = manual_type
        self.field_model.raw_edits = dict(raw)
        self.current_case.invalidate()
        self.refresh_form()
        self.refresh_cases()

    def apply_field_edit(self, field, value):
        if self.active_worker or field in {"registration_address", "living_address"}:
            return
        row = self.field_model.fields.index(field)
        if str(self.field_model.data(self.field_model.index(row, 1)) or "") == value.strip():
            return
        before = self.snapshot_edit()
        self.field_model.setData(self.field_model.index(row, 1), value)
        self.push_snapshot(before, "Изменить: " + FIELD_LABELS[field])

    def commit_edits(self):
        if (
            self._committing
            or not hasattr(self, "field_model")
            or self.field_model.case is not self.current_case
        ):
            return
        self._committing = True
        try:
            editor = self.fields.findChild(QLineEdit)
            index = self.fields.currentIndex()
            if editor and index.isValid():
                if str(self.field_model.data(index) or "") != editor.text().strip():
                    self.field_model.setData(index, editor.text())
                self.fields.closePersistentEditor(index)
            for name, field in self.review_form.fields.items():
                if field.editor.hasFocus():
                    self.apply_field_edit(name, field.editor.text())
                    break
            self.case_raw_edits[self.current_case.id] = dict(self.field_model.raw_edits)
        finally:
            self._committing = False

    def refresh_cases(self):
        self.cases.blockSignals(True)
        self.queue_table.blockSignals(True)
        self.cases.clear()
        self.queue_table.setRowCount(len(self.batch.cases))
        for row, case in enumerate(self.batch.cases):
            self.cases.addItem(QListWidgetItem(f"{case.name}\n{len(case.images)} фото · {case.status}"))
            problems = readiness(case, self.profile, require_confirmation=False)
            values = [
                case.name,
                str(len(case.images)),
                case.status,
                str(len(problems)),
                "Готов" if case.confirmed and not problems else "Нужна проверка",
            ]
            for col, value in enumerate(values):
                self.queue_table.setItem(row, col, QTableWidgetItem(value))
            if case.id == self.current_case.id:
                self.cases.setCurrentRow(row)
                self.queue_table.selectRow(row)
        self.cases.blockSignals(False)
        self.queue_table.blockSignals(False)
        self.queue_badge.setText(f"Клиентов в очереди: {len(self.batch.cases)}")

    def select_case(self, row):
        if not 0 <= row < len(self.batch.cases):
            return
        self.commit_edits()
        self.undo_stack.clear()
        self.current_case = self.batch.cases[row]
        self.field_model = FieldTableModel(
            self.current_case, self.profile.pin_length, self.profile.identifier_label
        )
        self.field_model.raw_edits = dict(self.case_raw_edits.get(self.current_case.id, {}))
        self.field_model.edited.connect(self.on_edited)
        self.fields.setModel(self.field_model)
        self.image_list.blockSignals(True)
        self.image_list.clear()
        for index, source in enumerate(self.current_case.images):
            self.image_list.addItem(f"Фото {index + 1}", source.id)
        self.image_list.setCurrentIndex(0)
        self.image_list.blockSignals(False)
        self.select_image(0)
        self.refresh_photos()
        self.refresh_form()
        step = (
            4
            if self.current_case.output_path
            else 2
            if self.current_case.data.full_name
            or self.current_case.evidence
            or self.current_case.pending_identity
            else 0
        )
        self.go_step(step)

    def refresh_form(self):
        if not hasattr(self, "field_model"):
            return
        self.field_model.refresh()
        preview_case = self.case_actions.ocr_state.preview(self.current_case)
        display_case = preview_case or self.current_case
        display_model = (
            FieldTableModel(display_case, self.profile.pin_length, self.profile.identifier_label)
            if preview_case
            else self.field_model
        )
        self.identity_status.setVisible(self.current_case.identity_status == "review")
        self.identity_status.setText("Подтвердить принадлежность фотографий")
        self.identity_status.setToolTip(self.current_case.identity_message)
        self.review_form.refresh(
            display_case, display_model, self.profile.required_fields, preview_only=bool(preview_case)
        )
        self.document_type.blockSignals(True)
        self.document_type.setCurrentIndex(
            max(0, self.document_type.findData(display_case.data.document_type))
        )
        self.document_type.blockSignals(False)
        self.upload_document_type.blockSignals(True)
        self.upload_document_type.setCurrentIndex(
            max(0, self.upload_document_type.findData(display_case.data.document_type))
        )
        self.upload_document_type.blockSignals(False)
        reports = self.current_case.mrz_reports.values()
        self.mrz_status.setVisible(
            bool(self.current_case.mrz_reports) or self.current_case.data.document_type == "kg_physical_id"
        )
        self.mrz_status.setText(
            "Есть расхождения между MRZ и распознанными данными"
            if any(r.conflicts for r in reports)
            else "MRZ проверена · сверьте документ"
            if any(r.valid for r in reports)
            else "MRZ не прошла проверку · подробнее"
            if self.current_case.mrz_reports
            else "MRZ ещё не проверялась"
        )
        self.side_review_button.setVisible(bool(self.current_case.side_proposals))
        no_address = (
            self.current_case.data.document_type == "kg_physical_id"
            and self.current_case.registration_status == RegistrationStatus.NOT_PRESENT
            and not self.current_case.data.registration_address.display()
        )
        self.registration_notice.setVisible(no_address)
        self.registration_actions.setVisible(no_address)
        filled = sum(bool(data_value(display_case.data, f)) for f in self.profile.required_fields)
        self.completeness.setText(
            ("Распознано: " if preview_case else "Обязательные поля: ")
            + f"{filled} из {len(self.profile.required_fields)}"
        )
        issues = readiness(self.current_case, self.profile, require_confirmation=False)
        self.review_summary.setText(
            f"Требуют внимания: {len(issues)}. Показать первое поле"
            if issues
            else "Все обязательные поля заполнены. Сверьте их с документом."
        )
        self.review_summary.setProperty("tone", "warning" if issues else "success")
        if preview_case:
            self.review_summary.setText(
                "Данные распознаны и показаны для просмотра. Подтвердите принадлежность фотографий одному человеку, затем проверьте поля."
            )
        if self.current_case.errors:
            self.review_summary.setText(
                self.review_summary.text() + "\n" + " · ".join(self.current_case.errors[:2])
            )
        template = self.current_case.template_path
        self.template_label.setText(template.name if template else "Выберите Excel-шаблон")
        self.profile_label.setText("Профиль: " + self.profile.name)
        self.photo_panel.refresh(display_case, self.profile)
        sheets = self.template_sheets.get(template, [])
        self.template_meta.setText(
            (
                f"Листов: {len(sheets)} · " + ", ".join(sheets)
                if sheets
                else "Листы будут проверены при выборе и перед записью"
            )
            + "\nОбязательные поля: "
            + ", ".join(FIELD_LABELS[f] for f in self.profile.required_fields)
        )
        self.mapping_table.setRowCount(len(self.profile.field_mappings))
        for row, mapping in enumerate(self.profile.field_mappings):
            label = FIELD_LABELS.get(
                mapping.field,
                mapping.field.replace("registration_address.", "Регистрация · ").replace(
                    "living_address.", "Проживание · "
                ),
            )
            special_labels = {
                "tool_full_name": "Инструменты",
                "phones": "Телефоны",
                "purchase_price": "Стоимость покупки",
                "deposit_amount": "Залог",
                "actual_return_date": "Дата возврата",
                "actual_return_time": "Время возврата",
                "registration_address.region": "Регистрация · регион",
                "registration_address.street": "Регистрация · улица и дом",
                "registration_address.full": "Полный адрес регистрации",
                "living_address.region": "Проживание · регион",
                "living_address.street": "Проживание · улица и дом",
                "living_address.full": "Полный адрес проживания",
            }
            label = special_labels.get(mapping.field, label)
            for col, text in enumerate([label, mapping.sheet, mapping.cell]):
                self.mapping_table.setItem(row, col, QTableWidgetItem(text))
        if self.current_step == 0:
            self.primary_button.setEnabled(bool(self.current_case.images) and not self.active_worker)

    def on_edited(self, _field):
        self.current_case.status = "На проверке"
        self.case_raw_edits[self.current_case.id] = dict(self.field_model.raw_edits)
        self.refresh_form()
        self.refresh_cases()

    def refresh_photos(self):
        loaded = bool(self.current_case.images)
        self.dropzone.illustration.setVisible(not loaded)
        self.formats_label.setVisible(not loaded)
        self.upload_summary.setVisible(not loaded)
        self.dropzone.title.setText("Добавьте ещё фотографии" if loaded else "Добавьте фотографии документа")
        self.dropzone.setMinimumHeight(145 if loaded else 220)
        self.dropzone.setMaximumHeight(172 if loaded else 16777215)
        self.case_title.setText(self.current_case.name)
        self.dropzone.layout().setContentsMargins(
            24, 14 if self.current_case.images else 28, 24, 14 if self.current_case.images else 28
        )
        while self.thumbnail_layout.count():
            item = self.thumbnail_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        for index, source in enumerate(self.current_case.images):
            card = PhotoThumbnail(index, source.selected_path, source.selected_rotation, source.role)
            card.selected.connect(self.show_photo)
            card.role_changed.connect(self.set_image_role)
            card.menu_requested.connect(self.photo_menu)
            self.thumbnail_layout.addWidget(card)
        self.thumbnail_layout.addStretch()
        if hasattr(self, "thumbnail_scroll"):
            self.thumbnail_scroll.setVisible(bool(self.current_case.images))

    def show_photo(self, index):
        self.go_step(2)
        self.image_list.setCurrentIndex(index)

    def set_image_role(self, index, role):
        if self.active_worker or not 0 <= index < len(self.current_case.images):
            return
        side = DocumentSide(role)
        if side in {DocumentSide.FRONT, DocumentSide.BACK}:
            for i, image in enumerate(self.current_case.images):
                if i != index and image.role == side:
                    image.role = DocumentSide.ADDITIONAL
        self.current_case.images[index].role = side
        self.current_case.side_proposals.pop(self.current_case.images[index].id, None)
        self.current_case.invalidate()
        self.undo_stack.clear()
        self.refresh_photos()
        self.refresh_form()

    def photo_menu(self, index, position):
        if self.active_worker:
            return
        self.image_list.setCurrentIndex(index)
        menu = QMenu(self)
        menu.addAction("Просмотреть", lambda: self.show_photo(index))
        menu.addAction("Обрезать / проверить фото", lambda: self.case_actions.edit_document_photo(index))
        menu.addAction("Заменить файлом", lambda: self.case_actions.replace_document_photo(index))
        menu.addAction(
            "Переснять эту сторону",
            lambda: self.case_actions.capture_documents(self.current_case.images[index].role),
        )
        menu.addAction("Поменять стороны местами", self.case_actions.swap_document_sides)
        menu.addAction("Повернуть на 90°", self.case_actions.rotate_image)
        for label, offset in [("Раньше", -1), ("Позже", 1)]:
            menu.addAction(label, lambda o=offset: self.reorder_photo(index, o)).setEnabled(
                0 <= index + offset < len(self.current_case.images)
            )
        menu.addSeparator()
        menu.addAction("Перенести / убрать", self.case_actions.manage_image)
        menu.exec(position)

    def reorder_photo(self, index, offset):
        images = self.current_case.images
        images[index], images[index + offset] = images[index + offset], images[index]
        self.current_case.invalidate()
        self.undo_stack.clear()
        self.select_case(self.batch.cases.index(self.current_case))
        self.go_step(0)

    def select_image(self, index):
        self.preview_token += 1
        token = self.preview_token
        if not self.current_case.images:
            self.viewer.clear()
            return
        if not 0 <= index < len(self.current_case.images):
            return
        source = self.current_case.images[index].model_copy(deep=True)
        show_original = self.view_mode.currentIndex() == 1

        def prepare(_cancel, _progress):
            array = load_image(
                source.path if show_original else source.selected_path,
                source.rotation if show_original else source.selected_rotation,
            )
            h, w = array.shape[:2]
            if max(h, w) > 2600:
                array = cv2.resize(array, (round(w * 2600 / max(h, w)), round(h * 2600 / max(h, w))))
            return array, detect_document_region(array)

        def show(result):
            if token != self.preview_token:
                return
            self.viewer.set_image(*result)
            self.photo_info.setText(
                "Показан документ · оригинал доступен в меню просмотра"
                if result[1].detected
                else "Показан оригинал · область можно выделить вручную"
            )
            if self.pending_highlight and self.pending_highlight[0] == source.id:
                self.viewer.highlight(self.pending_highlight[1])

        self.run_worker(prepare, show, busy=False)

    def focus_review_field(self, name):
        row = self.field_model.fields.index(name)
        self.fields.setCurrentIndex(self.field_model.index(row, 1))
        evidence = self.current_case.evidence.get(name, FieldEvidence())
        from document_excel.domain.resolution import displayed_candidate

        candidate = displayed_candidate(evidence, getattr(self.current_case.data, name, None))
        if candidate:
            self.focus_candidate(candidate)
        elif evidence.expected_bbox:
            self.focus_candidate(self.field_hint(name, evidence))

    @staticmethod
    def field_hint(name, evidence):
        from document_excel.domain.models import FieldCandidate

        return FieldCandidate(
            field_name=name,
            raw_value="",
            normalized_value="",
            confidence=0,
            source_image_id=evidence.expected_image_id,
            source_bbox=evidence.expected_bbox,
            image_version=evidence.expected_image_version,
            validation_status=FieldStatus.NOT_RECOGNIZED,
            extraction_method="region-hint",
        )

    def focus_candidate(self, candidate):
        if candidate is None:
            return
        self.pending_highlight = (candidate.source_image_id, candidate.source_bbox)
        for index, source in enumerate(self.current_case.images):
            if source.id == candidate.source_image_id:
                if source.processed_confirmed:
                    self.view_mode.setCurrentIndex(0 if candidate.image_version == "processed" else 1)
                self.image_list.setCurrentIndex(index)
                self.viewer.highlight(candidate.source_bbox)
                break

    def show_first_issue(self):
        for name in self.field_model.fields:
            evidence = self.current_case.evidence.get(name, FieldEvidence())
            if (
                evidence.status in {FieldStatus.INVALID, FieldStatus.CONFLICT}
                or name in self.profile.required_fields
                and not data_value(self.current_case.data, name)
            ):
                self.review_form.reveal(name)
                return

    def show_field_details(self, name):
        self.commit_edits()
        dialog = QDialog(self)
        dialog.setWindowTitle(FIELD_LABELS[name])
        dialog.resize(560, 300)
        layout = QVBoxLayout(dialog)
        evidence = self.current_case.evidence.get(name, FieldEvidence())
        info = QLabel(evidence.message or "Сверьте значение с оригиналом")
        info.setWordWrap(True)
        layout.addWidget(info)
        choices = QComboBox()
        seen = set()
        side_names = {
            "FRONT": "Лицевая сторона",
            "BACK": "Оборотная сторона",
            "UNKNOWN": "Сторона не определена",
            "ADDITIONAL": "Дополнительное фото",
        }
        for candidate in evidence.candidates:
            key = (candidate.normalized_value, candidate.source_image_id, candidate.source_method)
            if key in seen:
                continue
            seen.add(key)
            choices.addItem(
                candidate.raw_value
                + " · "
                + ("MRZ" if candidate.source_method == "MRZ" else side_names[candidate.source_side.value]),
                candidate,
            )
        if not evidence.candidates and evidence.expected_bbox:
            choices.addItem("Область поля для повторного распознавания", self.field_hint(name, evidence))
        choices.currentIndexChanged.connect(lambda: self.focus_candidate(choices.currentData()))
        layout.addWidget(choices)
        self.focus_candidate(choices.currentData())
        repeat = self.make_button("Повторно распознать область поля", lambda: dialog.done(2))
        repeat.setObjectName("repeatFieldOcr")
        repeat.setEnabled(any(c.source_bbox is not None for c in evidence.candidates))

        def update_repeat():
            selected = choices.currentData()
            repeat.setText(
                "Открыть проверку MRZ"
                if selected and selected.source_method == "MRZ"
                else "Повторно распознать область поля"
            )
            repeat.setEnabled(selected is not None and selected.source_bbox is not None)

        choices.currentIndexChanged.connect(update_repeat)
        update_repeat()
        layout.addWidget(repeat)
        accept = self.make_button(
            "Использовать выбранное значение",
            lambda: (self.apply_field_edit(name, choices.currentData().normalized_value), dialog.accept()),
        )
        accept.setEnabled(
            bool(evidence.candidates)
            and any(
                c.normalized_value and c.validation_status != FieldStatus.INVALID for c in evidence.candidates
            )
            and name not in {"registration_address", "living_address"}
        )
        layout.addWidget(accept)
        layout.addWidget(
            self.make_button(
                "Копировать значение",
                lambda: QApplication.clipboard().setText(data_value(self.current_case.data, name)),
            )
        )
        close = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        close.rejected.connect(dialog.reject)
        layout.addWidget(close)
        if dialog.exec() == 2:
            if choices.currentData().source_method == "MRZ":
                self.case_actions.show_mrz()
            else:
                self.case_actions.repeat_field(name, choices.currentData())

    def update_zoom(self, value):
        self.zoom_label.setText(f"{value}%")

    def run_worker(self, function, result, busy=True, cancellable=True):
        if busy and self.active_worker:
            self.notice.setText("Сначала завершите текущую операцию.")
            return
        worker = Worker(function, cancellable)
        self.background_workers.append(worker)
        if busy:
            self.active_worker = worker
            self.set_busy(True)
            self.cancel_button.setEnabled(cancellable)
        receiver = WorkerCallbacks(self, worker, result)
        worker.callbacks = receiver
        worker.signals.result.connect(receiver.receive, Qt.ConnectionType.QueuedConnection)
        worker.signals.error.connect(self.error, Qt.ConnectionType.QueuedConnection)
        worker.signals.progress.connect(self.on_progress, Qt.ConnectionType.QueuedConnection)
        worker.signals.finished.connect(receiver.finish, Qt.ConnectionType.QueuedConnection)
        self.pool.start(worker)

    def on_progress(self, value, text):
        if self.current_step == 1:
            self.progress.setRange(0, 0)
            self.substage_label.setText(text)
        else:
            self.progress.setRange(0, 100)
            self.progress.setValue(value)
            self.notice.setText(text)

    def worker_finished(self, worker):
        if worker in self.background_workers:
            self.background_workers.remove(worker)
        if worker is self.active_worker:
            self.active_worker = None
            self.set_busy(False)
            if self.current_step == 1:
                self.go_step(2)
            else:
                self.go_step(self.current_step)
        if self.closing:
            QTimer.singleShot(100, self, self.close)

    def set_busy(self, busy):
        for widget in [
            self.add_button,
            self.folder_button,
            self.ocr_button,
            self.batch_ocr_button,
            self.export_button,
            self.batch_export_button,
            self.fields,
            self.cases,
            self.image_list,
            self.confirm_button,
            self.document_type,
            self.template_button,
            self.review_form,
            self.queue_table,
            self.dropzone,
            self.thumbnail_scroll,
            self.new_case_button,
            self.settings_button,
        ]:
            widget.setEnabled(not busy)
        for button in self.step_indicator.buttons:
            button.setEnabled(not busy)
        self.cancel_button.setEnabled(busy)
        self.primary_button.setEnabled(
            busy
            if self.current_step == 1
            else not busy and (self.current_step != 0 or bool(self.current_case.images))
        )

    def cancel_work(self):
        if self.active_worker:
            if hasattr(self, "batch_export_queue"):
                self.batch_export_queue.clear()
            self.active_worker.cancel()
            self.substage_label.setText("Завершаем текущий проход…")
            self.notice.setText("Отмена запрошена. Данные клиента будут сохранены в памяти.")
        elif self.viewer.selecting:
            self.viewer.cancel_selection()

    def message(self, text):
        QMessageBox.information(self, "Документ → Excel", str(text))

    def error(self, text):
        self.notice.setText(str(text))
        if not self.closing:
            QMessageBox.warning(self, "Нужно внимание", str(text))

    def show_settings(self):
        self.commit_edits()
        if self.active_worker:
            return
        dialog = SettingsDialog(self.settings, self)
        if dialog.exec():
            value = dialog.value()
            if value.temp_directory:
                try:
                    with tempfile.TemporaryFile(dir=value.temp_directory):
                        pass
                except OSError:
                    self.error("Временная папка недоступна. Выберите существующую папку с правом записи.")
                    return
            self.settings = value
            tempfile.tempdir = value.temp_directory or None
            self.pool.setMaxThreadCount(value.max_parallel)
            apply_theme(QApplication.instance(), value)
            self.save_settings()
            self.diagnose()

    def save_settings(self):
        try:
            self.repository.save(self.settings)
        except AppError as exc:
            self.error(exc)

    def diagnose(self):
        settings = self.settings.model_copy(deep=True)

        def done(diagnostic):
            self.ocr_diagnostic = diagnostic
            if hasattr(self, "ocr_status"):
                self.ocr_status.setText("Настройте OCR" if diagnostic.error else "OCR готов")

        self.run_worker(lambda _c, _p: create_engine(settings).diagnose(), done, busy=False)

    def show_diagnostics(self):
        value = self.ocr_diagnostic
        if value and not value.error:
            neural = self.settings.ocr_engine in {"paddle", "rapidocr"}
            details = value.version + "\nЯзыки: " + ", ".join(value.languages)
            details += "\nЛокальные модели / программа: " + value.executable
            if neural:
                details += f"\nУстройство: CPU · потоки: {self.settings.paddle_cpu_threads}"
                details += "\nДетектор: PP-OCRv5 mobile · чтение: Cyrillic / English PP-OCRv5"
            self.message(details)
            return
        self.message(
            (value.error or value.version + "\nМодели: " + ", ".join(value.languages))
            if value
            else "Диагностика ещё выполняется"
        )

    def show_raw_ocr(self):
        import json

        dialog = QDialog(self)
        dialog.setWindowTitle("Диагностика OCR · значения скрыты")
        dialog.resize(700, 560)
        layout = QVBoxLayout(dialog)
        text = QTextEdit()
        text.setReadOnly(True)
        case = self.current_case
        diagnostic_case = self.case_actions.ocr_state.preview(case) or case
        report = {
            "identity_status": case.identity_status,
            "pending_identity": bool(case.pending_identity),
            "runs": [
                {
                    "run_id": r.id,
                    "status": r.status,
                    "started": r.started_at.isoformat(),
                    "completed": r.completed_at.isoformat() if r.completed_at else None,
                    "images": len(r.image_hashes),
                }
                for r in case.ocr_runs.values()
            ],
            "events": case.diagnostic_events,
            "blocks": case.diagnostic_blocks,
            "processing": diagnostic_case.processing_metrics,
            "fields": {
                name: {
                    "status": e.status.value,
                    "reason": e.resolution_reason,
                    "manual_conflict": e.manual_conflict,
                    "candidates": [
                        {
                            "method": c.extraction_method,
                            "image_id": c.source_image_id,
                            "side": c.source_side.value,
                            "run_id": c.run_id,
                            "bbox": c.source_bbox,
                            "status": c.validation_status.value,
                            "rejections": len(c.validation_errors),
                        }
                        for c in e.candidates
                    ],
                }
                for name, e in diagnostic_case.evidence.items()
            },
        }
        text.setPlainText(json.dumps(report, ensure_ascii=False, indent=2))
        layout.addWidget(text)

        def reveal():
            if (
                QMessageBox.question(
                    dialog,
                    "Временный просмотр реквизитов",
                    "Показать полный OCR-текст и кандидатов текущего клиента в этом окне? Данные останутся в памяти и не попадут в обычный журнал.",
                )
                == QMessageBox.StandardButton.Yes
            ):
                text.setPlainText(
                    json.dumps(
                        {
                            "ocr": self.case_ocr_texts.get(case.id, {}),
                            "candidates": {
                                name: [c.model_dump(mode="json") for c in e.candidates]
                                for name, e in (
                                    self.case_actions.ocr_state.preview(case) or case
                                ).evidence.items()
                            },
                        },
                        ensure_ascii=False,
                        indent=2,
                    )
                )
                dialog.setWindowTitle("Чувствительные данные · временный просмотр")

        layout.addWidget(self.make_button("Временно показать чувствительные значения", reveal))
        dialog.exec()
        text.clear()
        dialog.deleteLater()

    def record_history(self, operation, result):
        if self.settings.history_enabled:
            try:
                self.history.add(operation, result)
            except OSError:
                logging.getLogger("document_excel").warning("history_write_failed")

    def show_history(self):
        self.commit_edits()
        self.history_text.setPlainText(
            "\n".join(
                f"{r.get('time', '')}   {r.get('operation', '')}   {r.get('result', '')}"
                for r in reversed(self.history.read())
            )
            or "История пока пуста."
        )
        self.pages.setCurrentIndex(2)

    def clear_history(self):
        try:
            self.history.clear()
            self.show_history()
        except OSError:
            self.error("Не удалось удалить историю.")

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "review_splitter"):
            from document_excel.presentation.workflow_view import update_responsive

            update_responsive(self, self.width())

    def closeEvent(self, event):
        cameras = [
            dialog
            for dialog in self.findChildren(QDialog)
            if dialog.property("documentCameraDialog") and dialog.service.busy
        ]
        if cameras:
            self.closing = True
            for dialog in cameras:
                dialog.reject()
            QTimer.singleShot(100, self, self.close)
            event.ignore()
            return
        if self.background_workers or self.pool.activeThreadCount():
            self.closing = True
            for worker in self.background_workers:
                worker.cancel()
            QTimer.singleShot(200, self, self.close)
            event.ignore()
            return
        self.settings.window_width, self.settings.window_height = self.width(), self.height()
        self.capture_controller.cleanup()
        try:
            self.repository.save(self.settings)
        except AppError:
            logging.getLogger("document_excel").warning("settings_write_failed_on_close")
        event.accept()
