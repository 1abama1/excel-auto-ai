from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QObject, Qt, QTimer, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QFileDialog,
    QInputDialog,
)

from document_excel.domain.errors import AppError, OutputExistsError
from document_excel.domain.models import WorkflowMode
from document_excel.infrastructure.contract_filename import ContractFilenameGenerator
from document_excel.infrastructure.excel import (
    create_demo_template,
    readiness,
)
from document_excel.infrastructure.storage import copy_originals
from document_excel.presentation.dialogs import (
    ExportDialog,
    MappingDialog,
)


class ExcelActions(QObject):
    def __init__(self, window) -> None:
        super().__init__(window)
        self.window = window
        self.filenames = ContractFilenameGenerator()

    def photo_exception(self, allowed):
        if self.window.active_worker or not self.window.profile.allow_photo_exception:
            return
        case = self.window.current_case
        confirmed = case.confirmed
        case.photo_exception = allowed
        case.invalidate()
        if confirmed:
            # Only the explicitly acknowledged photo exception changed; the
            # previously reviewed client data and photos remain unchanged.
            case.confirmed_digest = case.digest()
        self.window.refresh_form()

    def choose_template(self):
        if self.window.active_worker:
            return
        self.window.commit_edits()
        name, _ = QFileDialog.getOpenFileName(
            self.window,
            "Выберите готовый договор Excel",
            self.window.settings.template_directory,
            "Excel (*.xlsx)",
        )
        if name:
            path = Path(name)
            case = self.window.current_case

            def loaded(description):
                sheets, profile = description
                case.template_path = path
                case.workbook_profile = profile.model_dump(mode="json")
                case.output_path = None
                case.invalidate()
                self.window.template_sheets[path] = sheets
                if profile != self.window.profile:
                    self.use_profile(profile)
                self.window.settings.default_template = str(path)
                self.window.save_settings()
                self.window.refresh_form()
                self.window.notice.setText(
                    "Договор распознан. Можно продолжить"
                    if profile.field_mappings
                    else "Структура договора неизвестна. В разделе «Дополнительно» настройте расположение полей."
                )

            self.window.run_worker(
                lambda _c, _p: self.window.workbooks.describe(
                    path,
                    self.window.profile,
                    strict=case.workflow_mode == WorkflowMode.EXISTING_CONTRACT,
                    registry=self.window.registry,
                ),
                loaded,
            )

    def edit_mapping(self):
        self.edit_profile(self.window.profile)

    def edit_profile(self, candidate):
        if self.window.active_worker:
            return
        path = self.window.current_case.template_path
        if not path:
            self.window.error("Сначала выберите XLSX, чтобы увидеть реальные листы.")
            return

        def loaded(sheets):
            dialog = MappingDialog(candidate, sheets, self.window)
            if dialog.exec():
                try:
                    self.use_profile(self.window.workbooks.bind_profile(path, dialog.result_profile))
                except AppError as exc:
                    self.window.error(exc)

        self.window.run_worker(lambda _c, _p: self.window.workbooks.sheets(path), loaded)

    def new_mapping(self):
        from document_excel.domain.models import new_id
        from document_excel.infrastructure.excel import TemplateProfile

        self.edit_profile(TemplateProfile(id=new_id(), tool_sheet=None, field_mappings=[]))

    def use_profile(self, profile):
        self.window.commit_edits()
        try:
            saved = self.window.registry.save(profile)
            self.window.profile = profile
            self.window.undo_stack.clear()
            self.window.settings.default_profile = str(saved)
            case = self.window.current_case
            case.workbook_profile = profile.model_dump(mode="json")
            case.output_path = None
            case.invalidate()
            self.window.field_model.pin_length = profile.pin_length
            self.window.field_model.identifier_label = profile.identifier_label
            self.window.profile_label.setText("Текущий профиль: " + profile.name)
            self.window.save_settings()
            self.window.refresh_form()
        except AppError as exc:
            self.window.error(exc)

    def import_profile(self):
        if self.window.active_worker:
            return
        name, _ = QFileDialog.getOpenFileName(self.window, "Импорт mapping", "", "JSON (*.json)")
        if name:
            try:
                self.edit_profile(self.window.registry.load(Path(name)))
            except AppError as exc:
                self.window.error(exc)

    def export_profile(self):
        name, _ = QFileDialog.getSaveFileName(self.window, "Экспорт mapping", "mapping.json", "JSON (*.json)")
        if name:
            from document_excel.infrastructure.storage import atomic_json

            try:
                atomic_json(Path(name), self.window.profile.model_dump())
            except OSError:
                self.window.error("Не удалось сохранить профиль.")

    def demo_template(self):
        if self.window.active_worker:
            return
        self.window.commit_edits()
        name, _ = QFileDialog.getSaveFileName(
            self.window, "Только демонстрационный шаблон", "DEMO-not-a-contract.xlsx", "Excel (*.xlsx)"
        )
        if name:
            case = self.window.current_case

            def done(sheets):
                case.template_path = Path(name)
                case.output_path = None
                case.invalidate()
                _, self.window.profile = self.window.workbooks.describe(
                    Path(name), self.window.profile, strict=True
                )
                case.workbook_profile = self.window.profile.model_dump(mode="json")
                self.window.template_sheets[Path(name)] = sheets
                self.window.refresh_form()
                self.window.notice.setText(
                    "Выбран демонстрационный XLSX. Совместимость с вашим договором ещё не проверена."
                )

            def create(_cancel, _progress):
                create_demo_template(Path(name))
                return self.window.workbooks.sheets(Path(name))

            self.window.run_worker(create, done)

    def export_current(self):
        if self.window.active_worker:
            return
        self.window.commit_edits()
        case = self.window.current_case
        if not case.template_path:
            self.window.error("Выберите Excel-шаблон.")
            return
        snapshot = case.model_copy(deep=True)
        profile = self.window.profile.model_copy(deep=True)
        self.window.run_worker(
            lambda _c, _p: self.window.workbooks.plan(snapshot.template_path, snapshot, profile),
            lambda plan: self.preview_and_save(plan, case, profile),
        )

    def preview_and_save(self, plan, case, profile):
        dialog = ExportDialog(plan, self.window)
        if not dialog.exec():
            return
        try:
            filename = self.filenames.generate(
                case.formatted_name, preserve=True, pattern=self.window.settings.filename_pattern
            )
        except AppError as exc:
            self.window.error(exc)
            return
        default = Path(self.window.settings.output_directory or str(plan.source.parent)) / filename
        numbered = False
        while True:
            name, _ = QFileDialog.getSaveFileName(
                self.window,
                "Создать договор · новая копия",
                str(default),
                "Excel (*.xlsx)",
                "",
                QFileDialog.Option.DontConfirmOverwrite,
            )
            if not name:
                return
            target = Path(name)
            if target.suffix.lower() != ".xlsx":
                target = target.with_suffix(".xlsx")
            automatic = target.name == filename
            if not target.exists():
                break
            choice, ok = QInputDialog.getItem(
                self.window,
                "Файл уже существует",
                "Исходный файл сохранится. Выберите действие:",
                ["Сохранить с дополнительным номером", "Изменить имя", "Выбрать другую папку", "Отменить"],
                0,
                False,
            )
            if not ok or choice == "Отменить":
                return
            if choice == "Сохранить с дополнительным номером":
                numbered = True
                break
            if choice == "Выбрать другую папку":
                folder = QFileDialog.getExistingDirectory(self.window, "Папка договора", str(target.parent))
                if not folder:
                    return
                default = Path(folder) / filename
            else:
                default = target
        if name:
            # The plan worker's finished signal is queued after result; defer the next job.
            QTimer.singleShot(
                0,
                self.window,
                lambda: self.save_plan(plan, target, case, profile, automatic=automatic, numbered=numbered),
            )

    def save_plan(self, plan, path, case, profile, *, automatic=False, numbered=False):
        def done(result):
            case.output_path = result
            case.status = "Excel сохранён"
            self.window.last_result = result
            self.window.settings.output_directory = str(result.parent)
            self.window.save_settings()
            self.window.refresh_cases()
            self.window.notice.setText("Excel проверен и сохранён: " + result.name)
            self.window.record_history("Экспорт Excel", "успешно")
            if case.id == self.window.current_case.id:
                self.window.refresh_form()
                QTimer.singleShot(
                    0,
                    self.window,
                    lambda: self.window.go_step(4) if case.id == self.window.current_case.id else None,
                )

        def write(_c, _p):
            target = path
            if automatic:
                target = path.parent / self.filenames.generate(
                    case.formatted_name, preserve=True, pattern=self.window.settings.filename_pattern
                )
            return self.write_target(plan, target, case, profile, numbered)

        self.window.run_worker(write, done, cancellable=False)

    def write_target(self, plan, target, case, profile, numbered):
        base = target
        for _ in range(32):
            target = self.filenames.available(base) if numbered else base
            try:
                return self.window.workbooks.write(plan, target, case, profile)
            except OutputExistsError:
                if not numbered:
                    raise
        raise AppError("Имя занято другим сохранением. Повторите сохранение в другую папку.")

    def template_generation(self):
        from document_excel.presentation.components import ConfirmDialog

        if not self.window.active_worker and ConfirmDialog.ask(
            self.window,
            "Дополнительный режим",
            "Включить заполнение шаблона? В этом режиме введённые инструменты, цены, залог и возврат могут заменять данные договора.",
        ):
            self.window.current_case.workflow_mode = WorkflowMode.TEMPLATE_GENERATION
            self.window.current_case.invalidate()
            self.window.refresh_form()

    def existing_mode(self):
        if not self.window.active_worker:
            self.window.current_case.workflow_mode = WorkflowMode.EXISTING_CONTRACT
            self.window.current_case.invalidate()
            self.window.refresh_form()

    def export_batch(self):
        if self.window.active_worker:
            return
        self.window.commit_edits()
        pending = [
            case
            for case in self.window.batch.cases
            if case.output_path is None or case.status != "Excel сохранён"
        ]
        if not pending:
            self.window.message("Все записи очереди уже сохранены.")
            return
        folder = QFileDialog.getExistingDirectory(
            self.window, "Папка отдельных договоров", self.window.settings.output_directory
        )
        if not folder:
            return
        self.window.batch_export_queue = list(pending)
        self.window.batch_export_total = len(pending)
        self.window.batch_export_folder = Path(folder)
        self.next_batch_export()

    def next_batch_export(self):
        if not self.window.batch_export_queue:
            self.window.progress.setValue(100)
            self.window.notice.setText("Пакетный экспорт завершён. Проверьте статусы записей в очереди.")
            return
        case = self.window.batch_export_queue.pop(0)
        self.batch_case = case
        position = self.window.batch_export_total - len(self.window.batch_export_queue)
        self.window.progress.setValue(round((position - 1) * 100 / self.window.batch_export_total))
        self.window.notice.setText(f"Экспорт договора {position} из {self.window.batch_export_total}")
        from document_excel.infrastructure.excel import TemplateProfile

        profile = (
            TemplateProfile.model_validate(case.workbook_profile)
            if case.workbook_profile
            else self.window.profile.model_copy(deep=True)
        )
        if not case.template_path or readiness(case, profile):
            case.status = "Экспорт: нужны данные"
            self.window.refresh_cases()
            QTimer.singleShot(0, self.window, self.next_batch_export)
            return

        def planned(plan):
            if ExportDialog(plan, self.window).exec():
                QTimer.singleShot(0, self.window, lambda: self.write_batch_plan(plan, case, profile))
            else:
                QTimer.singleShot(0, self.window, self.next_batch_export)

        self.window.run_worker(
            lambda _c, _p: self.window.workbooks.plan(
                case.template_path, case.model_copy(deep=True), profile
            ),
            planned,
        )
        if self.window.active_worker:
            self.window.active_worker.signals.error.connect(
                self.batch_case_error, Qt.ConnectionType.QueuedConnection
            )

    def batch_case_error(self, _text):
        self.batch_case.status = "Экспорт: ошибка"
        self.window.refresh_cases()
        QTimer.singleShot(0, self.window, self.next_batch_export)

    def write_batch_plan(self, plan, case, profile):
        self.batch_case = case

        def done(path):
            case.output_path = path
            case.status = "Excel сохранён"
            self.window.last_result = path
            self.window.refresh_cases()
            self.window.record_history("Пакетный экспорт", "сохранён договор")
            if case.id == self.window.current_case.id:
                self.window.refresh_form()

        def worker_function(_c, _p):
            target = self.window.batch_export_folder / self.filenames.generate(
                case.formatted_name, preserve=True, pattern=self.window.settings.filename_pattern
            )
            return self.write_target(plan, target, case, profile, True)

        self.window.run_worker(worker_function, done, cancellable=False)
        if self.window.active_worker:
            self.window.active_worker.signals.error.connect(
                self.mark_batch_write_error, Qt.ConnectionType.QueuedConnection
            )
            self.window.active_worker.signals.finished.connect(
                self.batch_write_finished, Qt.ConnectionType.QueuedConnection
            )

    def batch_write_finished(self):
        QTimer.singleShot(0, self.window, self.next_batch_export)

    def mark_batch_write_error(self, _text):
        self.batch_case.status = "Экспорт: ошибка записи"
        self.window.refresh_cases()

    def save_photos(self):
        if self.window.active_worker:
            return
        if not self.window.settings.copy_images:
            self.window.message(
                "Включите копирование фото в настройках. По умолчанию оригиналы не копируются."
            )
            return
        folder = QFileDialog.getExistingDirectory(
            self.window, "Папка копий оригинальных фото", self.window.settings.image_directory
        )
        if folder:
            paths = [
                path
                for source in self.window.current_case.images
                for path in [source.path, source.processed_path]
                if path is not None
            ]
            case_id = self.window.current_case.id
            self.window.run_worker(
                lambda _c, _p: copy_originals(paths, Path(folder), case_id),
                lambda copied: self.window.notice.setText(
                    f"Сохранено копий оригинальных фото: {len(copied)}."
                ),
            )
            self.window.settings.image_directory = folder
            self.window.save_settings()

    def open_result(self):
        path = self.window.current_case.output_path
        if path and path.exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
        else:
            self.window.message("Сначала сохраните Excel.")

    def open_result_folder(self):
        path = self.window.current_case.output_path
        if path and path.parent.exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(path.parent)))
        else:
            self.window.message("Сначала сохраните Excel.")
