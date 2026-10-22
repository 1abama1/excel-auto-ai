from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QObject, Qt, QTimer
from PySide6.QtGui import QPixmap, QTransform
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
)

from document_excel.application.ocr_state import OcrState
from document_excel.application.workflows import RecognitionService
from document_excel.domain.errors import AppError
from document_excel.domain.models import (
    FIELD_LABELS,
    ClientCase,
    DocumentBatch,
    DocumentSide,
    FieldCandidate,
    FieldEvidence,
    FieldStatus,
    RegistrationStatus,
    WorkflowMode,
    data_value,
)
from document_excel.infrastructure.excel import (
    readiness,
)
from document_excel.infrastructure.extraction import apply_candidates
from document_excel.presentation.components import ConfirmDialog
from document_excel.presentation.dialogs import (
    AddressDialog,
    ContractDialog,
)


class CaseActions(QObject):
    def __init__(self, window) -> None:
        super().__init__(window)
        self.window = window
        self.ocr_state = OcrState()

    def new_case(self):
        if self.window.active_worker:
            return
        self.window.commit_edits()
        case = ClientCase(
            name=f"Клиент {len(self.window.batch.cases) + 1}",
            workflow_mode=WorkflowMode.BEST_EFFORT_UPDATE,
        )
        self.window.batch.cases.append(case)
        self.window.refresh_cases()
        self.window.select_case(len(self.window.batch.cases) - 1)
        self.window.pages.setCurrentIndex(0)
        self.window.go_step(0)

    def rename_case(self, _item=None):
        name, ok = QInputDialog.getText(
            self.window, "Название записи · только в памяти", "Название", text=self.window.current_case.name
        )
        if ok and name.strip():
            self.window.current_case.name = name.strip()[:100]
            self.window.refresh_cases()

    def manage_case(self):
        if self.window.active_worker:
            return
        self.window.commit_edits()
        choices = [
            f"Объединить с: {i + 1}. {c.name}"
            for i, c in enumerate(self.window.batch.cases)
            if c.id != self.window.current_case.id
        ]
        choices.append("Удалить запись из памяти")
        choice, ok = QInputDialog.getItem(
            self.window, "Группа клиента", "Выберите действие", choices, 0, False
        )
        if not ok:
            return
        self.window.undo_stack.clear()
        source = self.window.current_case
        selected_case = source
        if choice != choices[-1]:
            target_index = int(choice.split(": ")[1].split(".")[0]) - 1
            target = self.window.batch.cases[target_index]
            self.window.case_service.merge(target, source)
            selected_case = target
        elif (
            QMessageBox.question(
                self.window,
                "Удалить запись",
                "Удалить данные клиента из памяти? Исходные фото и Excel останутся на диске.",
            )
            != QMessageBox.StandardButton.Yes
        ):
            return
        self.window.batch.cases = [case for case in self.window.batch.cases if case.id != source.id]
        self.window.case_ocr_texts.pop(source.id, None)
        if not self.window.batch.cases:
            self.window.batch.cases.append(
                ClientCase(name="Клиент 1", workflow_mode=WorkflowMode.BEST_EFFORT_UPDATE)
            )
        if selected_case is source:
            selected_case = self.window.batch.cases[0]
        self.window.refresh_cases()
        self.window.select_case(self.window.batch.cases.index(selected_case))
        self.window.case_raw_edits.pop(source.id, None)
        self.window.case_steps.pop(source.id, None)

    def choose_images(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self.window,
            "Добавить фотографии",
            "",
            "Изображения (*.jpg *.jpeg *.png *.webp *.bmp *.tif *.tiff)",
        )
        if paths:
            self.add_images([Path(path) for path in paths])

    def choose_folder(self):
        path = QFileDialog.getExistingDirectory(self.window, "Фото одного клиента")
        if path:
            self.add_images([Path(path)])

    def add_images(self, paths):
        if self.window.active_worker:
            self.window.notice.setText("Дождитесь текущей операции или нажмите «Отмена».")
            return
        self.window.commit_edits()
        errors = self.window.case_service.add_images(self.window.current_case, paths)
        self.window.undo_stack.clear()
        self.window.current_case.status = "Фото загружены"
        self.window.refresh_cases()
        self.window.select_case(self.window.batch.cases.index(self.window.current_case))
        self.window.go_step(0)
        self.window.notice.setText(
            " ".join(errors) if errors else "Фотографии добавлены. Нажмите «Распознать документы»."
        )

    def choose_side_image(self, side):
        if self.window.active_worker:
            return
        name, _ = QFileDialog.getOpenFileName(
            self.window,
            "Фотография паспорта",
            "",
            "Изображения (*.jpg *.jpeg *.png *.webp *.bmp *.tif *.tiff)",
        )
        if not name:
            return
        case = self.window.current_case
        previous = [image for image in case.images if image.role == side]
        if previous and not ConfirmDialog.ask(
            self.window, "Заменить фотографию", "Заменить выбранную сторону паспорта новым снимком?"
        ):
            return
        old_ids = {image.id for image in case.images}
        self.add_images([Path(name)])
        added = [image for image in case.images if image.id not in old_ids]
        if added:
            for image in previous:
                self.window.case_service.remove_evidence(case, image.id)
                case.images.remove(image)
            added[0].role = side
            case.invalidate()
            self.window.select_case(self.window.batch.cases.index(case))
            self.window.go_step(0)

    def rotate_image(self):
        index = self.window.image_list.currentIndex()
        if self.window.active_worker or index < 0:
            return
        self.window.commit_edits()
        self.window.undo_stack.clear()
        source = self.window.current_case.images[index]
        if source.processed_confirmed and source.processed_path:
            import numpy as np

            from document_excel.infrastructure.document_images import selected_image

            try:
                source = self.window.capture_controller.photo(
                    source.role,
                    np.rot90(selected_image(source), -1).copy(),
                    source=source,
                    retain=self.capture_directory(),
                    case_id=self.window.current_case.id,
                )
                self.window.current_case.images[index] = source
            except (AppError, OSError):
                self.window.error("Не удалось повернуть обработанную фотографию. Проверьте папку хранения.")
                return
        else:
            source.rotation = (source.rotation + 90) % 360
        self.window.case_ocr_texts.get(self.window.current_case.id, {}).pop(source.id, None)
        self.window.case_service.remove_evidence(self.window.current_case, source.id)
        self.window.select_image(index)
        self.window.refresh_photos()
        self.window.refresh_form()

    def capture_documents(self, side=None):
        if self.window.active_worker:
            return
        from document_excel.presentation.camera_dialog import CameraDialog

        self.window.commit_edits()
        case = self.window.current_case
        target = side if side in {DocumentSide.FRONT, DocumentSide.BACK} else None
        dialog = CameraDialog(self.window, side=target, physical=case.data.document_type == "kg_physical_id")
        accepted = []

        def commit(role, original, processed):
            retain = self.capture_directory()
            source = self.window.capture_controller.photo(
                role, processed, original=original, retain=retain, case_id=case.id
            )
            self.window.capture_controller.assign(case, source)
            accepted.append(source.id)
            self.window.undo_stack.clear()
            if self.window.current_case is case:
                self.window.refresh_photos()
                self.window.refresh_form()
            self.window.refresh_cases()

        dialog.commit_photo = commit
        dialog.exec()
        dialog.commit_photo = None
        dialog.deleteLater()
        if accepted and self.window.current_case is case and not self.window.closing:
            self.window.select_case(self.window.batch.cases.index(case))
            self.window.go_step(0)

    def capture_directory(self):
        settings = self.window.settings
        if settings.copy_images:
            if not settings.image_directory:
                folder = QFileDialog.getExistingDirectory(
                    self.window, "Папка оригиналов и обработанных фотографий"
                )
                if not folder:
                    raise AppError(
                        "Снимок ещё не принят. Выберите папку хранения или отключите постоянное хранение в настройках."
                    )
                settings.image_directory = folder
                self.window.save_settings()
            return Path(settings.image_directory)
        return None

    def edit_document_photo(self, index, replacement=None):
        if self.window.active_worker or not 0 <= index < len(self.window.current_case.images):
            return
        from document_excel.domain.models import SourceImage
        from document_excel.infrastructure.document_images import selected_image
        from document_excel.infrastructure.preprocessing import load_image
        from document_excel.presentation.camera_dialog import PhotoReviewDialog

        case = self.window.current_case
        source = case.images[index]
        side = source.role
        if side not in {DocumentSide.FRONT, DocumentSide.BACK}:
            value, ok = QInputDialog.getItem(
                self.window,
                "Сторона документа",
                "Куда относится фотография?",
                ["Лицевая сторона", "Оборотная сторона"],
                0,
                False,
            )
            if not ok:
                return
            side = DocumentSide.FRONT if value == "Лицевая сторона" else DocumentSide.BACK
        original_source = source
        if replacement is not None:
            source = SourceImage(path=replacement, role=side)
        dialog = None
        try:
            dialog = PhotoReviewDialog(load_image(source.path, source.rotation), self.window)
            if source.processed_confirmed:
                dialog.editor.current = selected_image(source)
                dialog.editor.refresh()
            if dialog.exec():
                photo = self.window.capture_controller.photo(
                    side,
                    dialog.editor.current,
                    source=source,
                    retain=self.capture_directory(),
                    case_id=case.id,
                )
                if original_source.id != source.id:
                    self.window.case_service.remove_evidence(case, original_source.id)
                    case.images = [s for s in case.images if s.id != original_source.id]
                self.window.capture_controller.assign(case, photo)
                self.window.undo_stack.clear()
                self.window.select_case(self.window.batch.cases.index(case))
                self.window.go_step(0)
                self.window.notice.setText(
                    "Фотография подтверждена. Повторите распознавание и сверку данных."
                )
        except (AppError, OSError) as exc:
            self.window.error(
                exc
                if isinstance(exc, AppError)
                else "Не удалось сохранить фотографию. Проверьте папку и свободное место."
            )
        finally:
            if dialog is not None:
                dialog.deleteLater()

    def replace_document_photo(self, index):
        name, _ = QFileDialog.getOpenFileName(
            self.window,
            "Заменить фотографию стороны",
            "",
            "Изображения (*.jpg *.jpeg *.png *.webp *.bmp *.tif *.tiff)",
        )
        if name:
            self.edit_document_photo(index, Path(name))

    def review_document_side(self, role):
        source = next(
            (
                i
                for i, image in enumerate(self.window.current_case.images)
                if image.role == DocumentSide(role)
            ),
            None,
        )
        if source is not None:
            self.edit_document_photo(source)

    def swap_document_sides(self):
        if not self.window.active_worker:
            self.window.capture_controller.swap(self.window.current_case)
            self.window.undo_stack.clear()
            self.window.refresh_photos()
            self.window.refresh_form()

    def manage_image(self):
        index = self.window.image_list.currentIndex()
        if self.window.active_worker or index < 0:
            return
        self.window.commit_edits()
        targets = [case for case in self.window.batch.cases if case.id != self.window.current_case.id]
        labels = [f"{i + 1}. {case.name}" for i, case in enumerate(targets)] + ["Убрать из клиента"]
        choice, ok = QInputDialog.getItem(
            self.window, "Фотография", "Перенести к другому клиенту или убрать ссылку", labels, 0, False
        )
        if not ok:
            return
        self.window.undo_stack.clear()
        source = self.window.current_case.images[index]
        self.window.case_ocr_texts.get(self.window.current_case.id, {}).pop(source.id, None)
        if choice == labels[-1]:
            self.window.current_case.images.pop(index)
            self.window.case_service.remove_evidence(self.window.current_case, source.id)
        else:
            self.window.case_service.move_image(
                self.window.current_case, targets[int(choice.split(".")[0]) - 1], source
            )
        self.window.refresh_cases()
        self.window.select_case(self.window.batch.cases.index(self.window.current_case))

    def start_ocr(self, _checked=False, all_cases=False, roi=None):
        if self.window.active_worker:
            return
        self.window.commit_edits()
        cases = self.window.batch.cases if all_cases else [self.window.current_case]
        snapshots = [
            (case.id, [image.model_copy(deep=True) for image in case.images]) for case in cases if case.images
        ]
        if not snapshots:
            self.window.error("Добавьте фотографии к клиенту.")
            return
        settings = self.window.settings.model_copy(deep=True)
        forced_types = {
            case.id: case.data.document_type
            for case in cases
            if case.document_type_manual
            and case.data.document_type in {"kg_physical_id", "kg_digital_id", "manual"}
        }
        if roi:
            selected_index = self.window.image_list.currentIndex()
            if not 0 <= selected_index < len(self.window.current_case.images):
                return
            selected = self.window.current_case.images[selected_index]
            snapshots = [(self.window.current_case.id, [selected.model_copy(deep=True)])]
            if selected.processed_confirmed and self.window.view_mode.currentIndex() == 1:
                snapshots[0][1][0].processed_confirmed = False

        tickets = {
            case_id: self.ocr_state.begin(
                next(c for c in cases if c.id == case_id), images, scope="roi" if roi else "full"
            )
            for case_id, images in snapshots
        }

        def run(cancel, progress):
            results = []
            service = RecognitionService(settings)
            for index, (case_id, images) in enumerate(snapshots):
                try:
                    value = service.run(
                        images,
                        cancel,
                        lambda p, t, i=index: progress(round((i * 100 + p) / len(snapshots)), t),
                        roi,
                        **({"document_type": forced_types[case_id]} if case_id in forced_types else {}),
                    )
                    value.run_id = tickets[case_id].id
                    results.append((case_id, value, ""))
                except AppError as exc:
                    if cancel.is_set():
                        raise
                    results.append((case_id, None, exc if hasattr(exc, "code") else str(exc)))
            return results

        self.window.undo_stack.clear()
        self.window.notice.setText("")
        self.ocr_previous_step = self.window.current_step
        self.ocr_view_case_id = self.window.current_case.id
        self.ocr_successful_case_ids = set()
        self.window.go_step(1)
        case_ids = {case_id for case_id, _images in snapshots}
        for case in cases:
            if case.id in case_ids:
                case.status = "Обрабатывается"
        self.window.refresh_cases()
        self.window.run_worker(run, lambda results: self.ocr_completed(results, tickets))
        self.ocr_case_ids = case_ids
        if self.window.active_worker:
            self.window.active_worker.signals.finished.connect(
                lambda: self.ocr_finished(tickets), Qt.ConnectionType.QueuedConnection
            )

    def ocr_finished(self, tickets=None):
        case_ids = (
            {c.id for c in self.window.batch.cases if c.id in tickets and c.active_run_id == tickets[c.id].id}
            if tickets
            else self.ocr_case_ids
        )
        if not case_ids:
            return
        for case in self.window.batch.cases:
            if case.id in case_ids and case.status == "Обрабатывается":
                if tickets and case.id in tickets:
                    self.ocr_state.stop(case, tickets[case.id].id, "CANCELLED")
                case.status = "На проверке" if case.evidence else "Фото загружены"
        self.window.refresh_cases()
        if self.window.current_case.id in case_ids:
            self.window.go_step(
                2
                if self.window.current_case.pending_identity
                or self.window.current_case.id in self.ocr_successful_case_ids
                else self.ocr_previous_step
            )
        elif self.window.current_case.id == getattr(self, "ocr_view_case_id", None):
            self.window.go_step(self.ocr_previous_step)

    def ocr_region(self, roi):
        if self.window.current_case.images:
            self.start_ocr(roi=roi)

    def repeat_field(self, name, candidate):
        """Re-read the selected evidence pixels, keeping manual values and clients isolated."""
        if name in {
            "surname_cyrillic",
            "surname_latin",
            "first_name_cyrillic",
            "first_name_latin",
            "patronymic_cyrillic",
        }:
            name = name.rsplit("_", 1)[0]
        if self.window.active_worker or candidate is None or candidate.source_bbox is None:
            return
        case = self.window.current_case
        if case.data.document_type == "kg_physical_id" and name in {"surname", "first_name", "patronymic"}:
            hint = case.field_regions.get(name, {})
            front = next(
                (s for s in case.images if s.id == hint.get("image_id") and s.role == DocumentSide.FRONT),
                None,
            )
            front = front or next((s for s in case.images if s.role == DocumentSide.FRONT), None)
            if front is None:
                self.window.field_detail.setText(
                    "Для повторного чтения ФИО добавьте лицевую сторону документа."
                )
                return
            bbox = hint.get("bbox") if hint.get("image_id") == front.id else None
            if bbox is None and candidate.source_image_id != front.id:
                self.window.field_detail.setText(
                    "Сначала распознайте лицевую сторону для поиска области ФИО."
                )
                return
            candidate = candidate.model_copy(
                update={
                    "source_image_id": front.id,
                    "source_side": DocumentSide.FRONT,
                    "source_bbox": bbox or candidate.source_bbox,
                    "image_version": hint.get("image_version", candidate.image_version),
                }
            )
        source = next(
            (s.model_copy(deep=True) for s in case.images if s.id == candidate.source_image_id), None
        )
        if source is None:
            return
        case_id, kind = case.id, case.data.document_type
        settings = self.window.settings.model_copy(deep=True)
        field_hint = case.field_regions.get(name, {})
        if candidate.image_version != "processed":
            source.processed_confirmed = False
        ticket = self.ocr_state.begin(
            case,
            [source],
            fields={name, "passport_type"} if name == "passport_number" else {name},
            scope="field",
        )

        def run(cancel, progress):
            if kind == "kg_digital_id":
                from document_excel.infrastructure.digital_reread import reread_digital_field
                from document_excel.infrastructure.ocr_factory import create_engine
                from document_excel.infrastructure.preprocessing import ImagePreprocessingService

                progress(10, "Повторное чтение выбранного цифрового поля")
                prepared = ImagePreprocessingService().prepare(
                    source.selected_path, "original", source.selected_rotation
                )
                hint = field_hint.get("bbox") if field_hint.get("image_id") == source.id else None
                engine = create_engine(settings)
                engine.cache_context = (source.case_id, source.id, "digital-field", name, ticket.id)
                values = reread_digital_field(
                    prepared.original, source, name, hint or candidate.source_bbox, engine, cancel
                )
                for value in values:
                    value.image_version = candidate.image_version
                progress(100, "Сверьте варианты поля")
                return values
            if kind == "kg_physical_id":
                import numpy as np

                from document_excel.infrastructure.ocr_contract import OcrResult
                from document_excel.infrastructure.ocr_factory import create_engine
                from document_excel.infrastructure.physical_id import PlasticCard
                from document_excel.infrastructure.preprocessing import ImagePreprocessingService
                from document_excel.infrastructure.region_ocr import RegionBasedOCR

                progress(10, "Повторное чтение выбранного поля")
                prepared = ImagePreprocessingService().prepare(
                    source.selected_path, "original", source.selected_rotation
                )
                pixels = prepared.original
                h, w = pixels.shape[:2]
                x, y, width, height = candidate.source_bbox
                if name in {"surname", "first_name", "patronymic"}:
                    # A previous word box may be truncated. Recover the whole
                    # printed group, retaining its full right-hand extent.
                    hint = field_hint.get("bbox") if field_hint.get("image_id") == source.id else None
                    if hint and field_hint.get("image_version") == candidate.image_version:
                        x, y, width, height = hint
                    else:
                        x = max(0, x - round(w * 0.03))
                        width = w - x
                left, top, right, bottom = max(0, x), max(0, y), min(w, x + width), min(h, y + height)
                if right <= left or bottom <= top:
                    raise AppError("Область поля вне изображения. Повторите OCR документа.")
                engine = create_engine(settings)
                engine.cache_context = (source.case_id, source.id, "manual-field", name, ticket.id)
                bbox = (left, top, right - left, bottom - top)
                if name in {"passport_issued_by", "passport_number"}:
                    values = RegionBasedOCR(engine).extract(
                        PlasticCard(pixels, np.eye(3), (w, h), False),
                        OcrResult([], source.id, "original", w, h),
                        cancel,
                        source.role,
                        field_name=name,
                        bbox=bbox,
                    )
                elif name in {"surname", "first_name", "patronymic"}:
                    from document_excel.infrastructure.full_name import FullNameExtractor, NameGroups

                    extractor = FullNameExtractor()
                    groups = NameGroups(regions={name: bbox})
                    values = extractor.recover(
                        pixels, OcrResult([], source.id, "original", w, h), groups, engine, cancel
                    ).candidates
                else:
                    from document_excel.domain.models import DATE_FIELDS, normalize, validate_field
                    from document_excel.infrastructure.field_extractors import DateExtractor

                    kind_value = (
                        "date"
                        if name in DATE_FIELDS
                        else "name"
                        if name in {"surname", "first_name", "patronymic"}
                        else "text"
                    )
                    read = engine.recognize_physical_value(
                        pixels[top:bottom, left:right],
                        source.id,
                        cancel,
                        "rus" if kind_value != "date" else "eng",
                        kind_value,
                    )
                    raw = read.text.strip()
                    try:
                        value = (
                            DateExtractor().value(name, raw) if name in DATE_FIELDS else normalize(name, raw)
                        )
                    except ValueError:
                        value = ""
                    values = []
                    if value:
                        error = validate_field(name, value)
                        values = [
                            FieldCandidate(
                                field_name=name,
                                raw_value=raw,
                                normalized_value=value,
                                confidence=min((b.confidence for b in read.blocks), default=0),
                                source_image_id=source.id,
                                source_bbox=bbox,
                                source_side=source.role,
                                extraction_method="selected-field-region",
                                validation_status=FieldStatus.INVALID if error else FieldStatus.REVIEW,
                                validation_errors=[error] if error else [],
                            )
                        ]
                for value in values:
                    value.image_version = candidate.image_version
                progress(100, "Сверьте варианты поля")
                return values
            output = RecognitionService(settings).run([source], cancel, progress, candidate.source_bbox)
            return [c for c in output.candidates if c.field_name == name]

        def completed(values):
            target = next((c for c in self.window.batch.cases if c.id == case_id), None)
            if target is None or not any(s.id == source.id for s in target.images):
                return
            if not self.ocr_state.current(target, ticket.id):
                from document_excel.application.ocr_state import event

                event(target, "STALE_RESULT_IGNORED", ticket.id, reason="field_run_changed")
                return
            if not values:
                self.ocr_state.stop(target, ticket.id, "FAILED")
                self.window.notice.setText(
                    "Поле не прочитано. Прежние варианты сохранены; исправьте вручную или переснимите."
                )
                return
            from document_excel.application.workflows import RecognitionOutput

            output = RecognitionOutput(values, kind, [], False, 0)
            self.ocr_state.commit(target, ticket.id, output)
            if self.window.current_case.id == case_id:
                self.window.refresh_form()
                self.window.focus_candidate(values[0])
                self.window.notice.setText("Поле перечитано. Откройте варианты и сверьте с фотографией.")

        self.window.run_worker(run, completed)
        if self.window.active_worker:
            worker = self.window.active_worker
            worker.signals.finished.connect(
                lambda: self.ocr_state.stop(
                    case, ticket.id, "CANCELLED" if worker.cancel_event.is_set() else "FAILED"
                ),
                Qt.ConnectionType.QueuedConnection,
            )

    def ocr_completed(self, results, tickets=None):
        current_succeeded = False
        failures = 0
        for case_id, output, error in results:
            case = next((c for c in self.window.batch.cases if c.id == case_id), None)
            if case is None:
                continue
            ticket = tickets.get(case_id) if tickets else None
            if not ticket and case.active_run_id:
                from document_excel.application.ocr_state import event

                event(case, "STALE_RESULT_IGNORED", reason="untracked_callback")
                continue
            if ticket and not self.ocr_state.current(case, ticket.id):
                from document_excel.application.ocr_state import event

                event(case, "STALE_RESULT_IGNORED", ticket.id, reason="run_or_image_changed")
                continue
            if error:
                if ticket:
                    self.ocr_state.stop(case, ticket.id, "FAILED")
                case.status = "Ошибка OCR"
                case.errors = [str(error)]
                case.processing_metrics.update(getattr(error, "metrics", {}))
                from document_excel.application.ocr_state import event

                event(
                    case,
                    "OCR_FAILED",
                    ticket.id if ticket else "",
                    status=getattr(error, "code", "OCR_ENGINE_ERROR"),
                )
                failures += 1
                continue
            if output is None:
                continue
            if ticket and not self.ocr_state.commit(case, ticket.id, output):
                continue
            if ticket and case.pending_identity:
                self.window.case_ocr_texts.setdefault(case.id, {}).update(getattr(output, "ocr_texts", {}))
                case.status = "Проверьте принадлежность фотографий"
                if hasattr(self, "ocr_successful_case_ids"):
                    self.ocr_successful_case_ids.add(case.id)
                current_succeeded |= case.id == self.window.current_case.id
                continue
            for image in case.images:
                side = getattr(output, "sides", {}).get(image.id)
                case.side_proposals.pop(image.id, None)
                if image.role == DocumentSide.UNKNOWN and side is not None:
                    image.role = DocumentSide(side)
                    if side == DocumentSide.UNKNOWN and output.document_type == "kg_physical_id":
                        case.side_proposals[image.id] = DocumentSide.UNKNOWN
                elif (
                    side in {DocumentSide.FRONT, DocumentSide.BACK}
                    and image.role in {DocumentSide.FRONT, DocumentSide.BACK}
                    and side != image.role
                    and output.document_type == "kg_physical_id"
                ):
                    case.side_proposals[image.id] = DocumentSide(side)
            texts = getattr(output, "ocr_texts", {})
            self.window.case_ocr_texts.setdefault(case.id, {}).update(texts)
            if not ticket:
                # Legacy synchronous test/import callers have no queued job.
                # An untracked callback cannot overwrite an active real run.
                if case.active_run_id:
                    continue
                case.mrz_reports.update(getattr(output, "mrz_reports", {}))
                apply_candidates(case, output.candidates, output.document_type, output.label_seen)
            case.status = "На проверке"
            case.errors = output.warnings
            if hasattr(self, "ocr_successful_case_ids"):
                self.ocr_successful_case_ids.add(case.id)
            current_succeeded |= case.id == self.window.current_case.id
        self.window.refresh_cases()
        self.window.refresh_photos()
        self.window.refresh_form()
        self.window.notice.setText(
            "Данные извлечены. Проверьте выделенные поля."
            if current_succeeded
            else "Некоторые документы не распознаны. Проверьте статусы в очереди."
            if failures
            else "Распознавание очереди завершено. Перейдите к проверке клиентов."
        )
        self.window.record_history("OCR", "завершено; проверьте поля")
        if current_succeeded:
            self.window.go_step(2)
            if self.window.current_case.pending_identity:
                self.window.notice.setText(
                    "Текст распознан. Подтвердите принадлежность фотографий: найденные поля доступны для просмотра."
                )
        if (
            current_succeeded
            and self.window.current_case.workflow_mode != WorkflowMode.BEST_EFFORT_UPDATE
            and not self.window.current_case.pending_identity
            and not self.window.current_case.data.registration_address.display()
            and self.window.current_case.registration_status != RegistrationStatus.NOT_PRESENT
        ):
            QTimer.singleShot(0, self.window, lambda: self.address_dialog(True, after_ocr=True))

    def review_identity(self):
        case = self.window.current_case
        if self.window.active_worker:
            return
        dialog = QDialog(self.window)
        dialog.setWindowTitle("Принадлежность фотографий")
        layout = QVBoxLayout(dialog)
        message = QLabel(
            "Не удалось надёжно связать стороны по реквизитам. Убедитесь, что все фотографии принадлежат одному человеку. До подтверждения результаты не объединяются."
        )
        message.setWordWrap(True)
        layout.addWidget(message)
        row = QHBoxLayout()
        for index, image in enumerate(case.images):
            preview = QLabel(f"Фото {index + 1}")
            picture = QPixmap(str(image.selected_path)).scaled(
                250, 170, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation
            )
            preview.setPixmap(picture)
            row.addWidget(preview)
        layout.addLayout(row)
        preview_case = self.ocr_state.preview(case)
        if preview_case:
            from PySide6.QtWidgets import QTextEdit

            from document_excel.domain.models import FIELD_LABELS, data_value

            details = QTextEdit()
            details.setReadOnly(True)
            details.setPlainText(
                "\n".join(
                    f"{label}: {data_value(preview_case.data, name)}"
                    for name, label in FIELD_LABELS.items()
                    if data_value(preview_case.data, name)
                )
            )
            layout.addWidget(details)
        accept = QPushButton("Подтверждаю: один человек")
        accept.clicked.connect(dialog.accept)
        layout.addWidget(accept)
        cancel = QPushButton("Проверить и разнести по клиентам")
        cancel.clicked.connect(dialog.reject)
        layout.addWidget(cancel)
        if dialog.exec():
            if not self.ocr_state.accept_identity(case):
                self.window.notice.setText("Фотографии изменились. Повторите распознавание.")
            else:
                self.window.notice.setText(
                    "Принадлежность фотографий подтверждена. Сверьте реквизиты и адрес."
                )
                self.window.go_step(2)
            self.window.refresh_form()
            self.window.refresh_photos()

    def review_sides(self):
        case = self.window.current_case
        if not case.side_proposals or self.window.active_worker:
            return
        dialog = QDialog(self.window)
        dialog.setWindowTitle("Проверить стороны документа")
        box = QVBoxLayout(dialog)
        box.addWidget(QLabel("Распознанные признаки отличаются от выбранных сторон. Проверьте фотографии."))
        choices = {}
        for index, source in enumerate(case.images):
            proposed = case.side_proposals.get(source.id)
            if proposed:
                row = QHBoxLayout()
                preview = QLabel()
                picture = QPixmap(str(source.selected_path)).transformed(
                    QTransform().rotate(source.selected_rotation)
                )
                preview.setPixmap(
                    picture.scaled(
                        170,
                        110,
                        Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation,
                    )
                )
                row.addWidget(preview)
                row.addWidget(
                    QLabel(f"Фото {index + 1}: выбрано {source.role.value}, определено {proposed.value}")
                )
                control = QComboBox()
                for label, value in [
                    ("Выберите сторону", DocumentSide.UNKNOWN),
                    ("Лицевая", DocumentSide.FRONT),
                    ("Оборотная", DocumentSide.BACK),
                    ("Дополнительный документ", DocumentSide.ADDITIONAL),
                ]:
                    control.addItem(label, value)
                control.setCurrentIndex(max(0, control.findData(proposed)))
                choices[source.id] = control
                row.addWidget(control)
                box.addLayout(row)
        row = QHBoxLayout()
        for text, decision in [
            ("Применить выбранные стороны", 1),
            ("Оставить выбранные стороны", 2),
            ("Отмена", 0),
        ]:
            button = QPushButton(text)
            button.clicked.connect(lambda _checked=False, value=decision: dialog.done(value))
            row.addWidget(button)
        box.addLayout(row)
        decision = dialog.exec()
        if decision:
            if decision == 1:
                for source in case.images:
                    if source.id in choices:
                        source.role = DocumentSide(choices[source.id].currentData())
            case.side_proposals.clear()
            for source in case.images:
                if source.role == DocumentSide.UNKNOWN:
                    case.side_proposals[source.id] = DocumentSide.UNKNOWN
            case.invalidate()
            self.window.refresh_photos()
            self.window.refresh_form()
        dialog.deleteLater()

    def show_mrz(self):
        case = self.window.current_case
        dialog = QDialog(self.window)
        dialog.setWindowTitle("Проверка MRZ · сведения текущего клиента")
        dialog.resize(760, 540)
        box = QVBoxLayout(dialog)
        text = QPlainTextEdit()
        text.setReadOnly(True)
        sections = []
        for report in case.mrz_reports.values():
            sections.append(
                "TD1 · контрольные цифры сошлись" if report.valid else "TD1 · проверка не пройдена"
            )
            sections.extend(report.lines)
            sections.extend(f"{name}: {'OK' if valid else 'ОШИБКА'}" for name, valid in report.checks.items())
            sections.extend(
                f"{name}: НЕ ПРОВЕРЕНО"
                for name, status in report.segment_status.items()
                if status == "NOT_CHECKED"
            )
            sections.extend(report.errors + report.warnings)
            sections.extend("Расхождение: " + FIELD_LABELS.get(name, name) for name in report.conflicts)
            sections.append(
                "Контрольные цифры не подтверждают подлинность документа. Сверьте печатные данные.\n"
            )
        text.setPlainText(
            "\n".join(sections) if sections else "MRZ ещё не распознана. Проверьте оборот и запустите OCR."
        )
        box.addWidget(text)
        row = QHBoxLayout()
        area = QPushButton("Показать область MRZ")

        def focus():
            report = next(iter(case.mrz_reports.values()), None)
            if report:
                candidate = FieldCandidate(
                    field_name="passport_number",
                    raw_value="",
                    normalized_value="",
                    confidence=0,
                    source_image_id=report.image_id,
                    source_bbox=report.source_bbox,
                    image_version=report.image_version,
                )
                self.window.focus_candidate(candidate)
            dialog.accept()

        area.clicked.connect(focus)
        row.addWidget(area)
        repeat = QPushButton("Повторить OCR")
        repeat.clicked.connect(lambda: dialog.done(2))
        row.addWidget(repeat)
        close = QPushButton("Закрыть")
        close.clicked.connect(dialog.reject)
        row.addWidget(close)
        box.addLayout(row)
        decision = dialog.exec()
        dialog.deleteLater()
        if decision == 2 and self.window.current_case is case:
            QTimer.singleShot(0, self.window, self.start_ocr)

    def change_document_type(self):
        control = self.sender()
        selected_type = (
            control.currentData()
            if hasattr(control, "currentData")
            else self.window.document_type.currentData()
        )
        self.window.commit_edits()
        before = self.window.snapshot_edit()
        self.window.current_case.data.document_type = selected_type
        self.window.current_case.document_type_manual = selected_type != "unknown"
        self.window.current_case.invalidate()
        self.window.push_snapshot(before, "Тип документа")
        self.window.refresh_form()

    def address_dialog(self, registration=True, after_ocr=False):
        if self.window.active_worker and not after_ocr:
            return
        self.window.commit_edits()
        before = self.window.snapshot_edit()
        field = "registration_address" if registration else "living_address"
        dialog = AddressDialog(getattr(self.window.current_case.data, field), registration, self.window)
        if dialog.exec():
            setattr(self.window.current_case.data, field, dialog.address())
            if registration:
                self.window.current_case.registration_status = (
                    RegistrationStatus.NOT_PRESENT
                    if dialog.absent.isChecked()
                    else RegistrationStatus.MANUALLY_ENTERED
                )
            self.window.current_case.evidence[field] = FieldEvidence(
                source="ручной ввод",
                status=FieldStatus.MANUALLY_CORRECTED,
                message="Отсутствует в документе"
                if dialog.absent.isChecked()
                else "Адрес проверен пользователем",
            )
            self.window.current_case.invalidate()
            self.window.push_snapshot(before, "Адрес регистрации" if registration else "Фактический адрес")
            self.window.on_edited(field)

    def edit_address_field(self, index):
        name = self.window.field_model.fields[index.row()]
        if name in {"registration_address", "living_address"}:
            self.address_dialog(name == "registration_address")

    def contract_dialog(self):
        if self.window.active_worker:
            return
        self.window.commit_edits()
        before = self.window.snapshot_edit()
        dialog = ContractDialog(self.window.current_case.contract, self.window)
        if dialog.exec():
            self.window.current_case.contract = dialog.value
            self.window.current_case.invalidate()
            self.window.push_snapshot(before, "Данные договора")
            self.window.refresh_form()

    def focus_field(self, index):
        field = self.window.field_model.fields[index.row()]
        evidence = self.window.current_case.evidence.get(field, FieldEvidence())
        self.window.field_detail.setText(evidence.message or "Сверьте значение с изображением.")
        self.window.alternatives.clear()
        for candidate in evidence.candidates:
            self.window.alternatives.addItem(
                candidate.raw_value + f" · {candidate.confidence:.2f} (эвристика)", candidate
            )
        from document_excel.domain.resolution import displayed_candidate

        candidate = displayed_candidate(evidence, getattr(self.window.current_case.data, field, None))
        if candidate:
            for i, source in enumerate(self.window.current_case.images):
                if source.id == candidate.source_image_id:
                    self.window.image_list.setCurrentIndex(i)
                    self.window.viewer.highlight(candidate.source_bbox)
                    break

    def accept_candidate(self):
        if self.window.active_worker or self.window.alternatives.currentIndex() < 0:
            return
        candidate = self.window.alternatives.currentData()
        self.window.commit_edits()
        before = self.window.snapshot_edit()
        if candidate.field_name == "registration_address":
            from document_excel.domain.models import Address

            if self.window.current_case.data.document_type == "kg_digital_id":
                from document_excel.domain.digital_address import digital_address_from_text

                self.window.current_case.data.registration_address = digital_address_from_text(
                    candidate.normalized_value
                )
            else:
                self.window.current_case.data.registration_address = Address(raw=candidate.normalized_value)
            self.window.current_case.registration_status = RegistrationStatus.USER_CONFIRMED
            self.window.current_case.evidence[candidate.field_name] = FieldEvidence(
                source="ручной выбор OCR",
                status=FieldStatus.MANUALLY_CORRECTED,
                candidates=self.window.current_case.evidence[candidate.field_name].candidates,
            )
            self.window.current_case.invalidate()
        else:
            self.window.apply_field_edit(candidate.field_name, candidate.normalized_value)
        if candidate.field_name == "registration_address":
            self.window.push_snapshot(before, "Выбор варианта адреса OCR")
        self.window.on_edited(candidate.field_name)

    def copy_field(self):
        index = self.window.fields.currentIndex()
        if index.isValid():
            QApplication.clipboard().setText(
                data_value(self.window.current_case.data, self.window.field_model.fields[index.row()])
            )

    def confirm_data(self):
        if self.window.active_worker:
            return
        if self.window.current_case.pending_identity:
            self.review_identity()
            if self.window.current_case.pending_identity:
                return
        self.window.commit_edits()
        case = self.window.current_case
        if case.workflow_mode == WorkflowMode.BEST_EFFORT_UPDATE:
            self.window.excel_actions.export_current()
            return
        old_status = case.registration_status
        if (
            case.data.registration_address.display()
            and case.registration_status != RegistrationStatus.NOT_PRESENT
        ):
            case.registration_status = RegistrationStatus.USER_CONFIRMED
        errors = readiness(case, self.window.profile, require_confirmation=False)
        if errors:
            case.registration_status = old_status
            self.window.error("\n".join(errors))
            return
        if not ConfirmDialog.ask(
            self.window,
            "Сверка с документом",
            "Подтверждаете ФИО, даты, ПИН, номер документа и адрес после проверки оригиналов?",
        ):
            case.registration_status = old_status
            return
        for evidence in case.evidence.values():
            if evidence.status not in {
                FieldStatus.MISSING,
                FieldStatus.INVALID,
                FieldStatus.CONFLICT,
                FieldStatus.NOT_PRESENT,
                FieldStatus.NOT_RECOGNIZED,
            }:
                evidence.status = FieldStatus.CONFIRMED
        from document_excel.application.ocr_state import event

        event(
            case,
            "FIELD_CONFIRMED",
            case.active_run_id or "",
            count=sum(e.status == FieldStatus.CONFIRMED for e in case.evidence.values()),
        )
        case.confirmed_digest = case.digest()
        case.status = "Подтверждено"
        self.window.refresh_form()
        self.window.refresh_cases()
        self.window.go_step(3)
        if case.workflow_mode.updates_existing:
            self.window.go_step(2)
            QTimer.singleShot(0, self.window, self.window.excel_actions.export_current)

    def clear_workspace(self):
        if self.window.active_worker:
            return
        self.window.commit_edits()
        if (
            QMessageBox.question(
                self.window,
                "Очистить рабочие данные",
                "Удалить все записи клиентов из памяти? Сохранённые Excel и копии фото останутся в выбранных папках.",
            )
            == QMessageBox.StandardButton.Yes
        ):
            from document_excel.application.fast_recognition import clear_cache
            from document_excel.infrastructure.paddle_ocr import clear_cache as clear_inference_cache

            clear_cache()
            clear_inference_cache()
            self.window.undo_stack.clear()
            self.window.case_ocr_texts.clear()
            self.window.last_result = None
            self.window.batch = DocumentBatch(
                cases=[ClientCase(name="Клиент 1", workflow_mode=WorkflowMode.BEST_EFFORT_UPDATE)]
            )
            self.window.select_case(0)
            self.window.case_raw_edits.clear()
            self.window.case_steps.clear()
            self.window.refresh_cases()
            self.window.go_step(0)
            QApplication.clipboard().clear()
