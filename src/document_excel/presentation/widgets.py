from __future__ import annotations

from pathlib import Path

import cv2
from PySide6.QtCore import QAbstractTableModel, QModelIndex, QRectF, Qt, Signal
from PySide6.QtGui import QColor, QImage, QPainter, QPen, QPixmap
from PySide6.QtWidgets import QGraphicsScene, QGraphicsView, QHeaderView, QTableView

from document_excel.domain.models import (
    DATE_FIELDS,
    FIELD_LABELS,
    Address,
    ClientCase,
    FieldEvidence,
    FieldStatus,
    data_value,
    normalize,
    parse_date,
    split_document_identifier,
    split_mkk_authority,
    validate_field,
)
from document_excel.infrastructure.excel import ExportPlan
from document_excel.infrastructure.preprocessing import detect_document_region, load_image

STATUS_TEXT = {
    FieldStatus.MISSING: "[ ] Нет данных",
    FieldStatus.NOT_RECOGNIZED: "[?] Не распознано",
    FieldStatus.NOT_PRESENT: "[—] Не предусмотрено",
    FieldStatus.INVALID: "! Ошибка",
    FieldStatus.REVIEW: "[?] Проверить",
    FieldStatus.VALID: "[OK] Проверка пройдена",
    FieldStatus.CONFLICT: "[!] Конфликт",
    FieldStatus.MANUALLY_CORRECTED: "[M] Введено вручную",
    FieldStatus.CONFIRMED: "[OK] Подтверждено",
}


class FieldTableModel(QAbstractTableModel):
    edited = Signal(str)

    def __init__(self, case: ClientCase, pin_length: int = 14, identifier_label: str = "ПИН / ИНН") -> None:
        super().__init__()
        self.case = case
        self.fields = list(FIELD_LABELS)
        self.pin_length = pin_length
        self.identifier_label = identifier_label
        self.raw_edits: dict[str, str] = {}

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.fields)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else 3

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole and orientation == Qt.Orientation.Horizontal:
            return ["Поле", "Значение", "Проверка / источник"][section]
        return None

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        field = self.fields[index.row()]
        value = self.raw_edits.get(field, data_value(self.case.data, field))
        evidence = self.case.evidence.get(field, FieldEvidence())
        if role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.EditRole):
            return [
                self.identifier_label if field == "personal_identifier" else FIELD_LABELS[field],
                value,
                STATUS_TEXT[evidence.status] + " · " + evidence.source,
            ][index.column()]
        if role == Qt.ItemDataRole.ToolTipRole:
            alternatives = "\n".join(
                f"{c.raw_value} · {c.variant} · {c.confidence:.2f} (эвристика)" for c in evidence.candidates
            )
            return (
                "\n".join(filter(None, [evidence.message, alternatives])) or "Значение можно ввести вручную."
            )
        if role == Qt.ItemDataRole.BackgroundRole and index.column() == 2:
            if evidence.status in {FieldStatus.INVALID, FieldStatus.MISSING}:
                return QColor("#fce5e2")
            if evidence.status in {FieldStatus.REVIEW, FieldStatus.CONFLICT}:
                return QColor("#fff0c2")
            return QColor("#dff3e8")
        if role == Qt.ItemDataRole.ForegroundRole and index.column() == 2:
            return QColor("#223344")
        return None

    def flags(self, index):
        flags = super().flags(index)
        if index.column() == 1 and self.fields[index.row()] not in {"registration_address", "living_address"}:
            flags |= Qt.ItemFlag.ItemIsEditable
        return flags

    def setData(self, index, value, role=Qt.ItemDataRole.EditRole):
        if role != Qt.ItemDataRole.EditRole or index.column() != 1:
            return False
        field = self.fields[index.row()]
        value = str(value).strip()
        if field == "passport_number" and (parts := split_document_identifier(value)):
            prefix, value = parts
            self.case.data.passport_type = prefix
            type_evidence = self.case.evidence.get("passport_type", FieldEvidence()).model_copy(deep=True)
            type_evidence.source = "ручной ввод"
            type_evidence.status = FieldStatus.MANUALLY_CORRECTED
            type_evidence.message = "Буквенный префикс отделён от номера, введённого пользователем."
            self.case.evidence["passport_type"] = type_evidence
        if field == "passport_issued_by" and (authority_parts := split_mkk_authority(value)):
            value, code = authority_parts
            self.case.data.passport_department_code = code
            code_evidence = self.case.evidence.get("passport_department_code", FieldEvidence()).model_copy(
                deep=True
            )
            code_evidence.source = "ручной ввод"
            code_evidence.status = FieldStatus.MANUALLY_CORRECTED
            code_evidence.message = "Цифры после МКК отделены от органа выдачи, введённого пользователем."
            self.case.evidence["passport_department_code"] = code_evidence
        error = validate_field(field, value, self.pin_length) if value else None
        # Invalid input is visible and blocks export; the last valid model value is cleared.
        evidence = self.case.evidence.get(field, FieldEvidence()).model_copy(deep=True)
        evidence.source = "ручной ввод"
        evidence.status = FieldStatus.INVALID if error else FieldStatus.MANUALLY_CORRECTED
        evidence.message = error or "Сверьте с оригиналом перед подтверждением."
        self.raw_edits.pop(field, None)
        if field in DATE_FIELDS:
            setattr(self.case.data, field, parse_date(value) if value and not error else None)
            if error:
                self.raw_edits[field] = value
                evidence.message = f"{value}: {error}"
        elif field not in {"registration_address", "living_address"}:
            if error and field in {"gender", "citizenship", "surname", "first_name", "patronymic"}:
                setattr(self.case.data, field, None)
                self.raw_edits[field] = value
            else:
                setattr(self.case.data, field, normalize(field, value) if value else None)
        else:
            setattr(self.case.data, field, Address(raw=value))
        self.case.evidence[field] = evidence
        name_field = field
        if field in {"surname", "first_name", "patronymic"}:
            from document_excel.infrastructure.extraction import name_script

            script = name_script(value)
            target = f"{field}_{script}"
            if target in type(self.case.data).model_fields:
                setattr(self.case.data, target, value or None)
                setattr(evidence, "original_" + script, value or None)
                self.case.evidence[target] = FieldEvidence(source="ручной ввод", status=evidence.status)
        elif field.endswith(("_cyrillic", "_latin")):
            name_field, script = field.rsplit("_", 1)
            base = self.case.evidence.get(name_field, FieldEvidence()).model_copy(deep=True)
            base.source = "ручной ввод"
            base.status = evidence.status
            setattr(base, "original_" + script, value or None)
            self.case.evidence[name_field] = base
            if script == "cyrillic":
                setattr(self.case.data, name_field, value or None)
            elif not getattr(self.case.data, name_field):
                setattr(
                    self.case.data,
                    name_field,
                    getattr(self.case.data, name_field + "_cyrillic", None) or value or None,
                )
        if name_field in {"surname", "first_name", "patronymic"}:
            full = self.case.evidence.get("full_name", FieldEvidence())
            if not full.source.startswith("ручной"):
                self.case.data.full_name = (
                    " ".join(
                        filter(
                            None,
                            [self.case.data.surname, self.case.data.first_name, self.case.data.patronymic],
                        )
                    )
                    or None
                )
                self.case.evidence["full_name"] = FieldEvidence(
                    source="составлено из частей имени", status=FieldStatus.REVIEW
                )
        self.case.invalidate()
        self.refresh()
        self.edited.emit(field)
        return True

    def refresh(self):
        self.dataChanged.emit(self.index(0, 0), self.index(len(self.fields) - 1, 2))


class PreviewModel(QAbstractTableModel):
    def __init__(self, plan: ExportPlan) -> None:
        super().__init__()
        self.changes = plan.changes

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.changes)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else 6

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole and orientation == Qt.Orientation.Horizontal:
            return ["Лист", "Ячейка", "Сейчас", "После записи", "Источник", "Предупреждение"][section]
        return None

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if index.isValid() and role == Qt.ItemDataRole.DisplayRole:
            c = self.changes[index.row()]
            value = [c.sheet, c.cell, c.old, c.new, c.source, c.warning][index.column()]
            return "" if value is None else str(value)
        return None


def configure_table(table: QTableView) -> None:
    table.setAlternatingRowColors(True)
    table.setWordWrap(False)
    table.verticalHeader().hide()
    table.verticalHeader().setDefaultSectionSize(38)
    table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
    table.horizontalHeader().setStretchLastSection(True)
    table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)


class ImageViewer(QGraphicsView):
    files_dropped = Signal(list)
    region_selected = Signal(object)
    zoom_changed = Signal(int)

    def __init__(self) -> None:
        super().__init__()
        self.setScene(QGraphicsScene(self))
        self.setAcceptDrops(True)
        self.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        self.setBackgroundBrush(QColor("#e5eaf1"))
        self.image_item = None
        self.highlight_item = None
        self.selecting = False
        self.roi_start = None
        self.image_width = 0
        self.image_height = 0
        self.array = None
        self.detection = None
        self.mode = 0

    def show_path(self, path: Path, rotation: int = 0) -> None:
        array = load_image(path, rotation)
        h, w = array.shape[:2]
        if max(h, w) > 2600:
            array = cv2.resize(array, (round(w * 2600 / max(h, w)), round(h * 2600 / max(h, w))))
        self.set_image(array, detect_document_region(array))

    def set_image(self, array, detection=None):
        self.array = array
        self.detection = detection
        self.set_mode(self.mode)

    def set_mode(self, mode):
        self.mode = mode
        if self.array is None:
            return
        array = self.array
        h, w = array.shape[:2]
        x = y = 0
        if mode != 1 and self.detection and self.detection.detected:
            x, y, cw, ch = self.detection.bbox
            array = array[y : y + ch, x : x + cw].copy()
        if mode == 2:
            lab = cv2.cvtColor(array, cv2.COLOR_RGB2LAB)
            lab[:, :, 0] = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(lab[:, :, 0])
            array = cv2.cvtColor(lab, cv2.COLOR_LAB2RGB)
        ah, aw = array.shape[:2]
        qimage = QImage(array.data, aw, ah, array.strides[0], QImage.Format.Format_RGB888).copy()
        self.scene().clear()
        self.highlight_item = None
        self.image_item = self.scene().addPixmap(QPixmap.fromImage(qimage))
        self.image_item.setPos(x, y)
        self.image_width, self.image_height = w, h
        self.scene().setSceneRect(QRectF(x, y, aw, ah))
        self.fit()

    def clear(self) -> None:
        self.scene().clear()
        self.image_item = None
        self.highlight_item = None
        self.image_width = self.image_height = 0
        self.array = None

    def fit(self) -> None:
        if self.image_item:
            self.fitInView(self.sceneRect(), Qt.AspectRatioMode.KeepAspectRatio)
            self.zoom_changed.emit(round(self.transform().m11() * 100))

    def fit_document(self):
        self.fit()

    def fit_width(self):
        if self.image_item:
            factor = max(0.03, (self.viewport().width() - 12) / self.sceneRect().width())
            self.resetTransform()
            self.scale(factor, factor)
            self.centerOn(self.sceneRect().center())
            self.zoom_changed.emit(round(factor * 100))

    def zoom(self, factor):
        if self.image_item and 0.03 <= self.transform().m11() * factor <= 12:
            self.scale(factor, factor)
            self.zoom_changed.emit(round(self.transform().m11() * 100))

    def highlight(self, bbox) -> None:
        if self.highlight_item:
            self.scene().removeItem(self.highlight_item)
            self.highlight_item = None
        if bbox:
            self.highlight_item = self.scene().addRect(QRectF(*bbox), QPen(QColor("#d98c00"), 4))
            self.ensureVisible(self.highlight_item, 40, 40)

    def wheelEvent(self, event):
        self.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        if (
            0.03 < self.transform().m11() < 12
            or (self.transform().m11() >= 12 and event.angleDelta().y() < 0)
            or (self.transform().m11() <= 0.03 and event.angleDelta().y() > 0)
        ):
            self.zoom(1.2 if event.angleDelta().y() > 0 else 1 / 1.2)
        event.accept()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        paths = [Path(url.toLocalFile()) for url in event.mimeData().urls() if url.isLocalFile()]
        self.files_dropped.emit(paths)
        event.acceptProposedAction()

    def select_region(self) -> None:
        self.selecting = True
        self.setDragMode(QGraphicsView.DragMode.NoDrag)
        self.setCursor(Qt.CursorShape.CrossCursor)

    def cancel_selection(self) -> None:
        self.selecting = False
        self.roi_start = None
        self.setCursor(Qt.CursorShape.ArrowCursor)
        self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        self.highlight(None)

    def mousePressEvent(self, event):
        if self.selecting and event.button() == Qt.MouseButton.LeftButton:
            self.roi_start = self.mapToScene(event.position().toPoint())
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.selecting and self.roi_start:
            rect = (
                QRectF(self.roi_start, self.mapToScene(event.position().toPoint()))
                .normalized()
                .intersected(self.sceneRect())
            )
            self.highlight((rect.x(), rect.y(), rect.width(), rect.height()))
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self.selecting and self.roi_start:
            rect = (
                QRectF(self.roi_start, self.mapToScene(event.position().toPoint()))
                .normalized()
                .intersected(self.sceneRect())
            )
            self.selecting = False
            self.roi_start = None
            self.setCursor(Qt.CursorShape.ArrowCursor)
            self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
            if rect.width() >= 20 and rect.height() >= 20:
                self.region_selected.emit(
                    tuple(int(v) for v in (rect.x(), rect.y(), rect.width(), rect.height()))
                )
        else:
            super().mouseReleaseEvent(event)
