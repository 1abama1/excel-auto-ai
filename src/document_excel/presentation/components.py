from __future__ import annotations

from PySide6.QtCore import QEvent, QPointF, QRectF, QSize, Qt, Signal
from PySide6.QtGui import QAction, QColor, QIcon, QPainter, QPen, QPixmap, QPolygonF
from PySide6.QtWidgets import (
    QCalendarWidget,
    QComboBox,
    QDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from document_excel.domain.models import DATE_FIELDS, FIELD_LABELS, FieldEvidence, FieldStatus, data_value


def icon(name: str, color: str = "#64748B", size: int = 20) -> QIcon:
    pixmap = QPixmap(size * 2, size * 2)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.scale(size * 2 / 24, size * 2 / 24)
    painter.setPen(
        QPen(QColor(color), 1.7, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
    )
    if name in {"document", "upload", "scan"}:
        painter.drawRoundedRect(QRectF(5, 3, 14, 18), 2, 2)
        for y in (9, 13, 17):
            painter.drawLine(9, y, 15, y)
        if name == "upload":
            painter.drawLine(12, 14, 12, 6)
            painter.drawLine(9, 9, 12, 6)
            painter.drawLine(15, 9, 12, 6)
    elif name in {"check", "success"}:
        painter.drawEllipse(QRectF(3, 3, 18, 18))
        painter.drawPolyline(QPolygonF([QPointF(7, 12), QPointF(10, 15), QPointF(17, 8)]))
    elif name in {"queue", "history"}:
        if name == "history":
            painter.drawEllipse(QRectF(3, 3, 18, 18))
            painter.drawPolyline(QPolygonF([QPointF(12, 7), QPointF(12, 12), QPointF(16, 14)]))
        else:
            for y in (6, 12, 18):
                painter.drawEllipse(QRectF(3, y - 1, 2, 2))
                painter.drawLine(9, y, 21, y)
    elif name in {"settings", "templates"}:
        for x in (5, 12, 19):
            painter.drawLine(x, 4, x, 20)
            painter.drawRoundedRect(QRectF(x - 2, 7 if x != 12 else 13, 4, 4), 1, 1)
    elif name in {"plus", "minus", "close"}:
        if name == "close":
            painter.drawLine(6, 6, 18, 18)
            painter.drawLine(18, 6, 6, 18)
        else:
            painter.drawLine(6, 12, 18, 12)
            if name == "plus":
                painter.drawLine(12, 6, 12, 18)
    elif name == "more":
        for x in (5, 12, 19):
            painter.drawEllipse(QRectF(x - 1, 11, 2, 2))
    elif name in {"back", "next"}:
        points = [(14, 6), (8, 12), (14, 18)] if name == "back" else [(10, 6), (16, 12), (10, 18)]
        painter.drawPolyline(QPolygonF([QPointF(x, y) for x, y in points]))
    elif name == "folder":
        painter.drawPolyline(
            QPolygonF(
                [
                    QPointF(3, 19),
                    QPointF(3, 6),
                    QPointF(10, 6),
                    QPointF(12, 9),
                    QPointF(21, 9),
                    QPointF(21, 19),
                    QPointF(3, 19),
                ]
            )
        )
    elif name in {"fit", "region", "width"}:
        for x, y, sx, sy in [(4, 4, 1, 1), (20, 4, -1, 1), (4, 20, 1, -1), (20, 20, -1, -1)]:
            painter.drawLine(x, y, x + sx * 5, y)
            painter.drawLine(x, y, x, y + sy * 5)
        if name == "width":
            painter.drawLine(7, 12, 17, 12)
    elif name == "rotate":
        painter.drawArc(QRectF(4, 4, 16, 16), 30 * 16, 290 * 16)
        painter.drawPolyline(QPolygonF([QPointF(16, 2), QPointF(20, 6), QPointF(15, 7)]))
    elif name == "source":
        painter.drawEllipse(QRectF(6, 6, 12, 12))
        painter.drawLine(12, 2, 12, 8)
        painter.drawLine(12, 16, 12, 22)
        painter.drawLine(2, 12, 8, 12)
        painter.drawLine(16, 12, 22, 12)
    elif name == "calendar":
        painter.drawRoundedRect(QRectF(4, 5, 16, 16), 2, 2)
        painter.drawLine(4, 10, 20, 10)
        painter.drawLine(8, 3, 8, 7)
        painter.drawLine(16, 3, 16, 7)
    else:
        painter.drawEllipse(QRectF(3, 3, 18, 18))
        painter.drawLine(12, 10, 12, 16)
        painter.drawPoint(12, 7)
    painter.end()
    pixmap.setDevicePixelRatio(2)
    return QIcon(pixmap)


class PrimaryButton(QPushButton):
    def __init__(self, text: str, callback=None):
        super().__init__(text)
        self.setObjectName("primary")
        self.setMinimumHeight(42)
        if callback:
            self.clicked.connect(callback)


class SecondaryButton(QPushButton):
    def __init__(self, text: str, callback=None, icon_name: str | None = None):
        super().__init__(text)
        self.setMinimumHeight(38)
        if icon_name:
            self.setIcon(icon(icon_name))
        if callback:
            self.clicked.connect(callback)


class IconButton(QPushButton):
    def __init__(self, name: str, tooltip: str, callback=None):
        super().__init__()
        self.setObjectName("iconButton")
        self.setIcon(icon(name))
        self.setFixedSize(34, 34)
        self.setIconSize(QSize(18, 18))
        self.setToolTip(tooltip)
        self.setAccessibleName(tooltip)
        if callback:
            self.clicked.connect(callback)


class TextField(QLineEdit):
    def __init__(self, placeholder: str = ""):
        super().__init__()
        self.setPlaceholderText(placeholder)
        self.setMinimumHeight(40)
        self.setMaxLength(2000)


class DateField(TextField):
    def __init__(self):
        super().__init__("ДД.ММ.ГГГГ")
        action = QAction(icon("calendar"), "Выбрать дату", self)
        action.triggered.connect(self.show_calendar)
        self.addAction(action, QLineEdit.ActionPosition.TrailingPosition)

    def show_calendar(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Выберите дату")
        layout = QVBoxLayout(dialog)
        calendar = QCalendarWidget()
        calendar.setGridVisible(False)
        layout.addWidget(calendar)

        def choose(value):
            self.setText(value.toString("dd.MM.yyyy"))
            self.editingFinished.emit()
            dialog.accept()

        calendar.clicked.connect(choose)
        dialog.exec()


class StatusIndicator(QLabel):
    def __init__(self):
        super().__init__()
        self.setObjectName("statusIndicator")
        self.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed)

    def update_status(self, status: FieldStatus, required: bool = False):
        text, tone = {
            FieldStatus.MISSING: ("Нет данных", "error" if required else "muted"),
            FieldStatus.NOT_RECOGNIZED: ("Не распознано", "warning"),
            FieldStatus.NOT_PRESENT: ("Не предусмотрено", "muted"),
            FieldStatus.INVALID: ("Ошибка", "error"),
            FieldStatus.REVIEW: ("Проверить", "warning"),
            FieldStatus.VALID: ("Найдено", "success"),
            FieldStatus.CONFLICT: ("Конфликт", "error"),
            FieldStatus.MANUALLY_CORRECTED: ("Вручную", "info"),
            FieldStatus.CONFIRMED: ("Подтверждено", "success"),
        }[status]
        self.setText(text)
        self.setProperty("tone", tone)
        self.style().unpolish(self)
        self.style().polish(self)


class Notification(QFrame):
    def __init__(self):
        super().__init__()
        self.setObjectName("notification")
        row = QHBoxLayout(self)
        row.setContentsMargins(14, 9, 14, 9)
        self.label = QLabel()
        self.label.setWordWrap(True)
        self.label.setTextFormat(Qt.TextFormat.PlainText)
        row.addWidget(self.label, 1)
        self.setVisible(False)

    def setText(self, text: str):
        self.label.setText(text)
        self.setVisible(bool(text))

    def text(self) -> str:
        return self.label.text()


class StepIndicator(QWidget):
    selected = Signal(int)
    TITLES = ["Договор и паспорт", "Распознавание", "Проверка данных", "Excel", "Готовый договор"]

    def __init__(self):
        super().__init__()
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)
        self.buttons = []
        for i, title in enumerate(self.TITLES):
            number = {0: 1, 2: 2, 4: 3}.get(i, i + 1)
            button = QPushButton(f"{number}  {title}")
            button.setObjectName("step")
            button.setMinimumHeight(36)
            button.clicked.connect(lambda _checked=False, index=i: self.selected.emit(index))
            layout.addWidget(button)
            self.buttons.append(button)
            button.setVisible(i in (0, 2, 4))
        layout.addStretch()
        self.update_step(0)

    def update_step(self, current: int):
        for i, button in enumerate(self.buttons):
            button.setProperty(
                "state",
                "active"
                if i == ({1: 0, 3: 2}.get(current, current))
                else "complete"
                if i < current
                else "pending",
            )
            button.setEnabled(i <= current and i != 1)
            button.style().unpolish(button)
            button.style().polish(button)


class EmptyState(QFrame):
    files_dropped = Signal(list)

    def __init__(self, title: str, subtitle: str):
        super().__init__()
        self.setObjectName("dropZone")
        self.setAcceptDrops(True)
        self.setMinimumHeight(220)
        box = QVBoxLayout(self)
        box.setContentsMargins(32, 28, 32, 28)
        box.setAlignment(Qt.AlignmentFlag.AlignCenter)
        illustration = QLabel()
        self.illustration = illustration
        illustration.setPixmap(icon("upload", "#2563EB", 54).pixmap(54, 54))
        illustration.setAlignment(Qt.AlignmentFlag.AlignCenter)
        box.addWidget(illustration)
        self.title = QLabel(title)
        self.title.setObjectName("sectionTitle")
        self.title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        box.addWidget(self.title)
        self.subtitle = QLabel(subtitle)
        self.subtitle.setObjectName("muted")
        self.subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.subtitle.setWordWrap(True)
        box.addWidget(self.subtitle)
        self.actions = QHBoxLayout()
        self.actions.setAlignment(Qt.AlignmentFlag.AlignCenter)
        box.addLayout(self.actions)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            self.setProperty("dragOver", True)
            self.style().polish(self)
            event.acceptProposedAction()

    def dragLeaveEvent(self, event):
        self.setProperty("dragOver", False)
        self.style().polish(self)

    def dropEvent(self, event):
        from pathlib import Path

        self.setProperty("dragOver", False)
        self.style().polish(self)
        self.files_dropped.emit([Path(u.toLocalFile()) for u in event.mimeData().urls() if u.isLocalFile()])
        event.acceptProposedAction()


class DocumentCard(QFrame):
    def __init__(self, title: str, description: str = ""):
        super().__init__()
        self.setObjectName("card")
        self.box = QVBoxLayout(self)
        self.box.setContentsMargins(24, 20, 24, 20)
        self.box.setSpacing(12)
        self.heading = QLabel(title)
        self.heading.setObjectName("sectionTitle")
        self.box.addWidget(self.heading)
        if description:
            label = QLabel(description)
            self.description = label
            label.setObjectName("muted")
            label.setWordWrap(True)
            label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
            self.box.addWidget(label)


class LoadingState(DocumentCard):
    def __init__(self, title: str, description: str):
        super().__init__(title, description)
        self.description.setMinimumHeight(self.description.fontMetrics().height() * 2 + 4)
        self.stage = QLabel("Подготовка изображения")
        self.stage.setObjectName("loadingStage")
        self.stage.setWordWrap(True)
        self.box.addWidget(self.stage)
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(8)
        self.box.addWidget(self.progress)
        self.details = QLabel("Подготовка · Поиск документа · Текст · Данные · Проверка")
        self.details.setObjectName("caption")
        self.details.setWordWrap(True)
        self.box.addWidget(self.details)


class ConfirmDialog(QMessageBox):
    @staticmethod
    def ask(parent, title: str, text: str) -> bool:
        return (
            QMessageBox.question(
                parent,
                title,
                text,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            == QMessageBox.StandardButton.Yes
        )


class PhotoThumbnail(QFrame):
    selected = Signal(int)
    role_changed = Signal(int, str)
    menu_requested = Signal(int, object)

    def __init__(self, index: int, path, rotation: int = 0, role: str = "UNKNOWN"):
        super().__init__()
        from PySide6.QtGui import QImageReader

        self.setObjectName("photoThumbnail")
        self.setFixedWidth(226)
        box = QVBoxLayout(self)
        box.setContentsMargins(12, 12, 12, 12)
        box.setSpacing(8)
        image = QPushButton()
        image.setObjectName("thumbnailImage")
        image.setFixedHeight(118)
        reader = QImageReader(str(path))
        reader.setAutoTransform(True)
        original_size = reader.size()
        if original_size.isValid():
            reader.setScaledSize(original_size.scaled(QSize(380, 236), Qt.AspectRatioMode.KeepAspectRatio))
        qimage = reader.read()
        if not qimage.isNull():
            from PySide6.QtGui import QTransform

            pix = QPixmap.fromImage(qimage).transformed(QTransform().rotate(rotation))
            image.setIcon(QIcon(pix))
            image.setIconSize(QSize(196, 110))
        else:
            image.setText("Не удалось открыть фото")
        image.clicked.connect(lambda: self.selected.emit(index))
        box.addWidget(image)
        row = QHBoxLayout()
        row.addWidget(QLabel(f"Фото {index + 1}"), 1)
        menu = IconButton("more", "Повернуть, переместить или убрать фото")
        menu.clicked.connect(
            lambda: self.menu_requested.emit(index, menu.mapToGlobal(menu.rect().bottomLeft()))
        )
        row.addWidget(menu)
        box.addLayout(row)
        self.role = QComboBox()
        for label, value in [
            ("Сторона не указана", "UNKNOWN"),
            ("Лицевая сторона", "FRONT"),
            ("Оборотная сторона", "BACK"),
            ("Дополнительный документ", "ADDITIONAL"),
        ]:
            self.role.addItem(label, value)
        self.role.setCurrentIndex(max(0, self.role.findData(str(role))))
        self.role.currentIndexChanged.connect(lambda: self.role_changed.emit(index, self.role.currentData()))
        box.addWidget(self.role)


class ReviewField(QWidget):
    value_edited = Signal(str, str)
    focus_requested = Signal(str)
    details_requested = Signal(str)
    address_requested = Signal(str)

    def __init__(self, name: str):
        super().__init__()
        self.name = name
        self.setObjectName("reviewField")
        box = QVBoxLayout(self)
        box.setContentsMargins(0, 0, 0, 0)
        box.setSpacing(5)
        top = QHBoxLayout()
        self.label = QLabel(FIELD_LABELS[name])
        self.label.setObjectName("fieldLabel")
        top.addWidget(self.label, 1)
        self.status_label = StatusIndicator()
        top.addWidget(self.status_label)
        box.addLayout(top)
        self.editor = DateField() if name in DATE_FIELDS else TextField("Не указано")
        self.editor.setAccessibleName(FIELD_LABELS[name])
        self.editor.installEventFilter(self)
        self.editor.editingFinished.connect(lambda: self.value_edited.emit(name, self.editor.text()))
        if name in {"registration_address", "living_address"}:
            self.editor.setReadOnly(True)
            action = QAction(icon("next"), "Открыть редактор адреса", self.editor)
            action.triggered.connect(lambda: self.address_requested.emit(name))
            self.editor.addAction(action, QLineEdit.ActionPosition.TrailingPosition)
        box.addWidget(self.editor)
        bottom = QHBoxLayout()
        self.source = QLabel()
        self.source.setObjectName("caption")
        self.source.setWordWrap(True)
        bottom.addWidget(self.source, 1)
        self.representation = QComboBox()
        self.representation.setObjectName("representation")
        self.representation.setMaximumWidth(110)
        self.representation.addItems(["Кириллица", "Латиница"])
        self.representation.currentIndexChanged.connect(self.choose_representation)
        bottom.addWidget(self.representation)
        self.locate = IconButton(
            "source", "Показать поле на документе", lambda: self.focus_requested.emit(name)
        )
        self.locate.setFixedSize(24, 24)
        bottom.addWidget(self.locate)
        self.details = IconButton(
            "more", "Варианты OCR и копирование", lambda: self.details_requested.emit(name)
        )
        self.details.setFixedSize(24, 24)
        bottom.addWidget(self.details)
        box.addLayout(bottom)
        self.validation = QLabel()
        self.validation.setObjectName("fieldError")
        self.validation.setWordWrap(True)
        self.validation.setVisible(False)
        box.addWidget(self.validation)
        self.proposal = QLabel()
        self.proposal.setWordWrap(True)
        self.proposal.setObjectName("caption")
        self.proposal.hide()
        box.addWidget(self.proposal)
        self._evidence = FieldEvidence()

    def eventFilter(self, watched, event):
        if watched is self.editor and event.type() == QEvent.Type.FocusIn:
            self.focus_requested.emit(self.name)
        return super().eventFilter(watched, event)

    def choose_representation(self, index):
        value = getattr(self._evidence, "original_latin" if index else "original_cyrillic", None)
        if value:
            self.editor.setText(value)
            self.value_edited.emit(self.name, value)

    def refresh(self, value: str, evidence: FieldEvidence, required: bool):
        self._evidence = evidence
        self.editor.blockSignals(True)
        self.editor.setText(value)
        self.editor.setToolTip(value)
        self.editor.blockSignals(False)
        status = evidence.status
        if value and status == FieldStatus.MISSING:
            status = FieldStatus.REVIEW
        self.status_label.update_status(status, required)
        self.editor.setProperty("invalid", status == FieldStatus.INVALID or required and not value)
        self.editor.style().unpolish(self.editor)
        self.editor.style().polish(self.editor)
        side = ""
        from document_excel.domain.resolution import displayed_candidate, manual

        candidate = displayed_candidate(evidence, value)
        if candidate:
            side = {
                "FRONT": "лицевая сторона",
                "BACK": "оборотная сторона",
                "ADDITIONAL": "дополнительный документ",
            }.get(str(getattr(candidate, "source_side", "")), "фотография документа")
            if candidate.source_method == "MRZ":
                side = "MRZ · " + side
        source = (
            "Введено пользователем"
            if manual(evidence)
            else f"Источник: {side}"
            if side
            else "Нужно заполнить"
            if not value
            else "Сверьте с документом"
        )
        self.source.setText(source)
        self.locate.setVisible(bool(evidence.candidates or evidence.expected_bbox))
        self.details.setVisible(bool(value or evidence.candidates or evidence.expected_bbox))
        proposed = next(
            (
                c
                for c in evidence.candidates
                if c.validation_status != FieldStatus.INVALID and not c.validation_errors
            ),
            None,
        )
        invalid_proposed = evidence.candidates[0] if not proposed and evidence.candidates else None
        self.proposal.setText(
            "Кандидат: " + proposed.normalized_value + " · не подтверждён"
            if not value and proposed
            else "Распознано: " + invalid_proposed.raw_value + " · требуется исправление"
            if not value and invalid_proposed
            else ""
        )
        self.proposal.setVisible(bool(self.proposal.text()))
        bilingual = bool(
            getattr(evidence, "original_cyrillic", None) and getattr(evidence, "original_latin", None)
        )
        self.representation.blockSignals(True)
        self.representation.setVisible(bilingual)
        if bilingual:
            self.representation.setCurrentIndex(
                1 if value == getattr(evidence, "original_latin", None) else 0
            )
        self.representation.blockSignals(False)
        self.validation.setText(
            evidence.message
            if status
            in {
                FieldStatus.INVALID,
                FieldStatus.CONFLICT,
                FieldStatus.NOT_RECOGNIZED,
                FieldStatus.NOT_PRESENT,
            }
            or evidence.manual_conflict
            else ""
        )
        self.validation.setVisible(bool(self.validation.text()))


class CollapsibleSection(QFrame):
    def __init__(self, title: str, opened: bool = True):
        super().__init__()
        self.setObjectName("formSection")
        box = QVBoxLayout(self)
        box.setContentsMargins(0, 0, 0, 16)
        box.setSpacing(14)
        self.toggle = QPushButton(title)
        self.toggle.setObjectName("sectionToggle")
        self.toggle.setCheckable(True)
        self.toggle.setChecked(opened)
        self.toggle.setIcon(icon("next"))
        box.addWidget(self.toggle)
        self.body = QWidget()
        self.grid = QGridLayout(self.body)
        self.grid.setContentsMargins(0, 0, 0, 0)
        self.grid.setHorizontalSpacing(14)
        self.grid.setVerticalSpacing(16)
        self.body.setVisible(opened)
        self.toggle.toggled.connect(self.body.setVisible)
        box.addWidget(self.body)


class ReviewForm(QScrollArea):
    value_edited = Signal(str, str)
    focus_requested = Signal(str)
    details_requested = Signal(str)
    address_requested = Signal(str)

    def __init__(self):
        super().__init__()
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setObjectName("reviewScroll")
        body = QWidget()
        body.setObjectName("reviewBody")
        box = QVBoxLayout(body)
        box.setContentsMargins(24, 22, 24, 20)
        box.setSpacing(0)
        self.fields: dict[str, ReviewField] = {}
        self.preview_only = False
        self.sections = []
        groups = [
            (
                "Личные данные",
                ["full_name", "birth_date", "personal_identifier"],
                True,
            ),
            ("Фамилия, имя и отчество", ["surname", "first_name", "patronymic"], True),
            (
                "Документ",
                [
                    "passport_type",
                    "passport_number",
                    "passport_issued_by",
                    "passport_department_code",
                    "passport_issued_date",
                ],
                True,
            ),
            ("Прописка", ["registration_address"], True),
            ("Контакты", ["whatsapp_phone", "additional_phone"], False),
            (
                "Дополнительно",
                [
                    "gender",
                    "citizenship",
                    "expiration_date",
                    "living_address",
                    "place_of_birth",
                    "surname_cyrillic",
                    "surname_latin",
                    "first_name_cyrillic",
                    "first_name_latin",
                    "patronymic_cyrillic",
                    "object_address",
                    "notes",
                ],
                False,
            ),
        ]
        for title, names, opened in groups:
            section = CollapsibleSection(title, opened)
            self.sections.append(section)
            for row, name in enumerate(names):
                field = ReviewField(name)
                field.value_edited.connect(
                    lambda n, v: self.value_edited.emit(n, v) if not self.preview_only else None
                )
                field.focus_requested.connect(self.focus_requested)
                field.details_requested.connect(self.details_requested)
                field.address_requested.connect(
                    lambda n: self.address_requested.emit(n) if not self.preview_only else None
                )
                self.fields[name] = field
                section.grid.addWidget(field, row, 0)
            box.addWidget(section)
        box.addStretch()
        self.setWidget(body)

    def refresh(self, case, model, required_fields, *, preview_only=False):
        self.preview_only = preview_only
        for name, widget in self.fields.items():
            if widget.editor.property("baseReadOnly") is None:
                widget.editor.setProperty("baseReadOnly", widget.editor.isReadOnly())
            widget.editor.setReadOnly(preview_only or bool(widget.editor.property("baseReadOnly")))
            widget.details.setEnabled(not preview_only)
            widget.locate.setEnabled(not preview_only)
            widget.representation.setEnabled(not preview_only)
            for button in widget.findChildren(QPushButton):
                button.setEnabled(not preview_only)
            row = model.fields.index(name)
            value = model.data(model.index(row, 1)) or data_value(case.data, name)
            widget.refresh(value, case.evidence.get(name, FieldEvidence()), name in required_fields)

    def commit(self):
        if getattr(self, "preview_only", False):
            return
        editor = self.focusWidget()
        for name, field in self.fields.items():
            if editor is field.editor:
                self.value_edited.emit(name, field.editor.text())

    def reveal(self, name):
        field = self.fields[name]
        for section in self.sections:
            if section.isAncestorOf(field):
                section.toggle.setChecked(True)
        self.ensureWidgetVisible(field, 0, 30)
        field.editor.setFocus()
