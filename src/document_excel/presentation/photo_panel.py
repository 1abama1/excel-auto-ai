"""Compact per-client photo preview for the Excel step."""

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QImage, QPixmap
from PySide6.QtWidgets import QCheckBox, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget

from document_excel.domain.errors import AppError
from document_excel.domain.models import DocumentSide
from document_excel.infrastructure.document_images import selected_image


class DocumentPhotoPanel(QWidget):
    review_requested = Signal(str)
    exception_changed = Signal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.cache = {}
        box = QVBoxLayout(self)
        box.setContentsMargins(0, 0, 0, 0)
        title = QLabel("Фотографии в договоре")
        title.setObjectName("sectionTitle")
        box.addWidget(title)
        row = QHBoxLayout()
        self.cards = {}
        for side, title, area in [
            (DocumentSide.FRONT, "Лицевая сторона", "A1:E24"),
            (DocumentSide.BACK, "Оборотная сторона", "F1:K24"),
        ]:
            card = QWidget()
            layout = QHBoxLayout(card)
            layout.setContentsMargins(0, 4, 8, 4)
            image = QLabel("Нет фото")
            image.setAlignment(Qt.AlignmentFlag.AlignCenter)
            image.setFixedSize(120, 82)
            layout.addWidget(image)
            info = QVBoxLayout()
            heading = QLabel(f"{title} · {area}")
            heading.setWordWrap(True)
            info.addWidget(heading)
            status = QLabel()
            status.setWordWrap(True)
            info.addWidget(status)
            review = QPushButton("Проверить фото")
            review.clicked.connect(lambda _=False, s=side: self.review_requested.emit(s.value))
            info.addWidget(review)
            layout.addLayout(info, 1)
            row.addWidget(card, 1)
            self.cards[side] = (image, status, review)
        box.addLayout(row)
        self.status = QLabel()
        self.status.setWordWrap(True)
        box.addWidget(self.status)
        self.exception = QCheckBox("Разрешаю сохранить без отсутствующих фотографий для этого клиента")
        self.exception.toggled.connect(self.exception_changed)
        box.addWidget(self.exception)

    def refresh(self, case, profile):
        complete = True
        for side, (image, status, review) in self.cards.items():
            sources = [s for s in case.images if s.role == side]
            source = sources[0] if len(sources) == 1 else None
            review.setEnabled(source is not None)
            if source is None:
                complete = False
                image.clear()
                image.setText("Нет фото" if not sources else "Выберите одно")
                status.setText(
                    "Добавьте фото или назначьте сторону на шаге загрузки."
                    if not sources
                    else "Несколько фото одной стороны: уточните выбор."
                )
                continue
            try:
                path = source.selected_path
                key = (str(path), path.stat().st_mtime_ns, source.selected_rotation)
                if key not in self.cache:
                    array = selected_image(source)
                    h, w = array.shape[:2]
                    pixmap = QPixmap.fromImage(
                        QImage(array.data, w, h, array.strides[0], QImage.Format.Format_RGB888).copy()
                    ).scaled(
                        120,
                        82,
                        Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation,
                    )
                    if len(self.cache) >= 8:
                        self.cache.clear()
                    self.cache[key] = pixmap
                image.setPixmap(self.cache[key])
                status.setText(
                    "Подтверждённый кадр"
                    if source.processed_confirmed
                    else "Оригинал · можно обрезать и проверить"
                )
            except (AppError, OSError):
                complete = False
                image.clear()
                image.setText("Недоступно")
                status.setText("Добавьте фотографию заново.")
        self.status.setText(
            f"Лист: {profile.photo_sheet}. Снимки будут встроены с сохранением пропорций."
            if profile.photo_sheet
            else "Выберите лист фотографий в «Расположение полей». Для этого профиля фотографии пока не вставляются."
        )
        if profile.photos_required and not complete:
            self.status.setText(self.status.text() + " Нужны обе стороны документа.")
        self.exception.blockSignals(True)
        self.exception.setChecked(case.photo_exception)
        self.exception.blockSignals(False)
        self.exception.setVisible(
            bool(profile.photo_sheet and profile.allow_photo_exception and not complete)
        )
