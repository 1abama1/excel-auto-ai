"""Camera and photo review UI. Captures stay full size; only live preview scales."""

from __future__ import annotations

import numpy as np
from PySide6.QtCore import QElapsedTimer, QRectF, Qt, QTimer, Signal
from PySide6.QtGui import QColor, QImage, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from document_excel.domain.errors import AppError
from document_excel.domain.models import DocumentSide
from document_excel.infrastructure.camera import CameraService
from document_excel.infrastructure.document_images import DocumentImageProcessor
from document_excel.infrastructure.preprocessing import detect_document_region
from document_excel.presentation.widgets import ImageViewer


class CameraPreview(QLabel):
    def __init__(self):
        super().__init__("Подключение камеры…")
        self.setMinimumSize(320, 220)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.picture = QPixmap()

    def frame(self, array):
        h, w = array.shape[:2]
        self.picture = QPixmap.fromImage(
            QImage(array.data, w, h, array.strides[0], QImage.Format.Format_RGB888).copy()
        )
        self.update()

    def paintEvent(self, event):
        if self.picture.isNull():
            super().paintEvent(event)
            return
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor("#101827"))
        size = self.picture.size().scaled(self.size(), Qt.AspectRatioMode.KeepAspectRatio)
        x, y = (self.width() - size.width()) // 2, (self.height() - size.height()) // 2
        painter.drawPixmap(
            x,
            y,
            self.picture.scaled(
                size, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation
            ),
        )
        width = min(size.width() * 0.85, size.height() * 0.8 * 1.586)
        frame = QRectF(
            self.width() / 2 - width / 2, self.height() / 2 - width / 1.586 / 2, width, width / 1.586
        )
        painter.setPen(QPen(QColor("#60a5fa"), 3, Qt.PenStyle.DashLine))
        painter.drawRoundedRect(frame, 10, 10)


class PhotoEditor(QWidget):
    def __init__(self):
        super().__init__()
        self.original = None
        self.current = None
        self.processor = DocumentImageProcessor()
        box = QVBoxLayout(self)
        box.setContentsMargins(0, 0, 0, 0)
        actions = QHBoxLayout()
        for text, method in [
            ("Оригинал", self.reset),
            ("Автообрезка", self.auto_crop),
            ("Обрезать вручную", self.manual_crop),
            ("Повернуть", self.rotate),
        ]:
            button = QPushButton(text)
            button.clicked.connect(method)
            actions.addWidget(button)
        box.addLayout(actions)
        self.viewer = ImageViewer()
        self.viewer.setMinimumSize(300, 220)
        self.viewer.region_selected.connect(self.crop)
        box.addWidget(self.viewer, 1)
        self.status = QLabel()
        self.status.setWordWrap(True)
        box.addWidget(self.status)

    def load(self, array):
        self.original = array.copy()
        self.current = self.processor.suggest(self.original)
        self.refresh()

    def refresh(self):
        self.current = np.ascontiguousarray(self.current)
        self.viewer.set_image(self.current)
        h, w = self.current.shape[:2]
        notes = self.processor.quality(self.current)
        self.status.setText(
            f"{w} × {h} · Проверьте, что документ виден целиком.\n"
            + (" ".join(notes) or "Сверьте читаемость мелкого текста перед подтверждением.")
        )

    def reset(self):
        if self.original is not None:
            self.current = self.original.copy()
            self.refresh()

    def auto_crop(self):
        if self.current is not None:
            self.current = self.processor.suggest(self.current)
            self.refresh()

    def manual_crop(self):
        self.status.setText("Выделите документ прямоугольником. Исходная фотография сохраняется отдельно.")
        self.viewer.select_region()

    def crop(self, area):
        x, y, width, height = area
        if self.current is not None and width >= 100 and height >= 100:
            h, w = self.current.shape[:2]
            left, top, right, bottom = max(0, x), max(0, y), min(w, x + width), min(h, y + height)
            if right - left >= 100 and bottom - top >= 100:
                self.current = self.current[top:bottom, left:right].copy()
                self.refresh()

    def rotate(self):
        if self.current is not None:
            self.current = np.rot90(self.current, -1).copy()
            self.refresh()


class PhotoReviewDialog(QDialog):
    def __init__(self, array, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Проверить фотографию документа")
        self.resize(980, 720)
        box = QVBoxLayout(self)
        self.editor = PhotoEditor()
        box.addWidget(self.editor, 1)
        self.editor.load(array)
        row = QHBoxLayout()
        cancel = QPushButton("Отмена")
        cancel.clicked.connect(self.reject)
        use = QPushButton("Использовать фото")
        use.setObjectName("primary")
        use.clicked.connect(self.accept)
        row.addWidget(cancel)
        row.addStretch()
        row.addWidget(use)
        box.addLayout(row)


class CameraDialog(QDialog):
    photo_accepted = Signal(object, object, object)

    def __init__(self, parent=None, *, side=None, service=None, physical=False):
        super().__init__(parent)
        self.setWindowTitle("Сфотографировать документ")
        self.setProperty("documentCameraDialog", True)
        self.resize(1040, 760)
        self.setMinimumSize(650, 520)
        self.service = service or CameraService(self)
        self.only_side = side
        self.physical = physical
        self._quality_clock = QElapsedTimer()
        self._quality_clock.start()
        self._last_quality = -1000
        self._closing = False
        self._finish_code = QDialog.DialogCode.Rejected
        self.accepted_sides = set()
        self.raw = None
        self.commit_photo = None
        self.live = False
        self.job = 0
        box = QVBoxLayout(self)
        row = QHBoxLayout()
        self.devices = QComboBox()
        self.devices.setMinimumWidth(180)
        self.devices.currentIndexChanged.connect(self.device_changed)
        row.addWidget(self.devices, 1)
        self.switch = QPushButton("Сменить камеру")
        self.switch.clicked.connect(
            lambda: self.devices.setCurrentIndex(
                (self.devices.currentIndex() + 1) % max(1, self.devices.count())
            )
        )
        row.addWidget(self.switch)
        self.refresh_button = QPushButton("Обновить камеры")
        self.refresh_button.clicked.connect(self.discover)
        row.addWidget(self.refresh_button)
        self.side = QComboBox()
        self.side.addItem("Лицевая сторона", DocumentSide.FRONT)
        self.side.addItem("Оборотная сторона", DocumentSide.BACK)
        self.side.setCurrentIndex(1 if side == DocumentSide.BACK else 0)
        self.side.currentIndexChanged.connect(self.side_changed)
        row.addWidget(self.side)
        box.addLayout(row)
        self.hint = QLabel()
        self.hint.setWordWrap(True)
        box.addWidget(self.hint)
        self.preview_quality = QLabel()
        self.preview_quality.setWordWrap(True)
        self.preview_quality.setObjectName("caption")
        self.preview_quality.setVisible(physical)
        box.addWidget(self.preview_quality)
        self.stack = QStackedWidget()
        self.preview = CameraPreview()
        self.editor = PhotoEditor()
        self.stack.addWidget(self.preview)
        self.stack.addWidget(self.editor)
        box.addWidget(self.stack, 1)
        self.status = QLabel("Поиск доступных камер…")
        self.status.setWordWrap(True)
        box.addWidget(self.status)
        controls = QHBoxLayout()
        self.cancel_button = QPushButton("Отмена")
        self.cancel_button.clicked.connect(self.reject)
        controls.addWidget(self.cancel_button)
        controls.addStretch()
        self.shutter = QPushButton("Сфотографировать")
        self.shutter.setObjectName("primary")
        self.shutter.setEnabled(False)
        self.shutter.clicked.connect(self.capture)
        controls.addWidget(self.shutter)
        self.retake = QPushButton("Переснять")
        self.retake.clicked.connect(self.resume)
        controls.addWidget(self.retake)
        self.use = QPushButton("Использовать фото")
        self.use.setObjectName("primary")
        self.use.clicked.connect(self.use_photo)
        controls.addWidget(self.use)
        box.addLayout(controls)
        self.service.devices_found.connect(self.found)
        self.service.preview.connect(self.frame)
        self.service.photograph.connect(self.captured)
        self.service.ready.connect(self.ready)
        self.service.error.connect(self.failed)
        self.side_changed()
        self.resume()
        QTimer.singleShot(0, self.discover)

    def side_changed(self):
        if self.physical:
            self.hint.setText(
                "Положите лицевую сторону карты перед камерой. Карта должна помещаться целиком."
                if self.side.currentData() == DocumentSide.FRONT
                else "Переверните карту и сфотографируйте оборотную сторону. Проверьте читаемость трёх строк MRZ и отсутствие бликов."
            )
            return
        self.hint.setText(
            "Поместите лицевую сторону документа в рамку"
            if self.side.currentData() == DocumentSide.FRONT
            else "Теперь сфотографируйте оборотную сторону"
        )

    def after_stop(self, callback):
        self.job += 1
        token = self.job
        self.service.stop()
        self.live = False
        self.shutter.setEnabled(False)

        def wait():
            if token != self.job or self._closing:
                return
            if self.service.busy:
                QTimer.singleShot(50, self, wait)
            else:
                callback()

        QTimer.singleShot(50, self, wait)

    def discover(self):
        self.status.setText("Поиск камер. Если их нет, проверьте подключение и разрешения Windows.")
        self.after_stop(self.service.discover)

    def found(self, devices):
        if self._closing:
            return
        self.devices.blockSignals(True)
        self.devices.clear()
        for device in devices:
            self.devices.addItem(device.name, device)
        self.devices.blockSignals(False)
        if devices:
            self.device_changed()
        else:
            self.failed(
                "Камеры не найдены или доступ запрещён. Подключите камеру, разрешите доступ для классических приложений Windows либо загрузите фотографии с диска."
            )

    def device_changed(self):
        device = self.devices.currentData()
        if device is not None:
            self.status.setText("Подключение камеры и проверка разрешений…")
            self.resume()
            self.after_stop(lambda: self.service.start(device))

    def ready(self, resolutions):
        if self._closing:
            return
        self.live = True
        self.shutter.setEnabled(self.stack.currentIndex() == 0)
        self.status.setText(
            "Проверенные разрешения: "
            + ", ".join(f"{w}×{h}" for w, h in resolutions)
            + ". Снимок использует полное разрешение потока."
        )

    def frame(self, array):
        if not self._closing and self.stack.currentIndex() == 0:
            self.preview.frame(array)
            now = self._quality_clock.elapsed()
            if self.physical and now - self._last_quality >= 1000:
                self._last_quality = now
                notes = [
                    note
                    for note in self.editor.processor.quality(array)
                    if not note.startswith("Мелкий текст")
                ]
                region = detect_document_region(array)
                if not region.detected:
                    notes.append("Поместите карту целиком в рамку; границы пока не определены уверенно.")
                else:
                    x, y, w, h = region.bbox
                    height, width = array.shape[:2]
                    if min(x, y, width - x - w, height - y - h) < min(width, height) * 0.02:
                        notes.append("Карта близко к краю кадра. Отодвиньте её, чтобы не обрезать текст.")
                self.preview_quality.setText(
                    " ".join(notes) or "Сверьте читаемость текста и отсутствие отражений."
                )

    def capture(self):
        self.shutter.setEnabled(False)
        self.service.capture()

    def captured(self, array):
        if self._closing:
            return
        self.raw = array.copy()
        self.preview_quality.hide()
        self.editor.load(array)
        self.stack.setCurrentIndex(1)
        self.shutter.hide()
        self.retake.show()
        self.use.show()
        self.devices.setEnabled(False)
        self.switch.setEnabled(False)

    def resume(self):
        self.preview_quality.setVisible(self.physical)
        self.stack.setCurrentIndex(0)
        self.raw = None
        self.shutter.show()
        self.shutter.setEnabled(self.live)
        self.retake.hide()
        self.use.hide()
        self.devices.setEnabled(True)
        self.switch.setEnabled(True)

    def use_photo(self):
        if self.raw is None or self.editor.current is None:
            return
        side = DocumentSide(self.side.currentData())
        if self.commit_photo is not None:
            try:
                self.commit_photo(side, self.raw.copy(), self.editor.current.copy())
            except (AppError, OSError) as exc:
                self.status.setText(
                    str(exc)
                    if isinstance(exc, AppError)
                    else "Не удалось сохранить снимок. Проверьте место на диске и папку фотографий."
                )
                return
        self.photo_accepted.emit(side, self.raw.copy(), self.editor.current.copy())
        self.accepted_sides.add(side)
        if self.only_side is not None or len(self.accepted_sides) == 2:
            self.accept()
        else:
            self.side.setCurrentIndex(1 if side == DocumentSide.FRONT else 0)
            self.resume()

    def failed(self, message):
        if not self._closing:
            self.status.setText(message)
            self.live = False
            self.shutter.setEnabled(False)

    def finish(self, code):
        if self._closing:
            return
        self._closing = True
        self._finish_code = code
        self.job += 1
        self.service.close()
        self.status.setText("Освобождение камеры…")
        self.shutter.setEnabled(False)
        self.use.setEnabled(False)
        self.cancel_button.setEnabled(False)
        self.poll_close()

    def poll_close(self):
        if self.service.busy:
            QTimer.singleShot(50, self, self.poll_close)
        else:
            super().done(self._finish_code)

    def accept(self):
        self.finish(QDialog.DialogCode.Accepted)

    def reject(self):
        self.finish(QDialog.DialogCode.Rejected)

    def closeEvent(self, event):
        event.ignore()
        self.reject()
