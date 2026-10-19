"""Render actual camera widgets using synthetic pixels; never opens hardware."""

import argparse
from pathlib import Path

import cv2
import numpy as np
from PySide6.QtCore import QObject, QTimer, Signal
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication

from document_excel.infrastructure.camera import CameraDevice
from document_excel.infrastructure.storage import Settings
from document_excel.presentation.camera_dialog import CameraDialog
from document_excel.presentation.theme import apply_theme


class SyntheticCamera(QObject):
    devices_found = Signal(object)
    preview = Signal(object)
    photograph = Signal(object)
    ready = Signal(object)
    error = Signal(str)
    busy = False

    def __init__(self):
        super().__init__()
        self.image = np.full((1080, 1920, 3), (55, 65, 75), np.uint8)
        cv2.rectangle(self.image, (280, 150), (1640, 950), (226, 237, 247), -1)
        for index, text in enumerate(
            ("SYNTHETIC DOCUMENT", "NOT A REAL ID", "FRONT / BACK", "PHOTO WORKFLOW CHECK")
        ):
            cv2.putText(
                self.image,
                text,
                (370, 310 + index * 155),
                cv2.FONT_HERSHEY_SIMPLEX,
                1.7,
                (25, 45, 65),
                3,
            )

    def discover(self):
        self.devices_found.emit([CameraDevice(0, "Synthetic camera")])

    def start(self, *_):
        self.ready.emit([(1920, 1080)])
        self.preview.emit(self.image)

    def capture(self):
        self.photograph.emit(self.image.copy())

    def stop(self):
        pass

    def close(self):
        pass


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--physical", action="store_true")
    args = parser.parse_args()
    app = QApplication([])
    apply_theme(app, Settings(theme="dark"))
    output = Path(
        "test-artifacts/physical-camera-visual" if args.physical else "test-artifacts/camera-visual"
    )
    output.mkdir(parents=True, exist_ok=True)
    for width, height in ((1000, 700), (800, 620)):
        dialog = CameraDialog(service=SyntheticCamera(), physical=args.physical)
        dialog.resize(width, height)
        dialog.show()
        QTest.qWait(350)
        dialog.grab().save(str(output / f"live-{width}.png"))
        dialog.capture()
        QTest.qWait(100)
        dialog.grab().save(str(output / f"review-{width}.png"))
        for button in (dialog.cancel_button, dialog.retake, dialog.use):
            assert dialog.rect().contains(button.geometry())
        dialog.reject()
        QTest.qWait(100)
        dialog.deleteLater()
        QTest.qWait(100)
    QTimer.singleShot(100, app.quit)
    app.exec()
    print("Synthetic camera UI: four screenshots; controls inside the dialog")


if __name__ == "__main__":
    main()
