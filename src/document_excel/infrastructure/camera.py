"""OpenCV camera I/O in owned worker threads; Qt only receives preview copies."""

from __future__ import annotations

import os
import threading
import time
from dataclasses import dataclass
from pathlib import Path

import cv2
from PySide6.QtCore import QObject, Signal


@dataclass(frozen=True)
class CameraDevice:
    index: int
    name: str


class CameraService(QObject):
    devices_found = Signal(object)
    preview = Signal(object)
    photograph = Signal(object)
    ready = Signal(object)
    error = Signal(str)
    stopped = Signal()

    def __init__(self, parent=None, capture_factory=None, indices=None):
        super().__init__(parent)
        self.factory = capture_factory or cv2.VideoCapture
        self.indices = indices
        self._stop = threading.Event()
        self._discovery_stop = threading.Event()
        self._shutter = threading.Event()
        self.worker_thread: threading.Thread | None = None
        self.discovery: threading.Thread | None = None

    @property
    def running(self):
        return self.worker_thread is not None and self.worker_thread.is_alive()

    @property
    def busy(self):
        return self.running or self.discovery is not None and self.discovery.is_alive()

    def _open(self, index):
        backends = [cv2.CAP_DSHOW, cv2.CAP_MSMF] if os.name == "nt" else [cv2.CAP_V4L2, cv2.CAP_ANY]
        for backend in backends:
            if self._discovery_stop.is_set():
                return None
            cap = None
            try:
                cap = self.factory(index, backend)
                if cap.isOpened():
                    return cap
            except (cv2.error, OSError):
                pass
            if cap is not None:
                cap.release()
        return None

    def discover(self):
        if self.busy:
            return
        self._discovery_stop.clear()

        def work():
            indices = self.indices
            if indices is None:
                indices = (
                    range(8)
                    if os.name == "nt"
                    else sorted(
                        int(p.name[5:]) for p in Path("/dev").glob("video[0-9]*") if p.name[5:].isdigit()
                    )
                )
            devices = []
            try:
                for index in indices:
                    if self._discovery_stop.is_set():
                        return
                    cap = self._open(index)
                    if cap is not None:
                        cap.release()
                        devices.append(CameraDevice(index, f"Камера {index + 1}"))
                if not self._discovery_stop.is_set():
                    self.devices_found.emit(devices)
            except (cv2.error, OSError):
                self.error.emit("Не удалось найти камеры. Проверьте подключение и разрешения Windows.")

        self.discovery = threading.Thread(target=work, daemon=True, name="camera-discovery")
        self.discovery.start()

    @staticmethod
    def negotiate(cap, stop):
        """Verify delivered full-size frames rather than trusting set() results."""
        supported = set()
        best = None
        cap.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter.fourcc(*"MJPG"))
        cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
        # Driver-negotiated candidates, including high-resolution photo formats.
        for width, height in [
            (4096, 3072),
            (3840, 2160),
            (2592, 1944),
            (2560, 1440),
            (1920, 1080),
            (1280, 720),
            (640, 480),
        ]:
            if stop.is_set():
                return None, []
            cap.set(cv2.CAP_PROP_FRAME_WIDTH, width)
            cap.set(cv2.CAP_PROP_FRAME_HEIGHT, height)
            ok, frame = cap.read()
            if ok and frame is not None and frame.size:
                actual = (frame.shape[1], frame.shape[0])
                supported.add(actual)
                if best is None or actual[0] * actual[1] > best[0] * best[1]:
                    best = actual
        if best is None:
            return None, []
        cap.set(cv2.CAP_PROP_FRAME_WIDTH, best[0])
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, best[1])
        # Flush stale buffers after changing format; do not enable the shutter
        # until the driver has delivered a frame at the selected resolution.
        delivered: tuple[int, int] | None = None
        for _ in range(8):
            if stop.is_set():
                return None, []
            ok, frame = cap.read()
            if ok and frame is not None and frame.size:
                delivered = (frame.shape[1], frame.shape[0])
                if delivered == best:
                    break
        if delivered is None:
            return None, []
        best = delivered
        return best, sorted(supported, key=lambda size: size[0] * size[1], reverse=True)

    def start(self, device: CameraDevice):
        if self.running:
            return
        self._stop.clear()
        self._discovery_stop.clear()
        self._shutter.clear()

        def work():
            cap = None
            try:
                cap = self._open(device.index)
                if cap is None:
                    self.error.emit(
                        "Камера недоступна. Разрешите доступ к камере для классических приложений Windows, закройте другие программы или выберите другую камеру."
                    )
                    return
                resolution, supported = self.negotiate(cap, self._stop)
                if self._stop.is_set():
                    return
                if resolution is None:
                    self.error.emit(
                        "Камера не передаёт кадры. Проверьте разрешение доступа и USB-подключение."
                    )
                    return
                self.ready.emit(supported)
                next_preview = 0.0
                failures = 0
                mismatches = 0
                while not self._stop.is_set():
                    ok, frame = cap.read()
                    if not ok or frame is None or not frame.size:
                        failures += 1
                        if failures >= 8:
                            self.error.emit(
                                "Поток камеры прерван. Камера отключена или занята. Подключите её снова и выберите «Обновить камеры»."
                            )
                            return
                        self._stop.wait(0.05)
                        continue
                    failures = 0
                    if self._shutter.is_set():
                        if (frame.shape[1], frame.shape[0]) != resolution:
                            mismatches += 1
                            if mismatches >= 10:
                                self.error.emit(
                                    "Камера изменила разрешение снимка. Переподключите её или выберите другую камеру."
                                )
                                return
                            continue
                        self._shutter.clear()
                        self.photograph.emit(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB).copy())
                    now = time.monotonic()
                    if now >= next_preview:
                        scale = min(1, 1000 / max(frame.shape[:2]))
                        small = cv2.resize(
                            frame, (round(frame.shape[1] * scale), round(frame.shape[0] * scale))
                        )
                        self.preview.emit(cv2.cvtColor(small, cv2.COLOR_BGR2RGB).copy())
                        next_preview = now + 0.1
                    self._stop.wait(0.005)
            except (cv2.error, OSError):
                if not self._stop.is_set():
                    self.error.emit("Ошибка камеры. Проверьте подключение, драйвер и разрешение доступа.")
            finally:
                if cap is not None:
                    cap.release()
                self.stopped.emit()

        self.worker_thread = threading.Thread(target=work, daemon=True, name="document-camera")
        self.worker_thread.start()

    def capture(self):
        if self.running:
            self._shutter.set()

    def stop(self):
        self._stop.set()
        self._shutter.clear()

    def close(self):
        self.stop()
        self._discovery_stop.set()
