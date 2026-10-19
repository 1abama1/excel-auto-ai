from __future__ import annotations

import logging
import threading
from collections.abc import Callable
from typing import Any

from PySide6.QtCore import QObject, QRunnable, Signal, Slot

from document_excel.domain.errors import AppError, Cancelled


class WorkerSignals(QObject):
    progress = Signal(int, str)
    result = Signal(object)
    error = Signal(str)
    finished = Signal()


class Worker(QRunnable):
    def __init__(self, function: Callable[..., Any], cancellable: bool = True) -> None:
        super().__init__()
        self.function = function
        self.cancellable = cancellable
        self.cancel_event = threading.Event()
        self.signals = WorkerSignals()

    def cancel(self) -> None:
        if self.cancellable:
            self.cancel_event.set()

    @Slot()
    def run(self) -> None:
        try:
            result = self.function(self.cancel_event, self.signals.progress.emit)
            if self.cancel_event.is_set():
                raise Cancelled()
            self.signals.result.emit(result)
        except AppError as exc:
            self.signals.error.emit(str(exc))
        except Exception as exc:
            # Never include exception repr, paths, OCR text or a traceback in persistent logs.
            logging.getLogger("document_excel").error("worker_failed type=%s", type(exc).__name__)
            self.signals.error.emit(
                "Не удалось выполнить операцию. Откройте диагностику; исходные файлы сохранены."
            )
        finally:
            self.signals.finished.emit()
