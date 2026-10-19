class AppError(Exception):
    """An error whose message is safe to show in the user interface."""


class OcrFailure(AppError):
    def __init__(self, message, code, metrics):
        super().__init__(message)
        self.code, self.metrics = code, metrics


class Cancelled(AppError):
    def __init__(self) -> None:
        super().__init__("Обработка отменена. Исходные файлы не изменены.")
