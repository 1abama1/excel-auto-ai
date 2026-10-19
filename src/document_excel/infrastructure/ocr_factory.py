"""Selection is explicit. An unavailable engine never silently falls back."""

from document_excel.infrastructure.ocr_contract import OcrEngine
from document_excel.infrastructure.storage import Settings


def create_engine(settings: Settings) -> OcrEngine:
    if settings.ocr_engine == "rapidocr":
        from document_excel.infrastructure.rapid_ocr import RapidOCREngine

        return RapidOCREngine(settings)
    if settings.ocr_engine == "paddle":
        from document_excel.infrastructure.paddle_ocr import PaddleOCREngine

        return PaddleOCREngine(settings)
    if settings.ocr_engine == "tesseract":
        from document_excel.infrastructure.ocr import TesseractEngine

        return TesseractEngine(settings)
    raise ValueError("Неизвестный OCR-движок")
