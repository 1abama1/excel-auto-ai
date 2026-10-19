"""Value-free outcome codes shared by local diagnostics and review messages."""

from enum import StrEnum


class OcrProblem(StrEnum):
    OCR_ENGINE_ERROR = "OCR_ENGINE_ERROR"
    NO_TEXT_DETECTED = "NO_TEXT_DETECTED"
    TEXT_DETECTED_BUT_FIELD_NOT_FOUND = "TEXT_DETECTED_BUT_FIELD_NOT_FOUND"
    FIELD_REJECTED_BY_VALIDATION = "FIELD_REJECTED_BY_VALIDATION"
    FIELD_CONFLICT = "FIELD_CONFLICT"
    STALE_RESULT = "STALE_RESULT"
    DOCUMENT_IMAGE_LOW_QUALITY = "DOCUMENT_IMAGE_LOW_QUALITY"
    FIELD_NOT_PRESENT = "FIELD_NOT_PRESENT"


MESSAGES = {
    OcrProblem.OCR_ENGINE_ERROR: "Техническая ошибка OCR. Проверьте локальные модели и диагностику.",
    OcrProblem.NO_TEXT_DETECTED: "Текст на изображении не обнаружен. Проверьте фотографию или выделите карточку.",
    OcrProblem.TEXT_DETECTED_BUT_FIELD_NOT_FOUND: "Текст документа распознан, но это поле определить не удалось. Повторите чтение его области или исправьте значение.",
    OcrProblem.FIELD_REJECTED_BY_VALIDATION: "Значение прочитано, но не прошло проверку формата. Сверьте предложенный вариант с фотографией.",
    OcrProblem.FIELD_CONFLICT: "Обнаружены разные варианты значения. Выберите правильный по фотографии.",
    OcrProblem.STALE_RESULT: "Фотографии или параметры изменились. Старый результат не применён; повторите распознавание.",
    OcrProblem.DOCUMENT_IMAGE_LOW_QUALITY: "Мелкий или нечёткий текст. Найденные данные сохранены; проверьте их или загрузите более чёткий исходник.",
    OcrProblem.FIELD_NOT_PRESENT: "Поле не предусмотрено на этой стороне документа.",
}


def safe_resolution_summary(case, image_id):
    from document_excel.domain.models import FieldStatus

    found, rejected, conflicts = set(), set(), set()
    reasons = {}
    for name, evidence in case.evidence.items():
        related = [c for c in evidence.candidates if c.source_image_id == image_id]
        if not related:
            continue
        valid = [c for c in related if c.validation_status != FieldStatus.INVALID and not c.validation_errors]
        if evidence.status == FieldStatus.CONFLICT:
            conflicts.add(name)
            reasons[name] = OcrProblem.FIELD_CONFLICT.value
        elif not valid:
            rejected.add(name)
            reasons[name] = OcrProblem.FIELD_REJECTED_BY_VALIDATION.value
        elif getattr(case.data, name, None):
            found.add(name)
    return {
        "accepted_fields": len(found),
        "rejected_fields": len(rejected),
        "conflict_fields": len(conflicts),
        "field_problems": reasons,
    }
