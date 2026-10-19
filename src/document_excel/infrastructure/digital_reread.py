"""Explicit single-field reading with semantic context supplied by the user."""

import cv2

from document_excel.domain.errors import AppError
from document_excel.infrastructure.digital_fields import DigitalBand, DigitalDocumentExtractor, DigitalLayout
from document_excel.infrastructure.digital_recovery import recover_digital_fields
from document_excel.infrastructure.document_profiles import KyrgyzDigitalIDProfile


def reread_digital_field(pixels, source, field, bbox, engine, cancel):
    h, w = pixels.shape[:2]
    x, y, bw, bh = bbox
    left, top = max(0, x), max(0, y)
    right, bottom = min(w, x + bw), min(h, y + bh)
    if right <= left or bottom <= top:
        raise AppError("Область поля вне изображения. Повторите OCR документа.")
    crop = pixels[top:bottom, left:right]
    scale = max(1.0, min(2.0, 1200 / max(crop.shape[:2])))
    if scale > 1.05:
        crop = cv2.resize(crop, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    result = engine.recognize(crop, source.id, "digital-selected-field", cancel)
    if not result.blocks:
        return []
    profile = KyrgyzDigitalIDProfile()
    layout = DigitalLayout(result, profile)
    # A value-only selection has no printed caption. Use the selected field
    # as context, without adding a fabricated OCR text block.
    layout.bands = {field: DigitalBand((0, 0, result.width, result.height), result.blocks)}
    found = [c for cls in DigitalDocumentExtractor.extractors for c in cls().extract(layout)]
    found, _, _ = recover_digital_fields(crop, result, profile, found, engine, cancel, layout=layout)
    allowed = {field, "passport_type"} if field == "passport_number" else {field}
    for c in found:
        if c.source_bbox:
            bx, by, cw, ch = c.source_bbox
            c.source_bbox = (
                left + round(bx / scale),
                top + round(by / scale),
                round(cw / scale),
                round(ch / scale),
            )
        c.source_side = source.role
        c.extraction_method = "digital-selected-field/" + c.extraction_method
    return [c for c in found if c.field_name in allowed]
