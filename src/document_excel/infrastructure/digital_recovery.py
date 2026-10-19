"""Bounded optical retries, separate from the block-only digital extractors."""

import copy
import re

import cv2
import numpy as np

from document_excel.domain.errors import AppError, Cancelled
from document_excel.domain.models import FieldStatus, validate_field
from document_excel.domain.resolution import RegistrationAddressResolver, name_script
from document_excel.infrastructure.digital_fields import (
    _LATIN_HOMOGLYPHS,
    DigitalBand,
    DigitalDocumentExtractor,
    DigitalLayout,
    bounds,
    candidate,
)
from document_excel.infrastructure.ocr_contract import OcrResult, TextBlock


def bilingual_split(pixels):
    """Locate an isolated diagonal slash, preserving the original RGB pixels."""
    gray = cv2.cvtColor(pixels, cv2.COLOR_RGB2GRAY) if pixels.ndim == 3 else pixels
    _, mask = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)
    count, labels, stats, _ = cv2.connectedComponentsWithStats(mask)
    h, w = mask.shape
    ink_y, ink_x = np.where(mask > 0)
    if not len(ink_x) or count <= 1:
        return None
    ink_left, ink_right = int(ink_x.min()), int(ink_x.max()) + 1
    glyph_height = int(stats[1:, cv2.CC_STAT_HEIGHT].max())
    proposals = []
    for index in range(1, count):
        x, y, cw, ch, area = stats[index]
        if not (
            ink_left + (ink_right - ink_left) * 0.2 < x < ink_left + (ink_right - ink_left) * 0.8
            and ch > glyph_height * 0.4
            and 0.12 < cw / max(ch, 1) < 0.65
            and 0.1 < area / max(cw * ch, 1) < 0.6
        ):
            continue
        ys, xs = np.where(labels[y : y + ch, x : x + cw] == index)
        if len(xs) < 4 or xs.std() < 1 or ys.std() < 1:
            continue
        if float(np.corrcoef(xs, ys)[0, 1]) < -0.8:
            proposals.append(x + cw)
    return int(proposals[0]) if len(proposals) == 1 else None


def recover_digital_fields(image, result, profile, found, engine, cancel, *, layout=None):
    read_line = getattr(engine, "recognize_text_line", None)
    layout = layout or DigitalLayout(result, profile)
    if not callable(read_line):
        return found, layout, {"calls": 0, "failed": False}
    output = list(found)
    calls = 0
    failed = False

    def safe_read(pixels, variant, script):
        nonlocal calls, failed
        if cancel.is_set():
            raise Cancelled()
        calls += 1
        try:
            return read_line(pixels, result.image_id, variant, cancel, script)
        except Cancelled:
            raise
        except AppError:
            failed = True
            return OcrResult([], result.image_id, variant, pixels.shape[1], pixels.shape[0])

    # Independent English recognition is needed only when the multilingual
    # Cyrillic recognizer could not preserve the printed Latin form.
    for name in ("surname", "first_name"):
        values = [
            c
            for c in output
            if c.field_name == name
            and c.validation_status != FieldStatus.INVALID
            and name_script(c.normalized_value) == "latin"
        ]
        band = layout.bands.get(name)
        if values or not band or not band.blocks:
            continue
        x, y, w, h = bounds(band.blocks)
        x, y = max(0, x - 3), max(0, y - 3)
        crop = image[y : min(image.shape[0], y + h + 6), x : min(image.shape[1], x + w + 6)]
        if cancel.is_set():
            raise Cancelled()
        split = bilingual_split(crop)
        latin_pixels = crop[:, max(0, split - 1) :] if split is not None else crop
        reread = safe_read(latin_pixels, f"digital-{name}-latin", "latin")
        pieces = re.split(r"\s*/\s*", reread.text.strip())
        # The Cyrillic half can look like Latin to an English model. Never
        # treat that half as a transliteration of the printed second half.
        if split is not None:
            value = reread.text.strip()
        elif len(pieces) == 2 and pieces[1]:
            value = pieces[1].strip()
        else:
            continue
        value = value.translate(_LATIN_HOMOGLYPHS)
        if name_script(value) != "latin" or validate_field(name, value):
            continue
        item = candidate(layout, name, value, band.blocks, "digital-latin-line")
        if split is not None:
            item.source_bbox = (x + max(0, split - 1), y, latin_pixels.shape[1], latin_pixels.shape[0])
        item.confidence = min((b.confidence for b in reread.blocks), default=0) * 0.95
        item.original_latin = value
        item.engine_name = reread.engine_name
        item.model_version = reread.model_version
        item.variant = reread.variant
        output.append(item)
    for name in ("surname", "first_name", "patronymic"):
        if any(
            c.field_name == name
            and c.validation_status != FieldStatus.INVALID
            and name_script(c.normalized_value) == "cyrillic"
            for c in output
        ):
            continue
        band = layout.bands.get(name)
        if not band or not band.blocks:
            continue
        x, y, w, h = bounds(band.blocks)
        left, top = max(0, x - 3), max(0, y - 3)
        crop = image[top : min(image.shape[0], y + h + 3), left : min(image.shape[1], x + w + 3)]
        split = bilingual_split(crop)
        pixels = crop[:, :split] if split is not None else crop
        reread = safe_read(pixels, f"digital-{name}-cyrillic", "cyrillic")
        from document_excel.infrastructure.field_extractors import cyrillic_field_value

        value = cyrillic_field_value(re.split(r"\s*/\s*", reread.text.strip())[0])
        if name_script(value) == "cyrillic" and not validate_field(name, value):
            item = candidate(layout, name, value, band.blocks, "digital-cyrillic-line")
            item.original_cyrillic = value
            item.engine_name, item.model_version = reread.engine_name, reread.model_version
            item.variant = reread.variant
            item.confidence = min((b.confidence for b in reread.blocks), default=0) * 0.95
            output.append(item)
    # Typed fields reuse the known semantic band. One original line plus one
    # contrast retry is the limit; no missing digits/dates are supplied.
    typed = {
        "birth_date",
        "passport_issued_date",
        "expiration_date",
        "personal_identifier",
        "passport_number",
    }
    for cls in DigitalDocumentExtractor.extractors:
        field = getattr(cls, "field", "")
        if field not in typed or field not in layout.bands:
            continue
        if any(c.field_name == field and c.validation_status != FieldStatus.INVALID for c in output):
            continue
        band = layout.bands[field]
        if not band.blocks:
            continue
        x, y, w, h = bounds(band.blocks)
        crop = image[
            max(0, y - 3) : min(image.shape[0], y + h + 3), max(0, x - 3) : min(image.shape[1], x + w + 3)
        ]
        for variant in ("original", "contrast"):
            if cancel.is_set():
                raise Cancelled()
            pixels = crop
            if variant == "contrast":
                gray = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY) if crop.ndim == 3 else crop
                pixels = cv2.createCLAHE(clipLimit=1.5, tileGridSize=(4, 2)).apply(gray)
            reread = safe_read(pixels, f"digital-{field}-{variant}", "latin")
            context = copy.copy(layout)
            placed = [TextBlock(b.text, b.confidence, band.bbox, b.line_key, b.script) for b in reread.blocks]
            context.bands = {field: DigitalBand(band.bbox, placed)}
            recovered = cls().extract(context)
            useful = [c for c in recovered if c.validation_status != FieldStatus.INVALID]
            if useful:
                for item in useful:
                    item.extraction_method = "digital-region-" + variant
                    item.engine_name = reread.engine_name
                    item.model_version = reread.model_version
                    item.variant = reread.variant
                output.extend(useful)
                break
    band = layout.bands.get("registration_address")
    primary = next(
        (
            c
            for c in output
            if c.field_name == "registration_address" and c.validation_status != FieldStatus.INVALID
        ),
        None,
    )
    if (
        band
        and band.blocks
        and primary
        and str(RegistrationAddressResolver().resolve([primary], label_seen=True)) == "PARTIALLY_FOUND"
    ):
        placed = []
        for row in band.rows()[:8]:
            x, y, w, h = bounds(row)
            left, top = max(0, x - 3), max(0, y - 3)
            crop = image[top : min(image.shape[0], y + h + 3), left : min(image.shape[1], x + w + 3)]
            reread = safe_read(crop, "digital-registration-line", "cyrillic")
            for b in reread.blocks:
                bx, by, bw, bh = b.bbox
                placed.append(
                    TextBlock(b.text, b.confidence, (left + bx, top + by, bw, bh), b.line_key, b.script)
                )
        if placed:
            context = copy.copy(layout)
            context.bands = {"registration_address": DigitalBand(band.bbox, placed)}
            from document_excel.infrastructure.digital_fields import DigitalRegistrationAddressExtractor

            for item in DigitalRegistrationAddressExtractor().extract(context):
                if (
                    item.normalized_value != primary.normalized_value
                    and item.validation_status != FieldStatus.INVALID
                ):
                    item.extraction_method = "digital-registration-lines"
                    output.append(item)
    return output, layout, {"calls": calls, "failed": failed}
