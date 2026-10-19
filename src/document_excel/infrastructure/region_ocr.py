"""Bounded, independent optical reads of official plastic-card fields."""

from __future__ import annotations

import re
from dataclasses import dataclass

import cv2
import numpy as np

from document_excel.domain.errors import AppError, Cancelled
from document_excel.domain.models import DocumentSide, FieldCandidate, FieldStatus
from document_excel.infrastructure.document_profiles import KyrgyzPhysicalIDProfile
from document_excel.infrastructure.preprocessing import deskew


@dataclass(frozen=True)
class FieldRegion:
    bbox: tuple[int, int, int, int]
    locator: str


class DocumentNumberExtractor:
    """KG physical profile: literal ID plus exactly seven observed digits.

    Position-confused prefixes are accepted only when another optical pass on
    the same region actually reads ID with the identical seven-digit suffix.
    No missing digit, D, or PIN is manufactured from the expected format.
    """

    @staticmethod
    def literal(value):
        compact = re.sub(r"\s+", "", value).upper()
        return compact if re.fullmatch(r"ID[0-9]{7}", compact) else None

    @classmethod
    def resolve(cls, reads):
        literal = {cls.literal(read.text) for read, _ in reads} - {None}
        output = []
        for read, variant in reads:
            compact = re.sub(r"\s+", "", read.text).upper()
            value = cls.literal(compact)
            positional = False
            if value is None:
                match = re.fullmatch(r"(?:[1IL|][D0O]|D|[I1L|])[0-9]{7}", compact)
                if match:
                    agreeing = [value for value in literal if value[2:] == compact[-7:]]
                    if len(agreeing) == 1:
                        value, positional = agreeing[0], True
            if value is not None:
                output.append((read, variant, value, positional))
        return output


class RegionBasedOCR:
    fields = {"passport_number", "passport_issued_by"}

    def __init__(self, engine):
        self.engine = engine
        self.profile = KyrgyzPhysicalIDProfile()

    def locate(self, card, result, side, name):
        from document_excel.infrastructure.physical_id import semantic_labels

        h, w = card.image.shape[:2]
        area = self.profile.ocr_regions.get(f"{side.value.lower()}.{name}")
        if area is None:
            return []
        rx, ry, rw, rh = area
        regions = [FieldRegion((round(rx * w), round(ry * h), round(rw * w), round(rh * h)), "profile")]
        labels = semantic_labels(result)
        for label, field in labels:
            x, y, width, height = label.bbox
            if field != name or abs(y / h - ry) > 0.22:
                continue
            left = max(0, min(round(rx * w), x - 4))
            right = min(w, max(x + width + 5, round((rx + rw) * w)))
            # A value printed on the label row has its own word bounds.
            inline = [
                b
                for b in result.blocks
                if b.line_key == label.line_key and b.bbox[0] >= x + width and re.search(r"[0-9]", b.text)
            ]
            if inline:
                right = min(w, max(b.bbox[0] + b.bbox[2] + 4 for b in inline))
                left = max(0, min(b.bbox[0] for b in inline) - 4)
                top = max(0, min(b.bbox[1] for b in inline) - 3)
                bottom = min(h, max(b.bbox[1] + b.bbox[3] for b in inline) + 3)
            else:
                top = max(0, y + height)
                below = [
                    b.bbox[1] - 1
                    for b, f in labels
                    if f != name and b.bbox[1] > top and abs(b.bbox[0] - x) < w * 0.16
                ]
                bottom = min(below + [min(h, top + round(h * max(rh, 0.08)))])
                adjacent = [
                    b.bbox[0] - 3
                    for b, f in labels
                    if f != name and b.bbox[0] > x + width and abs(b.bbox[1] - y) < height * 1.5
                ]
                right = min(adjacent + [right])
            if right - left >= 35 and bottom - top >= 8:
                regions.insert(0, FieldRegion((left, top, right - left, bottom - top), "label"))
        if name == "passport_number":
            # Preserve independently read visible numbers, including the repeat
            # under the barcode whose label can be absent or OCR-invisible.
            for line in result.lines():
                if self.literal_number_line(line.text):
                    x, y, width, height = line.bbox
                    if (
                        side == DocumentSide.FRONT
                        and x > w * 0.45
                        and y > h * 0.5
                        or side == DocumentSide.BACK
                        and x > w * 0.5
                        and y < h * 0.58
                    ):
                        regions.insert(
                            0,
                            FieldRegion(
                                (
                                    max(0, x - 4),
                                    max(0, y - 3),
                                    min(w - x + 4, width + 8),
                                    min(h - y + 3, height + 6),
                                ),
                                "word",
                            ),
                        )
            if side == DocumentSide.BACK:
                barcode = self.barcode_region(card.image)
                if barcode:
                    regions.insert(0, barcode)
        elif name == "passport_issued_by":
            # A six-digit suffix often survives full-image OCR even when the
            # small multilingual label and three letters do not. Its actual
            # baseline locates the entire issuer row without clipping letters
            # or including the label above/below it.
            for line in result.lines():
                x, y, width, height = line.bbox
                if (
                    re.fullmatch(r"[0-9]{6}", line.text.strip())
                    and w * 0.25 < x < w * 0.65
                    and h * 0.2 < y < h * 0.4
                ):
                    left = max(0, min(round(rx * w), x - round(w * 0.14)))
                    top = max(0, y - 3)
                    right = min(w, x + width + 4)
                    bottom = min(h, y + height + 3)
                    regions.insert(0, FieldRegion((left, top, right - left, bottom - top), "code-row"))
        # Prefer semantic/word geometry, keeping one layout fallback. Two
        # distinct proposals bound the total Tesseract work for each field.
        unique = list(dict.fromkeys(regions))
        return unique[:1] + ([unique[-1]] if len(unique) > 1 and unique[-1].bbox != unique[0].bbox else [])

    @staticmethod
    def literal_number_line(value):
        return bool(
            re.fullmatch(r"(?:(?:ID|[1Il|][DO0]|D)\s*[0-9]{7}|[0-9]{9})", value.strip(), re.IGNORECASE)
        )

    @staticmethod
    def barcode_region(image):
        h, w = image.shape[:2]
        left, top = round(w * 0.58), round(h * 0.04)
        sample = image[top : round(h * 0.4), left : round(w * 0.98)]
        gray = cv2.cvtColor(sample, cv2.COLOR_RGB2GRAY)
        mask = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
        joined = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, np.ones((3, max(3, round(w * 0.025))), np.uint8))
        contours, _ = cv2.findContours(joined, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        proposals = []
        for contour in contours:
            x, y, width, height = cv2.boundingRect(contour)
            if width > w * 0.16 and h * 0.035 < height < h * 0.22 and width / height > 2.5:
                stripes = mask[y : y + height, x : x + width].mean(axis=0) > 100
                if np.count_nonzero(np.diff(stripes)) < 14:
                    continue
                proposals.append((width * height, x, y, width, height))
        if not proposals:
            return None
        _, x, y, width, height = max(proposals)
        bx, by = max(left, left + x - 5), top + y + height + 1
        return FieldRegion((bx, by, min(w - bx, width + 10), min(h - by, round(h * 0.11))), "barcode")

    @staticmethod
    def variants(crop):
        gray = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY)
        clahe = cv2.createCLAHE(clipLimit=1.5, tileGridSize=(4, 2)).apply(gray)
        threshold = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 21, 9)
        return [
            ("original", crop),
            ("gray", gray),
            ("clahe", clahe),
            ("adaptive", threshold),
            ("upscale", gray),
        ]

    def extract(self, card, result, cancel, side, *, field_name=None, bbox=None):
        from document_excel.infrastructure.physical_id import clean_value_line

        output = []
        for field in sorted(self.fields):
            if field_name and field != field_name:
                continue
            regions = (
                [FieldRegion(bbox, "selected-area")]
                if bbox is not None
                else self.locate(card, result, side, field)
            )
            for region in regions:
                x, y, width, height = region.bbox
                crop = card.image[y : y + height, x : x + width]
                if crop.size == 0:
                    continue
                # Keep the genuine unprocessed crop as an independent pass.
                # Hough deskew on a short code can follow a glyph diagonal,
                # rotate the baseline and clip strokes, so it is an additional
                # number alternative, never a prerequisite for every read.
                if field == "passport_number":
                    crop = clean_value_line(crop, first_band=True)
                reads = []
                variants = self.variants(crop)
                allowed = getattr(self.engine, "regional_variants", None)
                if allowed:
                    variants = [(name, pixels) for name, pixels in variants if name in allowed]
                if field == "passport_number":
                    rotated = deskew(crop)
                    if not np.array_equal(rotated, crop):
                        variants.append(("deskew", rotated))
                for variant, pixels in variants:
                    if cancel.is_set():
                        raise Cancelled()
                    try:
                        read = self.engine.recognize_official_region(
                            pixels, result.image_id, cancel, field, variant
                        )
                        reads.append((read, variant))
                        # Neural recognition needs only a useful original or
                        # one contrast alternative, not every legacy filter.
                        if (
                            allowed
                            and read.text.strip()
                            and min((b.confidence for b in read.blocks), default=0) >= 0.9
                        ):
                            break
                    except Cancelled:
                        raise
                    except AppError:
                        continue
                if field == "passport_number":
                    resolved = DocumentNumberExtractor.resolve(reads)
                else:
                    from document_excel.infrastructure.field_extractors import IssuingAuthorityExtractor

                    resolved = []
                    for read, variant in reads:
                        value = IssuingAuthorityExtractor().value(read.text)
                        # The short-code profile requires Cyrillic letters and
                        # a six-digit office suffix. Mixed alphabets remain
                        # explicit alternatives, never silently transliterated.
                        if value:
                            resolved.append((read, variant, value, False))
                for read, variant, value, positional in resolved:
                    confidence = min((b.confidence for b in read.blocks), default=0.0)
                    method = f"region-{region.locator}-{variant}"
                    if positional:
                        method += "+corroborated-prefix"
                    base = FieldCandidate(
                        field_name=field,
                        raw_value=read.text,
                        normalized_value=value,
                        confidence=confidence,
                        source_image_id=result.image_id,
                        source_bbox=region.bbox,
                        source_side=side,
                        variant=f"official-{region.locator}-{variant}",
                        extraction_method=method,
                        validation_status=FieldStatus.REVIEW,
                    )
                    if field == "passport_number":
                        base.normalized_value = value[2:]
                        base.extraction_method += "+split-type"
                        output.append(
                            base.model_copy(
                                update={
                                    "field_name": "passport_type",
                                    "normalized_value": "ID",
                                    "raw_value": "ID",
                                }
                            )
                        )
                    output.append(base)
        return output
