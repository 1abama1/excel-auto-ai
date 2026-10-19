"""Geometry and dynamic field crops for plastic cards, independent of Excel."""

from __future__ import annotations

import re
import threading
from dataclasses import dataclass

import cv2
import numpy as np

from document_excel.domain.models import (
    FieldCandidate,
    FieldStatus,
    normalize,
    split_document_identifier,
    validate_field,
)
from document_excel.infrastructure.document_profiles import KyrgyzPhysicalIDProfile
from document_excel.infrastructure.extraction import FieldExtractor, label_matches
from document_excel.infrastructure.ocr import TextBlock
from document_excel.infrastructure.preprocessing import order_corners


@dataclass
class PlasticCard:
    image: np.ndarray
    inverse: np.ndarray
    original_size: tuple[int, int]
    detected: bool

    def original_bbox(self, bbox):
        if bbox is None:
            return None
        x, y, width, height = bbox
        points = np.array(
            [[[x, y], [x + width - 1, y], [x + width - 1, y + height - 1], [x, y + height - 1]]], np.float32
        )
        corners = cv2.perspectiveTransform(points, self.inverse)[0]
        left, top = np.floor(corners.min(axis=0)).astype(int)
        right, bottom = np.ceil(corners.max(axis=0)).astype(int) + 1
        left, top = max(0, left), max(0, top)
        right, bottom = min(self.original_size[0], right), min(self.original_size[1], bottom)
        return int(left), int(top), int(max(0, right - left)), int(max(0, bottom - top))


def rotate_plastic_card(card: PlasticCard, turns: int) -> PlasticCard:
    for _ in range(turns % 4):
        h = card.image.shape[0]
        inverse_rotation = np.array([[0, 1, 0], [-1, 0, h - 1], [0, 0, 1]], np.float64)
        card = PlasticCard(
            cv2.rotate(card.image, cv2.ROTATE_90_CLOCKWISE),
            card.inverse @ inverse_rotation,
            card.original_size,
            card.detected,
        )
    return card


def prepare_plastic_card(original: np.ndarray) -> PlasticCard:
    h, w = original.shape[:2]
    scale = min(1, 1300 / max(h, w))
    sample = cv2.resize(original, (round(w * scale), round(h * scale)))
    gray = cv2.cvtColor(sample, cv2.COLOR_RGB2GRAY)
    edge = cv2.Canny(cv2.GaussianBlur(gray, (5, 5), 0), 25, 90)
    edge = cv2.morphologyEx(edge, cv2.MORPH_CLOSE, np.ones((5, 5), np.uint8))
    contours, _ = cv2.findContours(edge, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    proposals = []
    for contour in sorted(contours, key=cv2.contourArea, reverse=True)[:30]:
        area = cv2.contourArea(contour)
        if not gray.size * 0.25 < area < gray.size * 0.995:
            continue
        polygon = cv2.approxPolyDP(contour, 0.025 * cv2.arcLength(contour, True), True)
        if len(polygon) != 4 or not cv2.isContourConvex(polygon):
            continue
        corners = order_corners(polygon) / scale
        width = max(np.linalg.norm(corners[1] - corners[0]), np.linalg.norm(corners[2] - corners[3]))
        height = max(np.linalg.norm(corners[3] - corners[0]), np.linalg.norm(corners[2] - corners[1]))
        if min(width, height) >= 160 and 1.25 < max(width, height) / min(width, height) < 2.05:
            proposals.append((area, corners, round(width), round(height)))
    if not proposals:
        return PlasticCard(original.copy(), np.eye(3), (w, h), False)
    _, corners, width, height = proposals[0]
    target = np.array([[0, 0], [width - 1, 0], [width - 1, height - 1], [0, height - 1]], np.float32)
    matrix = cv2.getPerspectiveTransform(corners, target)
    card = cv2.warpPerspective(original, matrix, (width, height), borderValue=(255, 255, 255))
    prepared = PlasticCard(card, np.linalg.inv(matrix), (w, h), True)
    return rotate_plastic_card(prepared, 1) if height > width else prepared


def prepare_region(image: np.ndarray, kind: str) -> np.ndarray:
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY) if image.ndim == 3 else image
    if kind == "number":
        return cv2.addWeighted(gray, 1.7, cv2.GaussianBlur(gray, (0, 0), 1.2), -0.7, 0)
    if kind == "date":
        return cv2.createCLAHE(clipLimit=1.5, tileGridSize=(4, 2)).apply(gray)
    return cv2.bilateralFilter(cv2.createCLAHE(clipLimit=2, tileGridSize=(8, 8)).apply(gray), 5, 35, 35)


def semantic_labels(result):
    """Alias-word bounds exclude unrelated values joined into the same OCR row."""
    output = []
    for line in result.lines():
        words = sorted(
            (word for word in result.blocks if word.line_key == line.line_key), key=lambda b: b.bbox[0]
        )
        spans = []
        cursor = 0
        for word in words:
            spans.append((cursor, cursor + len(word.text), word))
            cursor += len(word.text) + 1
        for field, match in label_matches(line.text):
            selected = [word for start, end, word in spans if start < match.end() and end > match.start()]
            if not selected:
                continue
            x = min(b.bbox[0] for b in selected)
            y = min(b.bbox[1] for b in selected)
            right = max(b.bbox[0] + b.bbox[2] for b in selected)
            bottom = max(b.bbox[1] + b.bbox[3] for b in selected)
            output.append(
                (
                    TextBlock(
                        line.text,
                        min(b.confidence for b in selected),
                        (x, y, right - x, bottom - y),
                        line.line_key,
                    ),
                    field,
                )
            )
    return output


def clean_value_line(image: np.ndarray, *, first_band=False) -> np.ndarray:
    """Keep a substantial text line and its left-anchored run, excluding nearby art."""
    if image.size == 0:
        return image
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    mask = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
    if first_band:
        occupied = np.flatnonzero((mask[:, : max(1, round(mask.shape[1] * 0.65))] > 0).sum(axis=1) > 2)
        groups = np.split(occupied, np.where(np.diff(occupied) > 2)[0] + 1)
        groups = [g for g in groups if len(g) >= max(6, image.shape[0] * 0.28)]
        if groups:
            group = groups[0]
            top = max(0, int(group[0]) - 2)
            bottom = min(image.shape[0], int(group[-1]) + 3)
            image = image[top:bottom].copy()
            mask = mask[top:bottom]
    h, w = mask.shape
    _, _, stats, _ = cv2.connectedComponentsWithStats(mask)
    boxes = sorted(
        [
            (x, y, width, height)
            for x, y, width, height, area in stats[1:]
            if height > h * 0.35 and width < h * 3 and area > max(3, h * 0.15)
        ],
        key=lambda b: b[0],
    )
    if not boxes or boxes[0][0] > w * 0.15:
        return image
    right = boxes[0][0] + boxes[0][2]
    count = 1
    for x, _y, width, _height in boxes[1:]:
        if x - right > max(6, h * 0.85):
            break
        right = max(right, x + width)
        count += 1
    if count < 2:
        return image
    return image[:, max(0, boxes[0][0] - 3) : min(w, right + 3)].copy()


def regional_candidates(card: PlasticCard, result, engine, cancel: threading.Event) -> list[FieldCandidate]:
    """Locate values below recognized labels, never coordinates from a specimen."""
    method = getattr(engine, "recognize_card_region", None)
    if not callable(method):
        return []
    labels = semantic_labels(result)
    output = []
    seen = set()
    h, w = card.image.shape[:2]
    for line, name in labels:
        if (
            name not in {"birth_date", "expiration_date", "passport_issued_date", "passport_number"}
            or name in seen
        ):
            continue
        seen.add(name)
        x, y, width, height = line.bbox
        # Inline VIZ values are already handled by the ordinary extractor.
        if re.search(r"\d{4}", line.text):
            continue
        right = min(
            [
                other.bbox[0] - 2
                for other, _ in labels
                if other.bbox[0] > x + width * 0.4 and abs(other.bbox[1] - y) < height * 2
            ]
            + [w]
        )
        bottom = min(
            [
                other.bbox[1] - 1
                for other, _ in labels
                if other.bbox[1] > y + height and abs(other.bbox[0] - x) < max(width, w * 0.15)
            ]
            + [min(h, y + height + max(height * 2, round(h * 0.11)))]
        )
        left, top = max(0, x - 3), min(h, y + height)
        if right - left < 45 or bottom - top < 8:
            continue
        region = card.image[top:bottom, left:right]
        kind = "number" if name == "passport_number" else "date"
        read = method(prepare_region(region, kind), result.image_id, cancel, kind)
        value = " ".join(read.text.split())
        if kind == "number":
            match = re.fullmatch(r"[A-ZА-Я]{1,4}\s*[0-9]{4,12}", value, re.IGNORECASE)
            if not match:
                continue
        else:
            try:
                value = normalize(name, value)
            except ValueError:
                continue
        # Reuse existing prefix split/format validation via a synthetic label block.
        read_confidence = min((block.confidence for block in read.blocks), default=0.0)
        read.blocks = [
            TextBlock(
                f"{FieldExtractor().aliases[name][0]}: {value}",
                read_confidence,
                (left, top, right - left, bottom - top),
                (1, 1, 1),
            )
        ]
        read.width, read.height = w, h
        for candidate in FieldExtractor(KyrgyzPhysicalIDProfile()).extract(read):
            candidate.extraction_method = "physical-region-" + kind
            if name == "passport_number" and split_document_identifier(value):
                candidate.extraction_method += "+split-type"
            candidate.validation_status = FieldStatus.REVIEW
            output.append(candidate)
    return output


def layout_candidates(card, result, engine, cancel, side, fields=None):
    """Re-read pixels in normalized layout proposals, refined by printed labels."""
    method = getattr(engine, "recognize_physical_value", None)
    if not callable(method) or side.value not in {"FRONT", "BACK"}:
        return []
    profile = KyrgyzPhysicalIDProfile()
    h, w = card.image.shape[:2]
    labels = semantic_labels(result)
    output = []
    specialized = []
    if callable(getattr(engine, "recognize_official_region", None)):
        from document_excel.infrastructure.region_ocr import RegionBasedOCR

        if fields is None:
            specialized = RegionBasedOCR(engine).extract(card, result, cancel, side)
        else:
            for field in sorted(fields & {"passport_issued_by", "passport_number"}):
                specialized.extend(
                    RegionBasedOCR(engine).extract(card, result, cancel, side, field_name=field)
                )
    specialized_fields = {c.field_name for c in specialized}
    for key, area in profile.ocr_regions.items():
        prefix, _, name = key.partition(".")
        if prefix != side.value.lower():
            continue
        if fields is not None and name not in fields:
            continue
        if name in specialized_fields:
            continue
        rx, ry, rw, rh = area
        x, y, right, bottom = round(rx * w), round(ry * h), round((rx + rw) * w), round((ry + rh) * h)
        anchors = [line for line, field in labels if field == name and abs(line.bbox[1] / h - ry) < 0.15]
        if anchors and name not in {
            "birth_date",
            "expiration_date",
            "passport_number",
            "passport_issued_date",
            "personal_identifier",
        }:
            anchor = min(anchors, key=lambda line: line.bbox[1])
            # A recognized English translation may start to the right of the
            # multilingual label and its printed value. Keep the value-column
            # proposal so the first letters are never cut off by that alias.
            x = min(x, max(0, anchor.bbox[0] - 3))
            y = min(h, anchor.bbox[1] + anchor.bbox[3])
            next_labels = [
                other.bbox[1] - 2
                for other, field in labels
                if field != name and other.bbox[1] > y and abs(other.bbox[0] - anchor.bbox[0]) < w * 0.15
            ]
            bottom = min(next_labels + [min(h, y + round((rh + 0.04) * h))])
        crop = card.image[y : min(h, bottom), x : min(w, right)]
        if crop.size == 0:
            continue
        if name in {"surname", "first_name"}:
            gray = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY)
            mask = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
            text_column = mask[:, : max(1, round(crop.shape[1] * 0.5))]
            occupied = np.flatnonzero((text_column > 0).sum(axis=1) > max(2, crop.shape[1] * 0.005))
            groups = np.split(occupied, np.where(np.diff(occupied) > 2)[0] + 1)
            groups = [g for g in groups if len(g) >= 4]
            if len(groups) == 2:
                readings = [
                    (crop[max(0, int(g[0]) - 1) : min(crop.shape[0], int(g[-1]) + 2)], lang)
                    for g, lang in zip(groups, ("rus", "eng"), strict=True)
                ]
            else:
                middle = crop.shape[0] // 2
                readings = [(crop[:middle], "rus"), (crop[middle:], "eng")]
        else:
            readings = [
                (
                    crop,
                    "rus"
                    if name in {"patronymic", "citizenship", "place_of_birth", "passport_issued_by", "gender"}
                    else "eng",
                )
            ]
        for pixels, lang in readings:
            kind = (
                "name"
                if name in {"surname", "first_name", "patronymic"}
                else "date"
                if name in {"birth_date", "expiration_date", "passport_issued_date"}
                else "text"
            )
            if name in {
                "surname",
                "first_name",
                "patronymic",
                "birth_date",
                "expiration_date",
                "passport_issued_date",
                "passport_number",
                "gender",
            }:
                pixels = clean_value_line(pixels, first_band=name not in {"surname", "first_name"})
            read = method(pixels, result.image_id, cancel, lang, kind)
            value = " ".join(read.text.split())
            if kind == "name" and lang == "rus" and not re.fullmatch(r"[А-ЯЁҮӨҢа-яёүөң -]{2,}", value):
                alternate = getattr(engine, "recognize_cyrillic_name", None)
                if callable(alternate):
                    alternate_read = alternate(pixels, result.image_id, cancel)
                    alternate_value = " ".join(alternate_read.text.split())
                    if re.fullmatch(r"[А-ЯЁҮӨҢа-яёүөң -]{2,}", alternate_value):
                        read, value = alternate_read, alternate_value
            if name == "passport_number" and read.blocks:
                from document_excel.infrastructure.mrz import MRZRecognizer

                left = min(b.bbox[0] for b in read.blocks)
                top = min(b.bbox[1] for b in read.blocks)
                right_edge = max(b.bbox[0] + b.bbox[2] for b in read.blocks)
                bottom_edge = max(b.bbox[1] + b.bbox[3] for b in read.blocks)
                word = pixels[max(0, top - 2) : bottom_edge + 2, max(0, left - 2) : right_edge + 2]
                glyphs, _ = MRZRecognizer.glyphs(word, expected=9, check_spacing=False)
                if glyphs:
                    split_x = (glyphs[1][0] + glyphs[1][2] + glyphs[2][0]) // 2
                    prefix = method(word[:, :split_x], result.image_id, cancel, "eng", "name").text.strip()
                    digits = method(word[:, split_x:], result.image_id, cancel, "eng", "digits").text.strip()
                    if re.fullmatch(r"[A-Z]{2}", prefix) and re.fullmatch(r"[0-9]{7}", digits):
                        value = prefix + digits
            if name == "passport_number" and not split_document_identifier(value):
                identifier_method = getattr(engine, "recognize_identifier_line", None)
                if callable(identifier_method):
                    reread = identifier_method(pixels, result.image_id, "physical-number", cancel)
                    if split_document_identifier(reread.text.strip()):
                        value = reread.text.strip()
            if name in {"birth_date", "expiration_date", "passport_issued_date"}:
                from document_excel.infrastructure.field_extractors import DateExtractor

                try:
                    value = DateExtractor().value(name, value)
                except ValueError:
                    continue
            if not value:
                continue
            if label_matches(value):
                continue
            observed_name = value
            if name in {"surname", "first_name", "patronymic"}:
                from document_excel.infrastructure.field_extractors import (
                    FirstNameExtractor,
                    PatronymicExtractor,
                    SurnameExtractor,
                )

                extractor = {
                    "surname": SurnameExtractor,
                    "first_name": FirstNameExtractor,
                    "patronymic": PatronymicExtractor,
                }[name]()
                value = extractor.value(value, "cyrillic" if lang == "rus" else "latin")
                if not value:
                    continue
            if name == "passport_number" and not re.fullmatch(r"[A-Z]{1,4}\s*[0-9]{4,12}", value):
                continue
            if name == "personal_identifier" and not re.fullmatch(r"[0-9]{14}", value.replace(" ", "")):
                continue
            if name == "gender" and value.upper().replace(" ", "") not in {
                "Э/М",
                "Э/М.",
                "А/F",
                "М",
                "Ж",
                "M",
                "F",
            }:
                continue
            if validate_field(name, value):
                continue
            read_confidence = min((block.confidence for block in read.blocks), default=0.0)
            read.blocks = [
                TextBlock(
                    profile.aliases.get(name, FieldExtractor().aliases.get(name, [name]))[0] + ": " + value,
                    read_confidence,
                    (x, y, right - x, bottom - y),
                    (1, 1, 1),
                )
            ]
            read.width, read.height = w, h
            for candidate in FieldExtractor(profile).extract(read):
                candidate.extraction_method += "+physical-layout"
                if (
                    name in {"surname", "first_name", "patronymic"}
                    and lang == "rus"
                    and observed_name != value
                ):
                    candidate.raw_value = observed_name
                    candidate.extraction_method += "+cyrillic-homoglyphs"
                candidate.validation_status = FieldStatus.REVIEW
                output.append(candidate)
    output.extend(specialized)
    return output


def normalize_physical_name_scripts(candidates, width, height, side):
    """Disambiguate identical glyph codes only inside the printed Cyrillic row.

    The lower Latin row and MRZ are never transliterated. Preserve raw OCR and
    require review; no phonetic, dictionary or customer-specific correction.
    """
    from document_excel.infrastructure.extraction import name_script
    from document_excel.infrastructure.field_extractors import cyrillic_field_value

    if side.value != "FRONT":
        return candidates
    profile = KyrgyzPhysicalIDProfile()
    for candidate in candidates:
        if (
            candidate.field_name not in {"surname", "first_name", "patronymic"}
            or candidate.source_bbox is None
        ):
            continue
        x, y, w, h = candidate.source_bbox
        rx, ry, rw, rh = profile.ocr_regions["front." + candidate.field_name]
        center_x, center_y = (x + w / 2) / width, (y + h / 2) / height
        limit = rh if candidate.field_name == "patronymic" else rh / 2
        if not (rx <= center_x <= rx + rw and ry <= center_y <= ry + limit):
            continue
        value = cyrillic_field_value(candidate.normalized_value)
        if value == candidate.normalized_value or name_script(value) != "cyrillic":
            continue
        candidate.normalized_value = value
        candidate.original_cyrillic, candidate.original_latin = value, None
        candidate.extraction_method += "+cyrillic-row-homoglyphs"
        candidate.validation_status = FieldStatus.REVIEW
    return candidates
