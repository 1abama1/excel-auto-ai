"""Printed FRONT name groups, assembled from geometry rather than OCR word order."""

from __future__ import annotations

from dataclasses import dataclass, field

import cv2
import numpy as np

from document_excel.domain.models import DocumentSide, FieldCandidate, FieldStatus
from document_excel.domain.resolution import name_script
from document_excel.infrastructure.document_profiles import KyrgyzPhysicalIDProfile
from document_excel.infrastructure.extraction import label_matches
from document_excel.infrastructure.field_extractors import NamePartExtractor, cyrillic_field_value
from document_excel.infrastructure.ocr_contract import TextBlock
from document_excel.infrastructure.physical_id import semantic_labels

NAMES = ("surname", "first_name", "patronymic")


@dataclass
class NameGroups:
    candidates: list[FieldCandidate] = field(default_factory=list)
    regions: dict = field(default_factory=dict)
    rows: dict = field(default_factory=dict)


class FullNameExtractor:
    """No customer values, dictionaries, maximum word length, or MRZ input."""

    def locate(self, result):
        labels = semantic_labels(result)
        h, w = result.height, result.width
        regions = {}
        profile = KyrgyzPhysicalIDProfile()
        for name in NAMES:
            rx, ry, rw, rh = profile.ocr_regions["front." + name]
            anchors = [b for b, f in labels if f == name and b.bbox[0] > w * 0.2]
            if anchors:
                first = min(anchors, key=lambda b: b.bbox[1])
                cluster = [b for b in anchors if abs(b.bbox[1] - first.bbox[1]) < h * 0.055]
                left = max(0, min(b.bbox[0] for b in cluster) - round(w * 0.015))
                top = max(b.bbox[1] + b.bbox[3] for b in cluster) + 1
                next_labels = [
                    b.bbox[1] - 1
                    for b, f in labels
                    if f != name and b.bbox[1] > top and abs(b.bbox[0] - first.bbox[0]) < w * 0.3
                ]
                bottom = min(
                    next_labels + [min(h, top + round(h * (0.12 if name != "patronymic" else 0.075)))]
                )
            else:
                left, top = round(rx * w), round(ry * h)
                bottom = round((ry + rh) * h)
            right = min(w, round(w * 0.99))
            if right > left and bottom > top:
                regions[name] = (left, top, right - left, bottom - top)
        return regions

    @staticmethod
    def rows(blocks, region):
        x, y, w, h = region
        selected = [
            b
            for b in blocks
            if x <= b.bbox[0] + b.bbox[2] / 2 <= x + w
            and y <= b.bbox[1] + b.bbox[3] / 2 <= y + h
            and not label_matches(b.text)
        ]
        rows: list[list[TextBlock]] = []
        for block in sorted(selected, key=lambda b: (b.bbox[1], b.bbox[0])):
            row = next(
                (
                    r
                    for r in rows
                    if abs((r[0].bbox[1] + r[0].bbox[3] / 2) - (block.bbox[1] + block.bbox[3] / 2))
                    < min(r[0].bbox[3], block.bbox[3]) * 0.5
                ),
                None,
            )
            if row is None:
                rows.append([block])
            else:
                row.append(block)
        output: list[TextBlock] = []
        for row in rows:
            row.sort(key=lambda b: b.bbox[0])
            text = row[0].text.strip()
            previous = row[0]
            for block in row[1:]:
                gap = block.bbox[0] - (previous.bbox[0] + previous.bbox[2])
                # A very narrow split inside a glyph run is a single word.
                join = gap <= min(previous.bbox[3], block.bbox[3]) * 0.22 or text.endswith("-")
                if name_script(previous.text) != name_script(block.text):
                    # Mixed scripts on one baseline are never fused into a name.
                    join = False
                text += ("" if join else " ") + block.text.strip()
                previous = block
            left = min(b.bbox[0] for b in row)
            top = min(b.bbox[1] for b in row)
            right = max(b.bbox[0] + b.bbox[2] for b in row)
            bottom = max(b.bbox[1] + b.bbox[3] for b in row)
            output.append(
                TextBlock(
                    text,
                    min(b.confidence for b in row),
                    (left, top, right - left, bottom - top),
                    (1, 1, len(output)),
                )
            )
        return output

    @staticmethod
    def candidate(name, raw, confidence, bbox, image_id, script, method):
        extractor = NamePartExtractor()
        extractor.field_name = name
        value = extractor.value(raw, script)
        if not value:
            return None
        return FieldCandidate(
            field_name=name,
            raw_value=raw,
            normalized_value=value,
            confidence=confidence,
            source_image_id=image_id,
            source_side=DocumentSide.FRONT,
            source_bbox=bbox,
            extraction_method=method,
            validation_status=FieldStatus.REVIEW,
            original_cyrillic=value if script == "cyrillic" else None,
            original_latin=value if script == "latin" else None,
        )

    def extract(self, result, side):
        groups = NameGroups()
        if side != DocumentSide.FRONT:
            return groups
        groups.regions = self.locate(result)
        for name, region in groups.regions.items():
            rows = self.rows(result.blocks, region)
            groups.rows[name] = rows
            for index, row in enumerate(rows):
                script = name_script(row.text)
                if index == 0 and (
                    len(rows) >= 2 or name_script(cyrillic_field_value(row.text)) == "cyrillic"
                ):
                    script = "cyrillic"
                if script not in {"cyrillic", "latin"} or name == "patronymic" and script != "cyrillic":
                    continue
                candidate = self.candidate(
                    name, row.text, row.confidence, row.bbox, result.image_id, script, "front-name-group"
                )
                if candidate:
                    candidate.engine_name = result.engine_name
                    candidate.model_version = result.model_version
                    candidate.variant = result.variant
                    groups.candidates.append(candidate)
        return groups

    @staticmethod
    def row_regions(image, region):
        x, y, w, h = region
        crop = image[y : y + h, x : x + w]
        gray = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY) if crop.ndim == 3 else crop
        mask = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
        occupied = np.flatnonzero((mask > 0).sum(axis=1) > max(3, w * 0.008))
        bands = np.split(occupied, np.where(np.diff(occupied) > 2)[0] + 1)
        bands = [b for b in bands if len(b) >= 5]
        return [
            (x, max(y, y + int(b[0]) - 3), w, min(y + h, y + int(b[-1]) + 4) - max(y, y + int(b[0]) - 3))
            for b in bands[:2]
        ]

    def recover(self, image, result, groups, engine, cancel):
        """Only missing/low-confidence rows. Width is the label group, never a partial word."""
        for name, region in groups.regions.items():
            optical_rows = groups.rows.get(name, [])
            readings = self.row_regions(image, region)
            if optical_rows:
                x, y, w, h = region
                readings = []
                for index, row in enumerate(optical_rows):
                    top = max(0, min(y, row.bbox[1] - 2)) if index == 0 else max(y, row.bbox[1] - 2)
                    bottom = min(image.shape[0], y + h, row.bbox[1] + row.bbox[3] + 2)
                    if index + 1 < len(optical_rows):
                        next_row = optical_rows[index + 1]
                        bottom = min(
                            bottom,
                            round(
                                (row.bbox[1] + row.bbox[3] / 2 + next_row.bbox[1] + next_row.bbox[3] / 2) / 2
                            ),
                        )
                    if bottom > top:
                        readings.append((x, top, w, bottom - top))
            for index, bbox in enumerate(readings):
                script = "cyrillic" if index == 0 else "latin"
                if name == "patronymic" and index:
                    continue
                existing = [
                    c
                    for c in groups.candidates
                    if c.field_name == name and name_script(c.normalized_value) == script
                ]
                complete = existing and not self.ink_outside(image, bbox, existing)
                if complete and min(c.confidence for c in existing) >= 0.9:
                    continue
                x, y, w, h = bbox
                pixels = image[y : y + h, x : x + w]
                if h < 36:
                    pixels = cv2.resize(pixels, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
                read = engine.recognize_physical_value(
                    pixels, result.image_id, cancel, "rus" if script == "cyrillic" else "eng", "name"
                )
                raw = " ".join(read.text.split())
                candidate = self.candidate(
                    name,
                    raw,
                    min((b.confidence for b in read.blocks), default=0),
                    bbox,
                    result.image_id,
                    script,
                    "front-name-group-retry",
                )
                if candidate:
                    if existing and min(c.confidence for c in existing) >= 0.9:
                        old = max(existing, key=lambda c: c.confidence)
                        previous = old.normalized_value.replace(" ", "")
                        complete_value = candidate.normalized_value.replace(" ", "")
                        if (
                            candidate.confidence < old.confidence * 0.98
                            or previous not in complete_value
                            or len(complete_value) <= len(previous)
                        ):
                            continue
                    candidate.engine_name = read.engine_name
                    candidate.model_version = read.model_version
                    candidate.variant = read.variant
                    # Replace a weak read of this same optical row with its
                    # complete wide-crop read; other images/MRZ remain independent.
                    groups.candidates = [c for c in groups.candidates if c not in existing] + [candidate]
        return groups

    @staticmethod
    def ink_outside(image, bbox, candidates):
        """A high-confidence partial word still needs its full optical row read."""
        x, y, w, h = bbox
        crop = image[y : y + h, x : x + w]
        gray = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY) if crop.ndim == 3 else crop
        mask = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
        _, _, stats, _ = cv2.connectedComponentsWithStats(mask)
        bounds = [c.source_bbox for c in candidates if c.source_bbox]
        if not bounds:
            return True
        left = min(b[0] for b in bounds)
        right = max(b[0] + b[2] for b in bounds)
        return any(
            h * 0.4 <= ch <= h
            and cw <= h * 1.7
            and area >= 8
            and (
                left - h * 1.5 <= x + cx < left - h * 0.4 or right + h * 0.4 < x + cx + cw <= right + h * 1.5
            )
            for cx, _, cw, ch, area in stats[1:]
        )
