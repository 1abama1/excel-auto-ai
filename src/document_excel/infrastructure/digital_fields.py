"""Digital card semantic bands. Extractors inspect existing blocks, never run OCR."""

import re
from dataclasses import dataclass

from document_excel.domain.models import (
    DocumentSide,
    FieldCandidate,
    FieldStatus,
    normalize,
    split_document_identifier,
    split_mkk_authority,
    validate_field,
)
from document_excel.domain.resolution import name_script
from document_excel.infrastructure.extraction import label_matches, name_values
from document_excel.infrastructure.field_extractors import DateExtractor, cyrillic_field_value
from document_excel.infrastructure.ocr_contract import TextBlock

_LATIN_HOMOGLYPHS = str.maketrans("АВСЕНКМОРТХасеорх", "ABCEHKMOPTXaceopx")


def bounds(blocks):
    left = min(b.bbox[0] for b in blocks)
    top = min(b.bbox[1] for b in blocks)
    right = max(b.bbox[0] + b.bbox[2] for b in blocks)
    bottom = max(b.bbox[1] + b.bbox[3] for b in blocks)
    return left, top, right - left, bottom - top


@dataclass
class DigitalBand:
    bbox: tuple[int, int, int, int]
    blocks: list[TextBlock]

    def rows(self):
        grouped: list[list[TextBlock]] = []
        for block in sorted(self.blocks, key=lambda b: (b.bbox[1] + b.bbox[3] / 2, b.bbox[0])):
            row = next(
                (
                    r
                    for r in grouped
                    if abs(r[0].bbox[1] + r[0].bbox[3] / 2 - block.bbox[1] - block.bbox[3] / 2)
                    < min(r[0].bbox[3], block.bbox[3]) * 0.65
                ),
                None,
            )
            if row is None:
                grouped.append([block])
            else:
                row.append(block)
        return [sorted(row, key=lambda b: b.bbox[0]) for row in grouped]


class DigitalLayout:
    def __init__(self, result, profile):
        self.result, self.profile = result, profile
        self.anchors: dict[str, list[TextBlock]] = {}
        self.inline: dict[str, list[TextBlock]] = {}
        for block in result.blocks:
            matches = label_matches(block.text, profile.aliases)
            for field in {name for name, _ in matches}:
                self.anchors.setdefault(field, []).append(block)
                last = max((m for n, m in matches if n == field), key=lambda m: m.end())
                suffix = block.text[last.end() :].strip(" :;/|-")
                # A slash on the header row separates translated captions. A
                # misspelled translated caption is not an inline name value.
                inline_separator = block.text[last.end() :].lstrip().startswith(":")
                if suffix and inline_separator and not label_matches(suffix, profile.aliases):
                    self.inline.setdefault(field, []).append(
                        TextBlock(suffix, block.confidence, block.bbox, block.line_key)
                    )
        self.bands = {}
        w, h = result.width, result.height
        for field, labels in self.anchors.items():
            # Repeated translated labels form one semantic header. A far-away
            # occurrence of NAME in another label is not a value position.
            labels = sorted(labels, key=lambda b: b.bbox[1])
            first = labels[0]
            group = [
                b
                for b in labels
                if abs(b.bbox[0] - first.bbox[0]) < w * 0.2
                and b.bbox[1] - first.bbox[1] < max(45, first.bbox[3] * 4)
            ]
            x, y, lw, lh = bounds(group)
            left = max(0, x - round(w * 0.025))
            right = w
            top = y + lh
            bottom = h
            for other, blocks in self.anchors.items():
                if other == field:
                    continue
                for b in blocks:
                    bx, by, bw, bh = b.bbox
                    if by >= top - 2 and abs(bx - x) < max(w * 0.16, lw * 0.55):
                        bottom = min(bottom, by)
                    elif abs(by - y) <= max(lh, bh) * 1.5 and bx > x + max(40, w * 0.22):
                        right = min(right, bx - round(w * 0.02))
            selected = list(self.inline.get(field, []))
            for block in result.blocks:
                bx, by, bw, bh = block.bbox
                if (
                    left <= bx + min(bw * 0.2, 10) < right
                    and top - 2 <= by + bh * 0.5 < bottom
                    and not label_matches(block.text, profile.aliases)
                    and (re.search(r"\w", block.text) or block.text.strip() == "/")
                ):
                    selected.append(block)
            self.bands[field] = DigitalBand((left, top, max(0, right - left), max(0, bottom - top)), selected)

    def side(self):
        fields = set(self.anchors)
        front = len(fields & set(self.profile.front_layout))
        back = len(fields & set(self.profile.back_layout))
        return (
            DocumentSide.FRONT
            if front >= 2 and front > back
            else DocumentSide.BACK
            if back >= 2 and back > front
            else DocumentSide.UNKNOWN
        )


def candidate(layout, field, raw, blocks, method="digital-label-band"):
    try:
        value = normalize(field, raw)
        error = validate_field(field, value)
    except ValueError:
        value, error = raw, "Неверный формат значения"
    return FieldCandidate(
        field_name=field,
        raw_value=raw,
        normalized_value=value,
        confidence=min((b.confidence for b in blocks), default=0) * 0.95,
        source_image_id=layout.result.image_id,
        source_side=layout.side(),
        source_bbox=bounds(blocks),
        validation_status=FieldStatus.INVALID if error else FieldStatus.REVIEW,
        validation_errors=[error] if error else [],
        extraction_method=method,
        variant=layout.result.variant,
        engine_name=layout.result.engine_name,
        model_version=layout.result.model_version,
    )


class DigitalNameExtractor:
    def extract(self, layout):
        output = []
        for name in ("surname", "first_name", "patronymic"):
            band_values = []
            band = layout.bands.get(name)
            if not band or not band.blocks:
                continue
            # Both printed scripts may share a baseline or occupy two rows.
            for row in band.rows()[:2]:
                text = " ".join(b.text for b in row)
                bilingual = re.split(r"\s*/\s*", text)
                parts = bilingual if len(bilingual) == 2 and all(bilingual) else name_values(text)
                for index, value in enumerate(parts):
                    expected_script = (
                        ("cyrillic" if index == 0 else "latin")
                        if len(bilingual) == 2 and all(bilingual)
                        else name_script(value)
                    )
                    if expected_script == "cyrillic":
                        value = cyrillic_field_value(value)
                    elif expected_script == "latin":
                        value = value.translate(_LATIN_HOMOGLYPHS)
                    found = candidate(layout, name, value, row)
                    script = name_script(value)
                    if script != expected_script:
                        found.validation_status = FieldStatus.INVALID
                        found.validation_errors.append(
                            "Алфавит имени не соответствует печатной форме; требуется уточнение."
                        )
                    elif script in {"cyrillic", "latin"}:
                        setattr(found, "original_" + script, value)
                    band_values.append(found)
            for script in ("cyrillic", "latin"):
                pieces = [
                    c
                    for c in band_values
                    if c.validation_status != FieldStatus.INVALID
                    and name_script(c.normalized_value) == script
                ]
                # Two wrapped rows of the same printed form constitute one
                # compound name. Alternatives from the same row stay separate.
                if len(pieces) == 2 and pieces[0].source_bbox[1] != pieces[1].source_bbox[1]:
                    merged = candidate(layout, name, " ".join(c.raw_value for c in pieces), band.blocks)
                    setattr(merged, "original_" + script, merged.normalized_value)
                    band_values = [c for c in band_values if c not in pieces] + [merged]
            output.extend(band_values)
        return output


class DigitalValueExtractor:
    field = ""

    def texts(self, band):
        # Separate blocks allow a page number beside a date to remain outside
        # the date candidate, while a spaced date can still use the whole row.
        return [(b.text, [b]) for b in band.blocks] + [(" ".join(b.text for b in r), r) for r in band.rows()]

    def extract(self, layout):
        band = layout.bands.get(self.field)
        if not band or not band.blocks:
            return []
        attempts = []
        for raw, blocks in self.texts(band):
            if self.field in DateExtractor.fields:
                try:
                    raw = DateExtractor().value(self.field, raw)
                except ValueError:
                    continue
            item = candidate(layout, self.field, raw, blocks)
            if item.validation_status != FieldStatus.INVALID:
                attempts.append(item)
        # Preserve an observed invalid candidate for review, never fabricate a
        # missing date or number from neighbouring fields.
        return attempts or [
            candidate(layout, self.field, " ".join(b.text for b in band.rows()[0]), band.rows()[0])
        ]


class DigitalBirthDateExtractor(DigitalValueExtractor):
    field = "birth_date"


class DigitalIssueDateExtractor(DigitalValueExtractor):
    field = "passport_issued_date"


class DigitalExpiryDateExtractor(DigitalValueExtractor):
    field = "expiration_date"


class DigitalPersonalNumberExtractor(DigitalValueExtractor):
    field = "personal_identifier"


class DigitalDocumentNumberExtractor(DigitalValueExtractor):
    field = "passport_number"

    def extract(self, layout):
        output = []
        for item in super().extract(layout):
            parts = split_document_identifier(item.normalized_value)
            if parts:
                prefix, number = parts
                item.normalized_value = number
                item.extraction_method += "+split-type"
                output.append(
                    item.model_copy(
                        update={
                            "field_name": "passport_type",
                            "raw_value": prefix,
                            "normalized_value": prefix,
                        }
                    )
                )
            output.append(item)
        return output


class DigitalIssuingAuthorityExtractor(DigitalValueExtractor):
    field = "passport_issued_by"

    def extract(self, layout):
        band = layout.bands.get(self.field)
        if not band or not band.blocks:
            return []
        rows = band.rows()[:2]
        blocks = [b for row in rows for b in row]
        item = candidate(layout, self.field, " ".join(b.text for b in blocks), blocks)
        parts = split_mkk_authority(item.normalized_value)
        if not parts:
            return [item]
        authority, code = parts
        item.normalized_value = cyrillic_field_value(authority)
        item.extraction_method += "+split-department"
        return [
            item,
            item.model_copy(
                update={"field_name": "passport_department_code", "raw_value": code, "normalized_value": code}
            ),
        ]


class DigitalRegistrationAddressExtractor(DigitalValueExtractor):
    field = "registration_address"

    def extract(self, layout):
        band = layout.bands.get(self.field)
        if not band or not band.blocks:
            return []
        rows = band.rows()
        left = min(b.bbox[0] for b in rows[0])
        rows = [row for row in rows if abs(min(b.bbox[0] for b in row) - left) < layout.result.width * 0.15][
            :8
        ]
        blocks = [b for row in rows for b in row]
        return [candidate(layout, self.field, " ".join(b.text for b in blocks), blocks)]


class DigitalSexExtractor(DigitalValueExtractor):
    field = "gender"


class DigitalCitizenshipExtractor(DigitalValueExtractor):
    field = "citizenship"

    def texts(self, band):
        return [(" ".join(b.text for b in row), row) for row in band.rows()[:2]]


class DigitalDocumentExtractor:
    extractors = (
        DigitalNameExtractor,
        DigitalBirthDateExtractor,
        DigitalPersonalNumberExtractor,
        DigitalDocumentNumberExtractor,
        DigitalIssuingAuthorityExtractor,
        DigitalIssueDateExtractor,
        DigitalExpiryDateExtractor,
        DigitalRegistrationAddressExtractor,
        DigitalSexExtractor,
        DigitalCitizenshipExtractor,
    )

    def extract(self, result, profile):
        layout = DigitalLayout(result, profile)
        found = [c for cls in self.extractors for c in cls().extract(layout)]
        unique: dict[tuple[str, str], FieldCandidate] = {}
        for item in found:
            key = (item.field_name, item.normalized_value)
            if key not in unique or item.confidence > unique[key].confidence:
                unique[key] = item
        return list(unique.values()), layout
