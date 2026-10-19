from __future__ import annotations

import json
import re
from dataclasses import dataclass
from dataclasses import field as dataclass_field
from difflib import SequenceMatcher
from importlib.resources import files

from document_excel.domain.models import (
    DocumentSide,
    FieldCandidate,
    FieldStatus,
    normalize,
    split_document_identifier,
    split_mkk_authority,
    validate_field,
)
from document_excel.domain.resolution import (
    NAME_FIELDS,
    name_script,
)
from document_excel.domain.resolution import RegistrationAddressResolver as RegistrationAddressResolver
from document_excel.domain.resolution import address_from_text as address_from_text
from document_excel.infrastructure.ocr import OcrResult, TextBlock


@dataclass
class DocumentProfile:
    id: str
    name: str
    signatures: list[str]
    minimum_signatures: int
    expected_fields: list[str]
    regions: dict
    formats: dict
    strategy: str
    aliases: dict[str, list[str]] = dataclass_field(default_factory=dict)
    supported_languages: list[str] = dataclass_field(default_factory=list)
    front_layout: list[str] = dataclass_field(default_factory=list)
    back_layout: list[str] = dataclass_field(default_factory=list)
    ocr_regions: dict = dataclass_field(default_factory=dict)
    mrz_strategy: str = "none"
    registration_policy: str = "detect"
    optional_fields: list[str] = dataclass_field(default_factory=list)

    def __post_init__(self) -> None:
        if self.strategy != "label-spatial":
            raise ValueError("Unsupported document extraction strategy")
        for region in self.regions.values():
            if len(region) != 4 or not all(0 <= value <= 1 for value in region):
                raise ValueError("Document regions must contain normalized x/y/width/height")


def load_document_profiles() -> list[DocumentProfile]:
    from document_excel.infrastructure.document_profiles import (
        KyrgyzDigitalIDProfile,
        KyrgyzPhysicalIDProfile,
        UnknownDocumentProfile,
    )

    resource = files("document_excel").joinpath("resources/document_profiles.json")
    return [DocumentProfile(**entry) for entry in json.loads(resource.read_text(encoding="utf-8"))] + [
        KyrgyzPhysicalIDProfile(),
        KyrgyzDigitalIDProfile(),
        UnknownDocumentProfile(),
    ]


def classify(result: OcrResult, profiles: list[DocumentProfile], *, digital_screen=False) -> str:
    from document_excel.infrastructure.document_profiles import DocumentTypeClassifier

    return DocumentTypeClassifier().classify(result, profiles, digital_screen=digital_screen)


ALIASES = {
    "place_of_birth": ["ТУУЛГАН ЖЕРИ", "МЕСТО РОЖДЕНИЯ", "PLACE OF BIRTH"],
    "full_name": ["ФИО", "Ф. И. О.", "FULL NAME"],
    "surname": ["ФАМИЛИЯ", "SURNAME", "ФАМИЛИЯСЫ"],
    "first_name": ["ИМЯ", "GIVEN NAMES", "GIVEN NAME", "NAME", "АТЫ"],
    "patronymic": ["ОТЧЕСТВО", "PATRONYMIC", "АТАСЫНЫН АТЫ"],
    "birth_date": ["ДАТА РОЖДЕНИЯ", "DATE OF BIRTH", "ТУУЛГАН КҮНҮ"],
    "personal_identifier": [
        "ПЕРСОНАЛЬНЫЙ НОМЕР",
        "PERSONAL NUMBER",
        "PERSONAL NO",
        "ПИН",
        "ИНН",
        "PIN",
        "ЖЕКЕ НОМЕРУ",
        "ЖЕКЕ НОМЕРИ",
        "ЖЕКЕ ИДЕНТИФИКАЦИЯЛЫК НОМЕР",
        "PERSONAL IDENTIFICATION NUMBER",
    ],
    "passport_number": [
        "НОМЕР ПАСПОРТА",
        "PASSPORT NO",
        "DOCUMENT NO",
        "DOCUMENT NUMBER",
        "НОМЕР ДОКУМЕНТА",
        "ДОКУМЕНТТИК №",
        "№ ДОКУМЕНТА",
        "ДОКУМЕНТТИН НОМЕРИ",
        "ДОКУМЕНТ НОМЕРИ",
        "ДОКУМЕНТ №",
    ],
    "passport_type": ["СЕРИЯ", "ТИП ДОКУМЕНТА", "TYPE"],
    "passport_issued_by": [
        "ОРГАН ВЫДАЧИ",
        "КЕМ ВЫДАН",
        "ISSUING AUTHORITY",
        "AUTHORITY",
        "БЕРГЕН МЕКЕМЕ",
        "БЕРГЕН ОРГАН",
    ],
    "passport_department_code": ["КОД ПОДРАЗДЕЛЕНИЯ", "DEPARTMENT CODE"],
    "passport_issued_date": ["ДАТА ВЫДАЧИ", "DATE OF ISSUE", "ISSUE DATE", "БЕРИЛГЕН КҮНҮ"],
    "registration_address": [
        "АДРЕС РЕГИСТРАЦИИ",
        "МЕСТО РЕГИСТРАЦИИ",
        "REGISTRATION ADDRESS",
        "ПРОПИСКА",
        "РЕГИСТРАЦИЯ",
        "АДРЕС ПРОПИСКИ",
        "КАТТАЛГАН ЖЕРИ",
        "КАТТОО ДАРЕГИ",
        "ТУРУКТУУ ЖАШАГАН ЖЕРИ",
    ],
    "expiration_date": [
        "ДАТА ОКОНЧАНИЯ",
        "БҮТКӨН КҮНҮ",
        "БУТКОН КУНУ",
        "ДЕЙСТВИТЕЛЕН ДО",
        "DATE OF EXPIRY",
        "EXPIRY DATE",
        "СРОК ДЕЙСТВИЯ",
        "КОЛДОНУУ МӨӨНӨТҮ",
    ],
    "gender": ["ПОЛ", "SEX", "GENDER", "ЖЫНЫСЫ"],
    "citizenship": ["ГРАЖДАНСТВО", "NATIONALITY", "CITIZENSHIP", "ЖАРАНДЫГЫ"],
    "_ethnicity": ["УЛУТУ", "НАЦИОНАЛЬНОСТЬ", "ETHNICITY"],
    "_signature": ["ПОДПИСЬ", "SIGNATURE", "КОЛУ"],
}

TYPED_FIELDS = {
    "birth_date",
    "passport_issued_date",
    "expiration_date",
    "personal_identifier",
    "passport_number",
}


def label_matches(text: str, aliases: dict[str, list[str]] | None = None) -> list[tuple[str, re.Match]]:
    found = []
    for field, labels in (aliases if aliases is not None else ALIASES).items():
        for label in labels:
            # Tiny OCR labels often lose spaces: DocumentNo / Dateofissue.
            pattern = r"(?<!\w)" + re.escape(label).replace(r"\ ", r"\s*") + r"(?!\w)"
            for match in re.finditer(pattern, text, re.IGNORECASE):
                found.append((field, match))
    # Longer semantic labels take precedence over overlapping short ones (NAME/FULL NAME).
    found.sort(key=lambda pair: len(pair[1].group()), reverse=True)
    accepted: list[tuple[str, re.Match]] = []
    for field, match in found:
        if not any(match.start() < old.end() and match.end() > old.start() for _, old in accepted):
            accepted.append((field, match))
    return sorted(accepted, key=lambda pair: pair[1].start())


class DocumentSideClassifier:
    """Conservative side detection; titles or a single label never establish a side."""

    def classify(self, result: OcrResult, document_type: str = "") -> DocumentSide:
        if document_type == "kg_physical_id":
            from document_excel.infrastructure.document_profiles import physical_side

            return physical_side(result)
        fields = {field for field, _ in label_matches(result.text)}
        front = fields & {"surname", "first_name", "patronymic", "birth_date", "gender", "citizenship"}
        back = fields & {
            "passport_number",
            "personal_identifier",
            "passport_issued_by",
            "passport_issued_date",
            "expiration_date",
            "registration_address",
        }
        if len(front) >= 2 and len(front) > len(back):
            return DocumentSide.FRONT
        if len(back) >= 2 and len(back) > len(front):
            return DocumentSide.BACK
        return DocumentSide.UNKNOWN


def name_values(value: str) -> list[str]:
    """Keep printed forms as alternatives, without inventing a transliteration."""
    parts = [part.strip() for part in re.split(r"\s*[/|;]\s*", value) if part.strip()]
    output: list[str] = []
    for part in parts:
        words = part.split()
        # Bilingual names may also be printed without a slash. Do not split a mixed-script word.
        scripts = [name_script(word) for word in words]
        if "cyrillic" in scripts and "latin" in scripts and "other" not in scripts:
            output.extend(
                " ".join(w for w, script in zip(words, scripts, strict=True) if script == kind)
                for kind in ("cyrillic", "latin")
            )
        else:
            output.append(part)
    return list(dict.fromkeys(output))


class FieldExtractor:
    def __init__(self, profile: DocumentProfile | None = None) -> None:
        self.profile = profile
        self.aliases = {**ALIASES, **(profile.aliases if profile else {})}

    def extract(self, result: OcrResult, quality_warning: bool = False) -> list[FieldCandidate]:
        if self.profile is not None and self.profile.id == "kg_digital_id":
            from document_excel.infrastructure.digital_fields import DigitalDocumentExtractor

            return DigitalDocumentExtractor().extract(result, self.profile)[0]
        lines = result.lines()
        candidates = []
        side = DocumentSideClassifier().classify(result)
        explicit_citizenship = bool(re.search(r"(?i)\b(?:ГРАЖДАНСТВО|ЖАРАНДЫГЫ|CITIZENSHIP)\b", result.text))
        for index, line in enumerate(lines):
            matches = label_matches(line.text, self.aliases)
            for match_index, (field, match) in enumerate(matches):
                if field.startswith("_"):
                    continue
                if field == "citizenship" and (
                    any(name == "_ethnicity" for name, _ in matches)
                    or explicit_citizenship
                    and match.group().upper() == "NATIONALITY"
                ):
                    # Nationality is citizenship on many passports, but a separate ethnicity on KG cards.
                    continue
                end = (
                    matches[match_index + 1][1].start() if match_index + 1 < len(matches) else len(line.text)
                )
                value = line.text[match.end() : end].strip(" :;|.-/")
                source = line
                method = "label-inline"
                next_same_field = match_index + 1 < len(matches) and matches[match_index + 1][0] == field
                if not value and next_same_field:
                    continue
                if field in NAME_FIELDS:
                    # A partly misread translated label can survive before the printed colon.
                    value = value.rsplit(":", 1)[-1].strip()
                    suffix = line.text[match.end() : end].lstrip()
                    if (
                        next_same_field
                        and suffix.startswith(("/", "|"))
                        and any(
                            value.upper() in alias.upper().split() for alias in self.aliases.get(field, [])
                        )
                    ):
                        continue
                    repeated_label = sum(name == field for name, _ in matches) >= 2
                    if repeated_label and suffix.startswith(("/", "|")) and ":" not in suffix:
                        nearby_name = self._nearby_value(lines, index, line, field)
                        if nearby_name and {name_script(v) for v in name_values(nearby_name.text)} == {
                            "cyrillic",
                            "latin",
                        }:
                            source = nearby_name
                            value = nearby_name.text.strip()
                            method = "label-spatial"
                needs_spatial = not value or field in TYPED_FIELDS and not re.search(r"\d", value)
                if needs_spatial:
                    if match_index + 1 < len(matches) and matches[match_index + 1][0] == field:
                        # Аты / Имя / Name is one multilingual label, not three empty fields.
                        continue
                    nearby = self._nearby_value(lines, index, line, field)
                    if nearby is None:
                        if not value or self._label_continuation(value, field):
                            continue
                    else:
                        source = nearby
                        value = nearby.text.strip()
                        method = "label-spatial"
                if field in {"registration_address", "passport_issued_by", "place_of_birth"} | NAME_FIELDS:
                    value, source = self._continuation(lines, source, value, field, result.width)
                if method == "label-spatial" and field == "citizenship":
                    source = self._aligned_value(result, source, line)
                    value = source.text
                confidence = min(line.confidence, source.confidence) * (
                    0.9 if method == "label-spatial" else 0.95
                )
                confidence *= 0.7 if quality_warning else 1
                region = self.profile.regions.get(field) if self.profile else None
                if region:
                    rx, ry, rw, rh = region
                    center_x = (source.bbox[0] + source.bbox[2] / 2) / result.width
                    center_y = (source.bbox[1] + source.bbox[3] / 2) / result.height
                    if not (rx <= center_x <= rx + rw and ry <= center_y <= ry + rh):
                        confidence *= 0.6
                # Warped coordinates cannot be used as original-image coordinates.
                bbox = None if result.variant in {"perspective", "deskewed", "roi"} else source.bbox
                values = name_values(value) if field in NAME_FIELDS else [self._typed_value(field, value)]
                for raw_value in values:
                    error = validate_field(field, raw_value)
                    try:
                        normalized = normalize(field, raw_value)
                    except ValueError:
                        normalized = raw_value
                    if field == "passport_number" and not re.search(r"\d", normalized):
                        error = "Цифры номера не распознаны. Проверьте строку с номером документа."
                    pattern = self.profile.formats.get(field) if self.profile else None
                    if pattern and not re.fullmatch(pattern, normalized):
                        error = "Значение не соответствует формату профиля документа"
                    status = (
                        FieldStatus.INVALID
                        if error
                        else FieldStatus.VALID
                        if confidence >= 0.82
                        else FieldStatus.REVIEW
                    )
                    identifier_parts = (
                        split_document_identifier(raw_value)
                        if field == "passport_number" and not error
                        else None
                    )
                    if (
                        field == "passport_number"
                        and not identifier_parts
                        and re.search(r"[OОIІ|l]", raw_value)
                        and not error
                    ):
                        status = FieldStatus.REVIEW
                    if field == "passport_number" and normalized.isdigit() and not error:
                        # Tiny glyphs can lose the printed series; never fabricate its letters.
                        status = FieldStatus.REVIEW
                    script = name_script(raw_value) if field in NAME_FIELDS else ""
                    candidate = FieldCandidate(
                        field_name=field,
                        raw_value=raw_value,
                        normalized_value=normalized,
                        confidence=confidence,
                        source_image_id=result.image_id,
                        source_bbox=bbox,
                        extraction_method=method,
                        validation_status=status,
                        variant=result.variant,
                        source_side=side,
                        original_cyrillic=raw_value if script == "cyrillic" else None,
                        original_latin=raw_value if script == "latin" else None,
                    )
                    if identifier_parts:
                        prefix, number = identifier_parts
                        candidate.normalized_value = number
                        candidate.extraction_method += "+split-type"
                        candidates.append(
                            candidate.model_copy(
                                update={
                                    "field_name": "passport_type",
                                    "raw_value": prefix,
                                    "normalized_value": prefix,
                                }
                            )
                        )
                    authority_parts = (
                        split_mkk_authority(raw_value)
                        if field == "passport_issued_by" and not error
                        else None
                    )
                    if authority_parts:
                        authority, department_code = authority_parts
                        candidate.normalized_value = authority
                        candidate.extraction_method += "+split-department"
                        candidates.append(
                            candidate.model_copy(
                                update={
                                    "field_name": "passport_department_code",
                                    "raw_value": department_code,
                                    "normalized_value": department_code,
                                }
                            )
                        )
                    candidates.append(candidate)
        return candidates

    @staticmethod
    def _typed_value(field: str, value: str) -> str:
        if field in {"birth_date", "passport_issued_date", "expiration_date"}:
            dates = re.findall(
                r"(?<!\d)(?:\d{4}[./-]\d{1,2}[./-]\d{1,2}|\d{1,2}(?:[./-]|\s+)\d{1,2}(?:[./-]|\s+)\d{4})(?!\d)",
                value,
            )
            if len(dates) == 1:
                return dates[0]
        if field in {"passport_number", "personal_identifier"}:
            return value.strip(" :;,./|")
        return value

    def _label_continuation(self, text: str, field: str) -> bool:
        """A nearby translation may be misspelled by OCR; only inspect known same-field labels."""
        if field not in TYPED_FIELDS or re.search(r"\d", text):
            return False
        compact = re.sub(r"[\W_]+", "", text).casefold()
        if len(compact) < 6:
            return False
        for alias in self.aliases.get(field, []):
            target = re.sub(r"[\W_]+", "", alias).casefold()
            if len(target) >= 6 and SequenceMatcher(None, compact, target).ratio() >= 0.8:
                return True
        return False

    @staticmethod
    def _aligned_value(result: OcrResult, source: TextBlock, label: TextBlock) -> TextBlock:
        """Keep value-column words when OCR joins an unrelated left-column label to the row."""
        words = [block for block in result.blocks if block.line_key == source.line_key]
        threshold = label.bbox[0] - max(15, label.bbox[3])
        aligned = [word for word in words if word.bbox[0] >= threshold]
        if len(aligned) == len(words) or not aligned:
            return source
        anchor = min(aligned, key=lambda word: word.bbox[1])
        aligned = [
            word
            for word in aligned
            if abs(word.bbox[1] + word.bbox[3] / 2 - anchor.bbox[1] - anchor.bbox[3] / 2)
            <= max(word.bbox[3], anchor.bbox[3]) * 0.75
        ]
        left = min(word.bbox[0] for word in aligned)
        if left > label.bbox[0] + max(40, label.bbox[2] / 2):
            return source
        top = min(word.bbox[1] for word in aligned)
        right = max(word.bbox[0] + word.bbox[2] for word in aligned)
        bottom = max(word.bbox[1] + word.bbox[3] for word in aligned)
        return TextBlock(
            " ".join(word.text for word in sorted(aligned, key=lambda word: word.bbox[0])),
            min(word.confidence for word in aligned),
            (left, top, right - left, bottom - top),
            source.line_key,
        )

    def _continuation(
        self, lines: list[TextBlock], source: TextBlock, value: str, field: str, width: int
    ) -> tuple[str, TextBlock]:
        source_index = lines.index(source)
        maximum = (
            8
            if field == "registration_address"
            else 2
            if field in {"passport_issued_by", "place_of_birth"}
            else 1
        )
        previous = source
        for following in lines[source_index + 1 : source_index + 1 + maximum]:
            gap = following.bbox[1] - (previous.bbox[1] + previous.bbox[3])
            if (
                label_matches(following.text, self.aliases)
                or gap > max(35, previous.bbox[3] * 2)
                or gap < -5
                or abs(following.bbox[0] - previous.bbox[0]) > width * 0.15
                or re.search(r"(?i)\b(?:QR|ПОДЕЛИТЬСЯ|СКАЧАТЬ|НАЗАД|SHARE|DOWNLOAD)\b", following.text)
            ):
                break
            if field in NAME_FIELDS and (
                name_script(value) not in {"cyrillic", "latin"}
                or {name_script(value), name_script(following.text)} != {"cyrillic", "latin"}
            ):
                break
            value += (" / " if field in NAME_FIELDS else " ") + following.text
            left = min(source.bbox[0], following.bbox[0])
            right = max(source.bbox[0] + source.bbox[2], following.bbox[0] + following.bbox[2])
            source = TextBlock(
                value,
                min(source.confidence, following.confidence),
                (left, source.bbox[1], right - left, following.bbox[1] + following.bbox[3] - source.bbox[1]),
                source.line_key,
            )
            previous = following
        return value, source

    def _nearby_value(
        self, lines: list[TextBlock], index: int, label: TextBlock, field: str
    ) -> TextBlock | None:
        x, y, w, h = label.bbox
        candidates = []
        for line in lines[max(0, index - 2) : index + 5]:
            if line is label:
                continue
            lx, ly, _lw, lh = line.bbox
            matches = label_matches(line.text, self.aliases)
            if (
                matches
                and field == "citizenship"
                and lx < x
                and all(name == "_signature" for name, _ in matches)
            ):
                # The value column is recovered from word boxes by _aligned_value.
                matches = []
            if matches:
                if all(name == field for name, _ in matches):
                    continue
                if ly > y and abs(lx - x) < max(w, 100):
                    break
                continue
            if not re.search(r"\w", line.text):
                continue
            if self._label_continuation(line.text, field):
                continue
            right = abs(ly - y) < max(h, lh) * 0.7 and lx >= x + w
            below = 0 <= ly - (y + h) <= max(h * 2.5, 40) and abs(lx - x) < max(w, 100)
            if right or below:
                distance = abs(ly - y) * 2 + abs(lx - x - (w if right else 0))
                rank = 0
                if field in TYPED_FIELDS:
                    typed = self._typed_value(field, line.text)
                    rank = 2 if not re.search(r"\d", typed) else 1 if validate_field(field, typed) else 0
                candidates.append((rank, distance, line))
        return min(candidates, key=lambda entry: (entry[0], entry[1]))[2] if candidates else None


def apply_candidates(case, candidates, document_type, label_seen=False):
    from document_excel.domain.resolution import FieldResolver

    FieldResolver().resolve(case, candidates, document_type, label_seen)
