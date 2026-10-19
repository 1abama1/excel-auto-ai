"""Specialized policies over the existing label/spatial extraction framework."""

import re

from document_excel.domain.models import DocumentSide, DocumentType
from document_excel.infrastructure.extraction import ALIASES, DocumentProfile, label_matches


class KyrgyzPhysicalIDProfile(DocumentProfile):
    def __init__(self):
        super().__init__(
            id=DocumentType.KG_PHYSICAL_ID,
            name="Пластиковая ID-карта КР",
            signatures=["KYRGYZ", "КЫРГЫЗ", "IDENTITY CARD"],
            minimum_signatures=2,
            expected_fields=["surname", "first_name", "birth_date", "passport_number", "personal_identifier"],
            regions={},
            formats={"personal_identifier": "[0-9]{14}"},
            strategy="label-spatial",
            aliases={"place_of_birth": ["ТУУЛГАН ЖЕРИ", "МЕСТО РОЖДЕНИЯ", "PLACE OF BIRTH"]},
            supported_languages=["rus", "eng", "kir"],
            front_layout=[
                "surname",
                "first_name",
                "patronymic",
                "birth_date",
                "passport_number",
                "expiration_date",
            ],
            back_layout=[
                "place_of_birth",
                "passport_issued_by",
                "passport_issued_date",
                "personal_identifier",
            ],
            ocr_regions={
                "mrz": (0.0, 0.57, 1.0, 0.43),
                "front.surname": (0.38, 0.265, 0.59, 0.085),
                "front.first_name": (0.38, 0.385, 0.59, 0.085),
                "front.patronymic": (0.38, 0.51, 0.59, 0.055),
                "front.gender": (0.38, 0.59, 0.22, 0.06),
                "front.citizenship": (0.38, 0.665, 0.59, 0.065),
                "front.birth_date": (0.38, 0.75, 0.29, 0.075),
                "front.passport_number": (0.67, 0.755, 0.29, 0.085),
                "front.expiration_date": (0.67, 0.87, 0.29, 0.075),
                "back.place_of_birth": (0.255, 0.11, 0.35, 0.16),
                "back.passport_issued_by": (0.25, 0.26, 0.38, 0.105),
                "back.passport_issued_date": (0.255, 0.39, 0.35, 0.075),
                "back.personal_identifier": (0.64, 0.23, 0.34, 0.075),
                "back.passport_number": (0.69, 0.40, 0.29, 0.075),
            },
            mrz_strategy="TD1",
            registration_policy="not_printed",
            optional_fields=["patronymic", "place_of_birth"],
        )


class KyrgyzDigitalIDProfile(DocumentProfile):
    def __init__(self):
        super().__init__(
            id=DocumentType.KG_DIGITAL_ID,
            name="Цифровое удостоверение КР",
            signatures=["КЫРГЫЗ", "KYRGYZ", "ЦИФРОВОЕ", "ТҮНДҮК"],
            minimum_signatures=2,
            expected_fields=[
                "surname",
                "first_name",
                "birth_date",
                "personal_identifier",
                "passport_number",
                "passport_issued_by",
                "passport_issued_date",
                "expiration_date",
                "registration_address",
                "gender",
                "citizenship",
            ],
            regions={},
            formats={"personal_identifier": "[0-9]{14}"},
            strategy="label-spatial",
            aliases={
                **ALIASES,
                "citizenship": ["ГРАЖДАНСТВО", "CITIZENSHIP", "ЖАРАНДЫГЫ"],
                "_ethnicity": ["УЛУТУ", "НАЦИОНАЛЬНОСТЬ", "NATIONALITY", "ETHNICITY"],
                "passport_number": ALIASES["passport_number"] + ["DOCUMENT №", "ДОКУМЕНТТИН №"],
            },
            supported_languages=["rus", "eng", "kir"],
            front_layout=[
                "first_name",
                "surname",
                "patronymic",
                "gender",
                "citizenship",
                "_ethnicity",
                "birth_date",
            ],
            back_layout=[
                "passport_issued_by",
                "passport_issued_date",
                "personal_identifier",
                "passport_number",
                "expiration_date",
                "registration_address",
            ],
            registration_policy="detect",
            optional_fields=["patronymic"],
        )


class UnknownDocumentProfile(DocumentProfile):
    def __init__(self):
        super().__init__("unknown", "Другой документ", [], 999, [], {}, {}, "label-spatial")


class DocumentTypeClassifier:
    def classify(self, result, profiles, *, digital_screen=False) -> str:
        text = result.text.upper()
        fields = {name for name, _ in label_matches(text)}
        country = any(value in text for value in ("КЫРГЫЗ", "KYRGYZ", "KGZ"))
        title = any(
            value in text for value in ("IDENTITY CARD", "IDENTITY", "ИДЕНТИФИКАЦИЯЛЫК", "ИДЕНТИФИКАЦИОННАЯ")
        )
        mrz = bool(re.search(r"[IAC][A-Z<]KGZ[A-Z0-9<]{18,}", re.sub(r"[ \t]", "", text)))
        portrait_fields = len(
            fields & {"surname", "first_name", "patronymic", "birth_date", "citizenship", "gender"}
        )
        if digital_screen and country and len(fields) >= 2:
            return DocumentType.KG_DIGITAL_ID
        if mrz or (
            country
            and title
            and portrait_fields >= 4
            and bool({"passport_number", "expiration_date"} & fields)
            and 1.25 < result.width / max(1, result.height) < 2.05
        ):
            return DocumentType.KG_PHYSICAL_ID
        if (
            country
            and portrait_fields + len(fields & {"personal_identifier", "passport_issued_by"}) >= 2
            and any(value in text for value in ("ЦИФРОВОЕ", "DIGITAL ID", "ТҮНДҮК", "TUNDUK"))
        ):
            return DocumentType.KG_DIGITAL_ID
        # Preserve legacy kg_id / registration behavior and imported profiles.
        scores = [
            (sum(marker in text for marker in p.signatures), p)
            for p in profiles
            if p.id not in {DocumentType.KG_PHYSICAL_ID, DocumentType.KG_DIGITAL_ID, "unknown"}
        ]
        eligible = sorted(
            ((n, p) for n, p in scores if n >= p.minimum_signatures), key=lambda pair: pair[0], reverse=True
        )
        if not eligible or len(eligible) > 1 and eligible[0][0] == eligible[1][0]:
            return "unknown"
        return eligible[0][1].id


def physical_side(result) -> DocumentSide:
    text = result.text.upper()
    if re.search(r"[IAC][A-Z<]KGZ[A-Z0-9<]{18,}", re.sub(r"[ \t]", "", text)):
        return DocumentSide.BACK
    fields = {name for name, _ in label_matches(text)}
    front = len(fields & {"surname", "first_name", "birth_date", "citizenship", "gender", "expiration_date"})
    back = len(
        fields & {"personal_identifier", "passport_issued_by", "passport_issued_date", "place_of_birth"}
    )
    if front >= 3 and front > back:
        return DocumentSide.FRONT
    if back >= 2 and back > front:
        return DocumentSide.BACK
    return DocumentSide.UNKNOWN
