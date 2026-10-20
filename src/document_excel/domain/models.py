from __future__ import annotations

import hashlib
import json
import re
from datetime import UTC, date, datetime
from enum import StrEnum
from pathlib import Path
from typing import Any
from uuid import uuid4

from pydantic import BaseModel, Field

from document_excel.domain.person_name import PersonNameFormatter


def new_id() -> str:
    return uuid4().hex


class RegistrationStatus(StrEnum):
    FOUND = "FOUND"
    PARTIALLY_FOUND = "PARTIALLY_FOUND"
    NOT_PRESENT = "NOT_PRESENT"
    NOT_RECOGNIZED = "NOT_RECOGNIZED"
    UNCERTAIN = "UNCERTAIN"
    MANUALLY_ENTERED = "MANUALLY_ENTERED"
    USER_CONFIRMED = "USER_CONFIRMED"


class FieldStatus(StrEnum):
    MISSING = "MISSING"
    NOT_PRESENT = "NOT_PRESENT"
    NOT_RECOGNIZED = "NOT_RECOGNIZED"
    INVALID = "INVALID"
    REVIEW = "REVIEW"
    VALID = "VALID"
    CONFLICT = "CONFLICT"
    MANUALLY_CORRECTED = "MANUALLY_CORRECTED"
    CONFIRMED = "CONFIRMED"


class DocumentSide(StrEnum):
    UNKNOWN = "UNKNOWN"
    FRONT = "FRONT"
    BACK = "BACK"
    ADDITIONAL = "ADDITIONAL"


class DocumentType(StrEnum):
    UNKNOWN = "unknown"
    KG_PHYSICAL_ID = "kg_physical_id"
    KG_DIGITAL_ID = "kg_digital_id"


class WorkflowMode(StrEnum):
    EXISTING_CONTRACT = "EXISTING_CONTRACT"
    TEMPLATE_GENERATION = "TEMPLATE_GENERATION"


class MRZReport(BaseModel):
    image_id: str
    lines: list[str] = Field(default_factory=list, repr=False)
    checks: dict[str, bool] = Field(default_factory=dict)
    segment_status: dict[str, str] = Field(
        default_factory=lambda: {
            name: "NOT_CHECKED" for name in ("document_number", "birth_date", "expiration_date", "composite")
        }
    )
    values: dict[str, str] = Field(default_factory=dict, repr=False)
    errors: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    conflicts: list[str] = Field(default_factory=list)
    source_bbox: tuple[int, int, int, int] | None = None
    image_version: str = "original"

    @property
    def valid(self) -> bool:
        return not self.errors and len(self.checks) == 4 and all(self.checks.values())


class Address(BaseModel):
    country: str = ""
    region: str = ""
    city: str = ""
    district: str = ""
    street: str = ""
    house: str = ""
    apartment: str = ""
    extra: str = ""
    raw: str = ""

    def display(self) -> str:
        parts = [self.country, self.region, self.city, self.district, self.street]
        parts += [
            f"д. {self.house}" if self.house else "",
            f"кв. {self.apartment}" if self.apartment else "",
            self.extra,
        ]
        return ", ".join(p for p in parts if p) or self.raw

    def legacy_street(self) -> str:
        return (
            ", ".join(
                p
                for p in [
                    self.country,
                    self.city,
                    self.district,
                    self.street,
                    f"д. {self.house}" if self.house else "",
                    f"кв. {self.apartment}" if self.apartment else "",
                    self.extra,
                ]
                if p
            )
            or self.raw
        )


class ClientData(BaseModel):
    full_name: str | None = None
    surname: str | None = None
    first_name: str | None = None
    patronymic: str | None = None
    surname_cyrillic: str | None = None
    surname_latin: str | None = None
    first_name_cyrillic: str | None = None
    first_name_latin: str | None = None
    patronymic_cyrillic: str | None = None
    place_of_birth: str | None = None
    birth_date: date | None = None
    personal_identifier: str | None = None
    passport_type: str | None = None
    passport_number: str | None = None
    passport_issued_by: str | None = None
    passport_department_code: str | None = None
    passport_issued_date: date | None = None
    registration_address: Address = Field(default_factory=Address)
    living_address: Address = Field(default_factory=Address)
    whatsapp_phone: str | None = None
    additional_phone: str | None = None
    object_address: str | None = None
    gender: str | None = None
    citizenship: str | None = None
    expiration_date: date | None = None
    document_type: str = "unknown"
    notes: str | None = None

    @property
    def document_number(self) -> str | None:
        """Full document identifier; contract cells retain separate type/digits."""
        if not self.passport_number:
            return None
        return (self.passport_type or "") + self.passport_number


class FieldCandidate(BaseModel):
    field_name: str
    raw_value: str
    normalized_value: str
    confidence: float = Field(ge=0, le=1)
    source_image_id: str
    source_bbox: tuple[int, int, int, int] | None = None
    extraction_method: str = "label-spatial"
    validation_status: FieldStatus = FieldStatus.REVIEW
    variant: str = "original"
    source_side: DocumentSide = DocumentSide.UNKNOWN
    original_cyrillic: str | None = None
    original_latin: str | None = None
    image_version: str = "original"
    source_method: str = "VISUAL_OCR"
    validation_errors: list[str] = Field(default_factory=list)
    run_id: str = ""
    image_hash: str = ""
    engine_name: str = "unknown"
    model_version: str = ""


class FieldEvidence(BaseModel):
    source: str = "нет данных"
    status: FieldStatus = FieldStatus.MISSING
    candidates: list[FieldCandidate] = Field(default_factory=list)
    message: str = ""
    original_cyrillic: str | None = None
    original_latin: str | None = None
    resolution_reason: str = ""
    manual_conflict: bool = False
    expected_image_id: str = ""
    expected_bbox: tuple[int, int, int, int] | None = None
    expected_image_version: str = "original"

    @property
    def source_method(self) -> str:
        if "ручн" in self.source or "пользоват" in self.source:
            return "MANUAL_INPUT"
        return (
            "MRZ"
            if self.candidates and all(c.source_method == "MRZ" for c in self.candidates)
            else "VISUAL_OCR"
        )

    @property
    def user_confirmed(self) -> bool:
        return self.status == FieldStatus.CONFIRMED


class SourceImage(BaseModel):
    id: str = Field(default_factory=new_id)
    path: Path
    rotation: int = 0
    role: DocumentSide = DocumentSide.UNKNOWN
    processed_path: Path | None = None
    processed_confirmed: bool = False
    captured: bool = False
    photo_digest: str | None = None
    case_id: str = ""

    @property
    def selected_path(self) -> Path:
        return self.processed_path if self.processed_confirmed and self.processed_path else self.path

    @property
    def selected_rotation(self) -> int:
        return 0 if self.processed_confirmed and self.processed_path else self.rotation


class ToolItem(BaseModel):
    name: str
    quantity: int = Field(default=1, ge=1, le=100000)
    inventory_number: str = ""
    price_per_day: float = Field(default=0, ge=0, le=1e12)


class ContractData(BaseModel):
    tools: list[ToolItem] = Field(default_factory=list)
    deposit_amount: float | None = Field(default=None, ge=0, le=1e12)
    purchase_price: float | None = Field(default=None, ge=0, le=1e12)
    actual_return_date: date | None = None
    actual_return_time: str | None = None


class OcrRun(BaseModel):
    id: str = Field(default_factory=new_id)
    case_id: str
    image_hashes: dict[str, str]
    read_hashes: dict[str, str] = Field(default_factory=dict)
    read_images: list[SourceImage] = Field(default_factory=list, repr=False)
    started_at: datetime = Field(default_factory=lambda: datetime.now(UTC))
    completed_at: datetime | None = None
    status: str = "RUNNING"
    fields: list[str] = Field(default_factory=list)
    scope: str = "full"
    document_type: str = "unknown"
    manual_type: bool = False


class ClientCase(BaseModel):
    id: str = Field(default_factory=new_id)
    name: str = "Новый клиент"
    images: list[SourceImage] = Field(default_factory=list)
    data: ClientData = Field(default_factory=ClientData)
    contract: ContractData = Field(default_factory=ContractData)
    evidence: dict[str, FieldEvidence] = Field(default_factory=dict)
    registration_status: RegistrationStatus = RegistrationStatus.UNCERTAIN
    status: str = "Готов к загрузке"
    errors: list[str] = Field(default_factory=list)
    confirmed_digest: str | None = None
    template_path: Path | None = None
    # Legacy service callers retain template generation; desktop creates cases
    # explicitly in EXISTING_CONTRACT mode.
    workflow_mode: WorkflowMode = WorkflowMode.TEMPLATE_GENERATION
    workbook_profile: dict[str, Any] | None = None
    output_path: Path | None = None
    photo_exception: bool = False
    document_type_manual: bool = False
    mrz_reports: dict[str, MRZReport] = Field(default_factory=dict, repr=False)
    side_proposals: dict[str, DocumentSide] = Field(default_factory=dict)
    ocr_runs: dict[str, OcrRun] = Field(default_factory=dict, repr=False)
    active_run_id: str | None = None
    diagnostic_events: list[dict] = Field(default_factory=list, repr=False)
    identity_status: str = "not_checked"
    identity_message: str = ""
    identity_approved: str = ""
    diagnostic_blocks: dict[str, list[dict]] = Field(default_factory=dict, repr=False)
    processing_metrics: dict[str, dict] = Field(default_factory=dict, repr=False)
    field_regions: dict[str, dict] = Field(default_factory=dict, repr=False)
    pending_identity: Any = Field(default=None, exclude=True, repr=False)

    def digest(self) -> str:
        payload = [
            self.workflow_mode.value,
            str(self.template_path)
            if self.template_path and self.workflow_mode == WorkflowMode.EXISTING_CONTRACT
            else None,
            self.workbook_profile,
            self.formatted_name,
            self.data.model_dump(mode="json"),
            self.contract.model_dump(mode="json"),
            self.registration_status.value,
            self.photo_exception,
            self.document_type_manual,
            self.side_proposals,
            self.identity_status,
            self.identity_approved,
            {k: v.model_dump(mode="json") for k, v in self.mrz_reports.items()},
            [i.model_dump(mode="json") for i in self.images],
            {k: v.status.value for k, v in self.evidence.items()},
        ]
        return hashlib.sha256(json.dumps(payload, sort_keys=True, ensure_ascii=False).encode()).hexdigest()

    @property
    def formatted_name(self) -> str:
        evidence = self.evidence.get("full_name")
        if evidence and evidence.source_method == "MANUAL_INPUT":
            return PersonNameFormatter.format(self.data.full_name, preserve=True)
        parts = [self.data.surname, self.data.first_name, self.data.patronymic]
        composed = " ".join(p for p in parts if p)
        if composed and composed == self.data.full_name:
            return " ".join(
                PersonNameFormatter.format(
                    value,
                    preserve=bool(
                        self.evidence.get(key) and self.evidence[key].source_method == "MANUAL_INPUT"
                    ),
                )
                for key, value in zip(("surname", "first_name", "patronymic"), parts, strict=True)
                if value
            )
        return PersonNameFormatter.format(self.data.full_name)

    @property
    def confirmed(self) -> bool:
        return self.confirmed_digest == self.digest()

    def invalidate(self) -> None:
        self.confirmed_digest = None
        self.status = "На проверке"


class DocumentBatch(BaseModel):
    cases: list[ClientCase] = Field(default_factory=list)


DATE_FIELDS = {"birth_date", "passport_issued_date", "expiration_date"}
FIELD_LABELS = {
    "full_name": "Полное ФИО",
    "surname": "Фамилия",
    "first_name": "Имя",
    "patronymic": "Отчество",
    "surname_cyrillic": "Фамилия (кириллица)",
    "surname_latin": "Фамилия (латиница)",
    "first_name_cyrillic": "Имя (кириллица)",
    "first_name_latin": "Имя (латиница)",
    "patronymic_cyrillic": "Отчество (кириллица)",
    "place_of_birth": "Место рождения",
    "birth_date": "Дата рождения",
    "personal_identifier": "ПИН / ИНН",
    "passport_type": "Тип / серия документа",
    "passport_number": "Номер документа",
    "passport_issued_by": "Орган выдачи",
    "passport_department_code": "Код подразделения",
    "passport_issued_date": "Дата выдачи",
    "registration_address": "Адрес регистрации",
    "living_address": "Фактический адрес",
    "whatsapp_phone": "WhatsApp / телефон",
    "additional_phone": "Дополнительный телефон",
    "object_address": "Адрес объекта",
    "gender": "Пол",
    "citizenship": "Гражданство",
    "expiration_date": "Действителен до",
    "notes": "Примечание",
}


def parse_date(value: str) -> date:
    value = value.strip()
    for pattern, order in [
        (r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", (0, 1, 2)),
        (r"(\d{1,2})(?:[./-]|\s+)(\d{1,2})(?:[./-]|\s+)(\d{4})", (2, 1, 0)),
    ]:
        match = re.fullmatch(pattern, value)
        if match:
            parts = [int(x) for x in match.groups()]
            try:
                return date(*(parts[i] for i in order))
            except ValueError:
                break
    raise ValueError("Введите существующую дату: ДД.ММ.ГГГГ.")


def normalize(field: str, value: str) -> str:
    value = " ".join(value.split())
    if field in DATE_FIELDS:
        return parse_date(value).isoformat()
    if field == "personal_identifier":
        return re.sub(r"\s", "", value)
    if field in {"passport_type", "passport_number"}:
        return value.upper().replace(" ", "")
    if field == "gender":
        token = re.sub(r"[\s/.]+", "", value).casefold()
        if token in {"м", "m", "эм", "эркек", "male", "мужской"}:
            return "MALE"
        if token in {"ж", "f", "аf", "аял", "female", "женский"}:
            return "FEMALE"
        if token in {"unknown", "не определён", "не определен", "<"}:
            return "UNKNOWN"
        return value
    if field == "citizenship":
        if value.casefold() in {
            "kgz",
            "кыргыз республикасы",
            "кыргызская республика",
            "кыргызстан",
            "киргизская республика",
            "kyrgyz republic",
            "kyrgyzstan",
        }:
            return "KGZ"
        return value
    return value


def split_document_identifier(value: str) -> tuple[str, str] | None:
    """Separate a visible letter prefix from its numeric body without repairing OCR."""
    match = re.fullmatch(r"\s*([A-Za-zА-Яа-яЁё]{1,4})[\s-]*([0-9]{4,12})\s*", value)
    return (match[1].upper(), match[2]) if match else None


def split_mkk_authority(value: str) -> tuple[str, str] | None:
    """The digits after the printed MKK authority are its department code."""
    match = re.fullmatch(r"\s*([МM][КK][КK])[\s:№-]*([0-9]{1,12})\s*", value, re.IGNORECASE)
    return (match[1].upper(), match[2]) if match else None


def validate_field(field: str, value: str, pin_length: int = 14) -> str | None:
    if not value.strip():
        return "Значение отсутствует"
    if len(value) > 2000 or any(ord(c) < 32 and c not in "\n\t" for c in value):
        return "Недопустимая длина или управляющие символы"
    try:
        normalized = normalize(field, value)
    except ValueError as exc:
        return str(exc)
    if field == "personal_identifier" and not re.fullmatch(rf"[0-9]{{{pin_length}}}", normalized):
        return f"Ожидается {pin_length} цифр; исправьте по документу"
    if field == "gender" and normalized not in {"MALE", "FEMALE", "UNKNOWN"}:
        return "Не удалось определить пол. Проверьте документ."
    if field == "citizenship" and not (
        re.fullmatch(r"[A-Z]{3}", normalized)
        or re.fullmatch(r"[А-ЯЁҮӨҢа-яёүөңA-Za-z][А-ЯЁҮӨҢа-яёүөңA-Za-z '-]{3,80}", normalized)
    ):
        return "Не удалось определить гражданство. Проверьте документ."
    if field in {
        "surname",
        "first_name",
        "patronymic",
        "surname_cyrillic",
        "first_name_cyrillic",
        "patronymic_cyrillic",
        "surname_latin",
        "first_name_latin",
    } and not re.fullmatch(r"[А-ЯЁҮӨҢа-яёүөңA-Za-z][А-ЯЁҮӨҢа-яёүөңA-Za-z '\u2019-]{1,99}", normalized):
        return "Не удалось прочитать часть имени. Сверьте соответствующую подпись."
    if field == "passport_number" and not re.fullmatch(r"[A-ZА-ЯЁ0-9-]{4,20}", normalized):
        return "Проверьте номер документа (4–20 букв/цифр)"
    if field == "passport_number" and re.fullmatch(rf"[0-9]{{{pin_length}}}", normalized):
        return "Похоже на ПИН. Проверьте отдельный номер документа."
    if field in {"whatsapp_phone", "additional_phone"} and not re.fullmatch(r"\+?[\d ()-]{7,25}", value):
        return "Проверьте формат телефона"
    if field in DATE_FIELDS and parse_date(value) > date.today() and field != "expiration_date":
        return "Дата не может быть в будущем"
    return None


def data_value(data: ClientData, field: str) -> str:
    value = getattr(data, field)
    if field == "gender":
        return {"MALE": "Мужской", "FEMALE": "Женский", "UNKNOWN": "Не определён"}.get(
            str(value), str(value or "")
        )
    if field == "citizenship" and value == "KGZ":
        return "Кыргызская Республика"
    if isinstance(value, Address):
        return value.display()
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    return str(value) if value is not None else ""
