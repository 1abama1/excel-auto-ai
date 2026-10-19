"""Semantic admission and resolution of evidence; independent of Qt and OCR."""

from __future__ import annotations

import re

from document_excel.domain.models import (
    Address,
    ClientCase,
    DocumentSide,
    FieldCandidate,
    FieldEvidence,
    FieldStatus,
    RegistrationStatus,
    normalize,
    validate_field,
)


def manual(evidence):
    return evidence.status == FieldStatus.MANUALLY_CORRECTED or any(
        s in evidence.source.casefold() for s in ("ручн", "пользоват")
    )


def displayed_candidate(evidence, value):
    """Evidence for the value currently shown, independent of confidence order."""
    if manual(evidence) or not value:
        return None
    matches = []
    for candidate in evidence.candidates:
        try:
            comparable = normalize(candidate.field_name, str(value)).strip().upper()
        except ValueError:
            continue
        if candidate.normalized_value.strip().upper() == comparable:
            matches.append(candidate)
    return max(
        matches,
        key=lambda c: (c.source_method == "VISUAL_OCR", c.source_side == DocumentSide.FRONT, c.confidence),
        default=None,
    )


class FieldResolver:
    front = {"surname", "first_name", "patronymic", "birth_date", "expiration_date", "gender", "citizenship"}
    back = {"passport_issued_by", "passport_issued_date", "personal_identifier"}

    def resolve(self, case, candidates, document_type, label_seen=False):

        admitted = []
        for original in candidates:
            if original.extraction_method == "composed-name":
                continue
            candidate = original.model_copy(deep=True)
            if candidate.field_name not in type(case.data).model_fields:
                continue
            errors = list(candidate.validation_errors)
            if candidate.validation_status == FieldStatus.INVALID and not errors:
                errors.append("Кандидат отклонён при извлечении; требуется повторное чтение.")
            try:
                candidate.normalized_value = normalize(candidate.field_name, candidate.normalized_value)
            except ValueError:
                errors.append("Неверный формат значения")
            error = validate_field(candidate.field_name, candidate.normalized_value)
            if error:
                errors.append(error)
            if document_type == "kg_physical_id" and candidate.source_method != "MRZ":
                if candidate.field_name == "citizenship" and not (
                    candidate.normalized_value == "KGZ"
                    or re.fullmatch(r"[A-Z]{3}", candidate.normalized_value)
                ):
                    errors.append(
                        "Не удалось прочитать официальное наименование гражданства; сверяйте с кодом MRZ."
                    )
                expected = (
                    DocumentSide.FRONT
                    if candidate.field_name in self.front
                    else DocumentSide.BACK
                    if candidate.field_name in self.back
                    else None
                )
                if expected and candidate.source_side not in {
                    expected,
                    DocumentSide.UNKNOWN,
                    DocumentSide.ADDITIONAL,
                }:
                    errors.append("Поле прочитано с неподходящей стороны документа")
                hint = case.field_regions.get(candidate.field_name)
                if (
                    hint
                    and candidate.field_name in {"surname", "first_name", "patronymic"}
                    and candidate.source_bbox
                    and hint.get("image_id") == candidate.source_image_id
                ):
                    x, y, w, h = candidate.source_bbox
                    a, b, c, d = hint["bbox"]
                    if not (
                        a - w * 0.1 <= x + w / 2 <= a + c + w * 0.1
                        and b - h * 0.25 <= y + h / 2 <= b + d + h * 0.25
                    ):
                        errors.append("Текст находится вне области соответствующей части имени")
            candidate.validation_errors = list(dict.fromkeys(errors))
            if errors:
                candidate.validation_status = FieldStatus.INVALID
            else:
                candidate.validation_status = FieldStatus.REVIEW
            admitted.append(candidate)
        _apply_validated_candidates(case, admitted, document_type, label_seen)
        for name, evidence in case.evidence.items():
            valid = [
                c
                for c in evidence.candidates
                if c.validation_status != FieldStatus.INVALID and not c.validation_errors
            ]
            if manual(evidence):
                current = getattr(case.data, name, None)
                if isinstance(current, Address):
                    current = current.display()
                try:
                    current = normalize(name, str(current or ""))
                except ValueError:
                    current = ""
                evidence.manual_conflict = bool(valid and any(c.normalized_value != current for c in valid))
                if evidence.manual_conflict:
                    evidence.message = (
                        "Ручное значение сохранено; новый OCR с ним расходится. Сверьте варианты."
                    )
                continue
            if evidence.candidates and not valid:
                if name in type(case.data).model_fields:
                    setattr(
                        case.data,
                        name,
                        Address()
                        if name.endswith("_address") and name in {"registration_address", "living_address"}
                        else None,
                    )
                evidence.status = FieldStatus.INVALID
                evidence.message = (
                    valid[0].validation_errors[0] if valid else "Не удалось определить. Проверьте документ."
                )
            evidence.resolution_reason = (
                "Независимые источники расходятся; требуется выбор пользователя"
                if evidence.status == FieldStatus.CONFLICT
                else "Привязка к полю и формат проверены; требуется сверка с оригиналом"
                if valid
                else "Допустимый кандидат не найден"
            )
        self.resolve_name_parts(case)
        self.compose_name(case)
        if document_type == "kg_physical_id":
            from document_excel.domain.models import FIELD_LABELS, RegistrationStatus

            expected_fields = self.front | self.back | {"passport_number", "passport_type"}
            for field in expected_fields:
                evidence = case.evidence.setdefault(field, FieldEvidence())
                if not getattr(case.data, field, None) and not evidence.candidates and not manual(evidence):
                    evidence.status = FieldStatus.NOT_RECOGNIZED
                    evidence.message = (
                        "Не удалось распознать: "
                        + FIELD_LABELS[field]
                        + ". Покажите область, повторите чтение или введите вручную."
                    )
            for field, hint in case.field_regions.items():
                evidence = case.evidence.setdefault(field, FieldEvidence())
                evidence.expected_image_id = hint["image_id"]
                evidence.expected_bbox = hint["bbox"]
                evidence.expected_image_version = hint.get("image_version", "original")
            if case.registration_status == RegistrationStatus.NOT_PRESENT:
                evidence = case.evidence.setdefault("registration_address", FieldEvidence())
                if not manual(evidence):
                    evidence.status = FieldStatus.NOT_PRESENT
                    evidence.message = (
                        "На ID-карте нет регистрации. Введите адрес или добавьте документ с пропиской."
                    )

        if document_type == "kg_digital_id":
            from document_excel.infrastructure.document_profiles import KyrgyzDigitalIDProfile
            from document_excel.infrastructure.ocr_diagnostics import MESSAGES, OcrProblem

            for field in KyrgyzDigitalIDProfile().expected_fields:
                evidence = case.evidence.setdefault(field, FieldEvidence())
                if not getattr(case.data, field, None) and not evidence.candidates and not manual(evidence):
                    evidence.status = FieldStatus.NOT_RECOGNIZED
                    evidence.message = MESSAGES[OcrProblem.TEXT_DETECTED_BUT_FIELD_NOT_FOUND]
                hint = case.field_regions.get(field)
                if hint:
                    evidence.expected_image_id = hint["image_id"]
                    evidence.expected_bbox = hint["bbox"]
                    evidence.expected_image_version = hint.get("image_version", "original")

    @staticmethod
    def resolve_name_parts(case):

        for name in ("surname", "first_name", "patronymic"):
            evidence = case.evidence.get(name)
            if not evidence or manual(evidence):
                continue
            forms = {}
            for script in ("cyrillic", "latin"):
                values = [
                    c
                    for c in evidence.candidates
                    if c.validation_status != FieldStatus.INVALID
                    and name_script(c.normalized_value) == script
                ]
                unique = {c.normalized_value for c in values}
                forms[script] = next(iter(unique)) if len(unique) == 1 else None
                if script == "latin" and len(unique) > 1:
                    printed = {
                        c.normalized_value
                        for c in values
                        if c.source_method == "VISUAL_OCR"
                        and c.source_side == DocumentSide.FRONT
                        and c.confidence >= 0.9
                    }
                    if len(printed) == 1:
                        # TD1 name characters have no check digit. Keep the
                        # printed form visible while retaining the conflict.
                        forms[script] = next(iter(printed))
                field = name + "_" + script
                if field in type(case.data).model_fields and not manual(
                    case.evidence.get(field, FieldEvidence())
                ):
                    setattr(case.data, field, forms[script])
                    if values:
                        case.evidence[field] = FieldEvidence(
                            source="OCR: официальная форма",
                            status=FieldStatus.CONFLICT if len(unique) > 1 else FieldStatus.REVIEW,
                            candidates=[c.model_copy(update={"field_name": field}) for c in values],
                        )
            evidence.original_cyrillic = forms["cyrillic"]
            evidence.original_latin = forms["latin"]
            setattr(case.data, name, forms["cyrillic"] or forms["latin"])
            if forms["cyrillic"] and evidence.status == FieldStatus.CONFLICT:
                evidence.status = FieldStatus.REVIEW
                evidence.message = (
                    "Кириллица взята с печатного поля. Латинские варианты требуют отдельной сверки."
                )

    @staticmethod
    def compose_name(case):

        previous = case.evidence.get("full_name", FieldEvidence())
        if (
            case.data.full_name
            and not any(name in case.evidence for name in ("surname", "first_name", "patronymic"))
            and previous.source != "OCR: составлено из подписанных полей"
        ):
            return
        if (
            manual(previous)
            or previous.candidates
            and previous.source != "OCR: составлено из подписанных полей"
        ):
            return
        parts = [case.data.surname, case.data.first_name, case.data.patronymic]
        if case.data.document_type in {"kg_physical_id", "kg_digital_id"}:
            parts = [
                getattr(case.data, name)
                if manual(case.evidence.get(name, FieldEvidence()))
                else getattr(case.data, name + "_cyrillic")
                for name in ("surname", "first_name", "patronymic")
            ]
        components = [case.evidence.get(n, FieldEvidence()) for n in ("surname", "first_name", "patronymic")]
        if not parts[0] or not parts[1]:
            case.data.full_name = None
            case.evidence["full_name"] = FieldEvidence(
                source="OCR: составлено из подписанных полей",
                status=FieldStatus.NOT_RECOGNIZED,
                message="Не удалось распознать фамилию или имя; другие части не подставлялись.",
            )
            return
        case.data.full_name = " ".join(p for p in parts if p)
        case.evidence["full_name"] = FieldEvidence(
            source="OCR: составлено из подписанных полей",
            status=FieldStatus.REVIEW,
            candidates=[
                c.model_copy(
                    update={
                        "field_name": "full_name",
                        "raw_value": case.data.full_name,
                        "normalized_value": case.data.full_name,
                        "extraction_method": "composed-name",
                    }
                )
                for e in components
                for c in e.candidates
                if c.validation_status != FieldStatus.INVALID
                and c.normalized_value == getattr(case.data, c.field_name, None)
            ],
            message="Составлено в порядке: фамилия, имя, отчество. Сверьте каждую часть.",
            original_cyrillic=" ".join(e.original_cyrillic for e in components if e.original_cyrillic)
            or None,
            original_latin=" ".join(e.original_latin for e in components if e.original_latin) or None,
        )


NAME_FIELDS = {"full_name", "surname", "first_name", "patronymic"}


_CYRILLIC = r"А-Яа-яЁёҮүӨөҢң"


_LATIN = r"A-Za-z"


def name_script(value: str) -> str:
    cyrillic = bool(re.search(f"[{_CYRILLIC}]", value))
    latin = bool(re.search(f"[{_LATIN}]", value))
    return "cyrillic" if cyrillic and not latin else "latin" if latin and not cyrillic else "other"


def address_from_text(value: str) -> Address:
    # Only explicitly labelled fragments are split. Unstructured text remains intact in raw.
    region_match = re.search(r"([^,;]*(?:обл(?:асть|\.)|облусу))", value, re.IGNORECASE)
    street_match = re.search(r"((?:ул(?:ица|\.)|көчөсү|street)\s*[^,;]+)", value, re.IGNORECASE)
    return Address(
        raw=value,
        region=region_match.group(1).strip() if region_match else "",
        street=street_match.group(1).strip() if street_match else "",
    )


class RegistrationAddressResolver:
    def resolve(
        self, candidates: list[FieldCandidate], *, label_seen: bool = False, explicitly_absent: bool = False
    ) -> RegistrationStatus:
        if explicitly_absent:
            return RegistrationStatus.NOT_PRESENT
        addresses = [c for c in candidates if c.field_name == "registration_address"]
        if not addresses:
            return RegistrationStatus.NOT_RECOGNIZED if label_seen else RegistrationStatus.UNCERTAIN
        best = max(addresses, key=lambda c: c.confidence)
        if len({c.normalized_value for c in addresses}) > 1:
            return RegistrationStatus.UNCERTAIN
        parsed = address_from_text(best.normalized_value)
        if best.confidence < 0.75 or not parsed.region or not parsed.street:
            return RegistrationStatus.PARTIALLY_FOUND
        return RegistrationStatus.FOUND


def _apply_validated_candidates(
    case: ClientCase, candidates: list[FieldCandidate], document_type: str, label_seen: bool = False
) -> None:
    case.invalidate()
    # Older cases can carry an explicit manual address status without evidence.
    # Restore that provenance before re-resolving after a photo replacement.
    if (
        case.registration_status == RegistrationStatus.MANUALLY_ENTERED
        and case.data.registration_address.display()
        and "registration_address" not in case.evidence
    ):
        case.evidence["registration_address"] = FieldEvidence(
            source="ручной ввод", status=FieldStatus.MANUALLY_CORRECTED
        )
    if not case.document_type_manual:
        case.data.document_type = document_type if document_type != "unknown" else case.data.document_type
    grouped: dict[str, list[FieldCandidate]] = {}
    for candidate in candidates:
        grouped.setdefault(candidate.field_name, []).append(candidate)
    for image in case.images:
        sides = {
            c.source_side
            for c in candidates
            if c.source_image_id == image.id and c.source_side in {DocumentSide.FRONT, DocumentSide.BACK}
        }
        if image.role == DocumentSide.UNKNOWN and len(sides) == 1:
            image.role = next(iter(sides))
    for field, incoming in grouped.items():
        evidence = case.evidence.setdefault(field, FieldEvidence())
        existing = evidence.candidates
        # Keep alternatives from other photos and OCR passes, deduplicate exact evidence.
        unique = {(c.source_image_id, c.variant, c.normalized_value): c for c in existing + incoming}
        evidence.candidates = sorted(
            unique.values(),
            key=lambda c: (
                c.validation_status == FieldStatus.INVALID or bool(c.validation_errors),
                -c.confidence,
            ),
        )[:30]
        if field in NAME_FIELDS:
            cyrillic = [c for c in evidence.candidates if name_script(c.normalized_value) == "cyrillic"]
            latin = [c for c in evidence.candidates if name_script(c.normalized_value) == "latin"]
            evidence.original_cyrillic = cyrillic[0].normalized_value if cyrillic else None
            evidence.original_latin = latin[0].normalized_value if latin else None
        if _manual_evidence(evidence):
            evidence.message = "Есть новые OCR-кандидаты; ручное значение сохранено."
            continue
        best = evidence.candidates[0]
        if field in NAME_FIELDS:
            preferred = [c for c in evidence.candidates if name_script(c.normalized_value) == "cyrillic"]
            if preferred:
                best = preferred[0]
            by_script: dict[str, set[str]] = {}
            for candidate in evidence.candidates:
                if candidate.validation_status == FieldStatus.INVALID:
                    continue
                by_script.setdefault(name_script(candidate.normalized_value), set()).add(
                    candidate.normalized_value.casefold()
                )
            conflict = any(len(values) > 1 for values in by_script.values())
        else:
            conflict = (
                len(
                    {
                        _comparable(field, c.normalized_value)
                        for c in evidence.candidates
                        if c.validation_status != FieldStatus.INVALID
                    }
                )
                > 1
            )
        evidence.source = "OCR"
        evidence.status = FieldStatus.CONFLICT if conflict else best.validation_status
        evidence.message = (
            "Несколько вариантов. Выберите по изображению."
            if conflict
            else "Оценка эвристическая; проверьте документ."
        )
        if field == "passport_number" and not conflict:
            agreeing = {
                "MRZ" if c.source_method == "MRZ" else c.source_side.value
                for c in evidence.candidates
                if c.validation_status != FieldStatus.INVALID
            }
            if {"FRONT", "BACK", "MRZ"} <= agreeing:
                evidence.status = FieldStatus.VALID
                evidence.message = "Совпали лицевая сторона, оборот и проверенная MRZ. Сверьте документ."
        if field == "passport_number" and "+split-type" in best.extraction_method:
            evidence.message += " Буквенный префикс сохранён отдельно в поле типа / серии документа."
        elif field == "passport_number" and re.search(r"[OОIІ|l]", best.raw_value):
            evidence.message += " Проверьте похожие символы O/0 и I/1. Автоматической замены не было."
        if (
            field in {"passport_issued_by", "passport_department_code"}
            and "+split-department" in best.extraction_method
        ):
            evidence.message += " Цифры после МКК сохранены отдельно как код подразделения."
        if (
            field == "passport_number"
            and best.normalized_value.isdigit()
            and "+split-type" not in best.extraction_method
            and not case.data.passport_type
        ):
            evidence.message += " Сверьте буквы серии и цифры номера: OCR мог пропустить мелкие символы."
        if best.validation_status == FieldStatus.INVALID:
            continue
        if conflict and any(c.source_method == "MRZ" for c in evidence.candidates):
            evidence.message = "Расхождение между MRZ и печатными данными. Выберите значение по фотографии; автоматический выбор не выполнен."
            continue
        if field == "registration_address":
            if case.data.document_type == "kg_digital_id":
                from document_excel.domain.digital_address import digital_address_from_text

                case.data.registration_address = digital_address_from_text(best.normalized_value)
            else:
                case.data.registration_address = Address(raw=best.normalized_value)
        else:
            from document_excel.domain.models import DATE_FIELDS, parse_date

            setattr(
                case.data,
                field,
                parse_date(best.normalized_value) if field in DATE_FIELDS else best.normalized_value,
            )
    composed_name = case.evidence.get("full_name", FieldEvidence())
    if (
        (not case.data.full_name or composed_name.source == "OCR: составлено из подписанных полей")
        and case.data.surname
        and case.data.first_name
        and not _manual_evidence(composed_name)
    ):
        case.data.full_name = " ".join(
            v for v in [case.data.surname, case.data.first_name, case.data.patronymic] if v
        )
        name_evidence = [
            case.evidence.get(f, FieldEvidence()) for f in ("surname", "first_name", "patronymic")
        ]
        case.evidence["full_name"] = FieldEvidence(
            source="OCR: составлено из подписанных полей",
            status=FieldStatus.CONFLICT
            if any(e.status == FieldStatus.CONFLICT for e in name_evidence)
            else FieldStatus.REVIEW,
            original_cyrillic=" ".join(e.original_cyrillic for e in name_evidence if e.original_cyrillic)
            or None,
            original_latin=" ".join(e.original_latin for e in name_evidence if e.original_latin) or None,
        )
    if not _manual_evidence(case.evidence.get("registration_address", FieldEvidence())):
        all_candidates = case.evidence.get("registration_address", FieldEvidence()).candidates
        case.registration_status = RegistrationAddressResolver().resolve(
            all_candidates,
            label_seen=label_seen,
            explicitly_absent=case.data.document_type == "kg_physical_id"
            and {DocumentSide.FRONT, DocumentSide.BACK} <= {i.role for i in case.images}
            and not all_candidates,
        )
    # Official scripts are independent values, never an invented transliteration.
    for name in ("surname", "first_name", "patronymic"):
        evidence = case.evidence.get(name, FieldEvidence())
        for script in ("cyrillic", "latin"):
            field_name = f"{name}_{script}"
            if field_name not in type(case.data).model_fields:
                continue
            if not _manual_evidence(case.evidence.get(field_name, FieldEvidence())):
                setattr(case.data, field_name, getattr(evidence, "original_" + script))
    for report in case.mrz_reports.values():
        for field in report.conflicts:
            evidence = case.evidence.setdefault(field, FieldEvidence())
            if not _manual_evidence(evidence):
                evidence.status = FieldStatus.CONFLICT
                evidence.message = (
                    "Расхождение печатной даты и MRZ. Полный год не подставлялся; сверьте источники."
                    if field in {"birth_date", "expiration_date"}
                    else "Расхождение печатных данных и MRZ. Сверьте оба источника."
                )
        if report.valid:
            for field, evidence in case.evidence.items():
                if (
                    evidence.status == FieldStatus.CONFLICT
                    and any(
                        c.source_method == "MRZ" and c.source_image_id == report.image_id
                        for c in evidence.candidates
                    )
                    and field not in report.conflicts
                ):
                    report.conflicts.append(field)


def _comparable(field: str, value: str) -> str:
    compact = re.sub(r"\s+", " ", value).strip().casefold()
    if field == "gender":
        token = re.sub(r"[\s/]+", "", compact)
        if token in {"м", "m", "эм", "эркек", "male"}:
            return "male"
        if token in {"ж", "f", "аял", "аf", "female"}:
            return "female"
    if field == "citizenship" and compact in {
        "kgz",
        "кыргыз республикасы",
        "кыргызская республика",
        "кыргызстан",
        "kyrgyz republic",
    }:
        return "kgz"
    return compact


def _manual_evidence(evidence: FieldEvidence) -> bool:
    return evidence.status == FieldStatus.MANUALLY_CORRECTED or any(
        marker in evidence.source.casefold() for marker in ("ручн", "пользоват")
    )
