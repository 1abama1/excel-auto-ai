from __future__ import annotations

import hashlib
import logging
import threading
import time
from collections.abc import Callable
from dataclasses import dataclass, replace
from dataclasses import field as dataclass_field
from pathlib import Path

from document_excel.domain.errors import AppError, Cancelled
from document_excel.domain.models import (
    ClientCase,
    DocumentSide,
    DocumentType,
    FieldCandidate,
    FieldStatus,
    MRZReport,
    SourceImage,
    new_id,
    validate_field,
)
from document_excel.infrastructure.extraction import (
    DocumentProfile,
    DocumentSideClassifier,
    FieldExtractor,
    classify,
    load_document_profiles,
)
from document_excel.infrastructure.identifier_ocr import refine_identifier
from document_excel.infrastructure.ocr_contract import OcrEngine, OcrResult
from document_excel.infrastructure.ocr_factory import create_engine
from document_excel.infrastructure.preprocessing import ImagePreprocessingService, PreparedImage
from document_excel.infrastructure.storage import Settings, collect_images


@dataclass
class RecognitionOutput:
    candidates: list[FieldCandidate]
    document_type: str
    warnings: list[str]
    label_seen: bool
    elapsed: float
    sides: dict[str, DocumentSide] = dataclass_field(default_factory=dict)
    # Deliberately in memory only. History/logging must never serialize OCR text.
    ocr_texts: dict[str, str] = dataclass_field(default_factory=dict)
    mrz_reports: dict[str, MRZReport] = dataclass_field(default_factory=dict)
    run_id: str = ""
    diagnostic_blocks: dict[str, list[dict]] = dataclass_field(default_factory=dict)
    field_regions: dict[str, dict] = dataclass_field(default_factory=dict)
    processing_metrics: dict[str, dict] = dataclass_field(default_factory=dict)


_SIDE_FIELDS = {
    DocumentSide.FRONT: {"surname", "first_name", "patronymic", "birth_date", "gender", "citizenship"},
    DocumentSide.BACK: {
        "personal_identifier",
        "passport_number",
        "passport_issued_by",
        "passport_issued_date",
        "expiration_date",
        "registration_address",
    },
}


@dataclass
class _RecognitionAttempt:
    result: OcrResult
    candidates: list[FieldCandidate]
    kind: str
    profile: DocumentProfile | None
    side: DocumentSide

    def quality(self) -> tuple[int, float, int, float]:
        """Rank usable distinct fields, coverage, invalid fields and confidence.

        Long OCR output and repeated bilingual labels do not improve the rank.
        A better segmentation replaces the earlier pass rather than mixing
        values from incompatible OCR attempts into artificial conflicts.
        """
        valid: dict[str, float] = {}
        invalid: set[str] = set()
        for candidate in self.candidates:
            if candidate.validation_status == FieldStatus.INVALID or validate_field(
                candidate.field_name, candidate.normalized_value
            ):
                invalid.add(candidate.field_name)
                continue
            valid[candidate.field_name] = max(valid.get(candidate.field_name, 0), candidate.confidence)
        expected = set(self.profile.expected_fields) if self.profile else set()
        if self.side in _SIDE_FIELDS:
            expected = expected & _SIDE_FIELDS[self.side] or _SIDE_FIELDS[self.side]
        coverage = len(set(valid) & expected) / len(expected) if expected else min(1, len(valid) / 4)
        confidence = sum(valid.values()) / len(valid) if valid else 0.0
        return len(valid), coverage, -len(invalid - valid.keys()), confidence

    def needs_alternative(self, roi: bool) -> bool:
        count, coverage, _, confidence = self.quality()
        if roi:
            return count < 1 or confidence < 0.7
        return count < 3 or coverage < 0.55 or confidence < 0.7


def merge_regional_candidates(global_candidates, regional):
    """A Latin crop must not erase the independently read Cyrillic name.

    Keep name evidence from both optical sources so a genuine disagreement
    becomes a conflict. Typed fields retain the existing region precedence.
    """
    replaced = {c.field_name for c in regional} - {"surname", "first_name", "patronymic", "full_name"}
    combined = [c for c in global_candidates if c.field_name not in replaced] + regional
    unique = {}
    for candidate in combined:
        unique[candidate.model_dump_json()] = candidate
    return list(unique.values())


class RecognitionService:
    def __init__(self, settings: Settings, engine: OcrEngine | None = None) -> None:
        self.settings = settings
        self.engine = engine or create_engine(settings)
        self.preprocessor = ImagePreprocessingService()

    def run(
        self,
        images: list[SourceImage],
        cancel: threading.Event,
        progress: Callable[[int, str], None],
        roi: tuple[int, int, int, int] | None = None,
        document_type: str | None = None,
    ) -> RecognitionOutput:
        if not images:
            raise AppError("Добавьте хотя бы одну фотографию.")
        if getattr(self.engine, "preferred_input", "") == "color" and roi is None:
            from document_excel.application.fast_recognition import recognize

            return recognize(self, images, cancel, progress, document_type)
        start = time.monotonic()
        candidates: list[FieldCandidate] = []
        notes: list[str] = []
        types: list[str] = []
        sides: dict[str, DocumentSide] = {}
        ocr_texts: dict[str, str] = {}
        mrz_reports: dict[str, MRZReport] = {}
        diagnostic_blocks: dict[str, list[dict]] = {}
        field_regions: dict[str, dict] = {}
        physical_seen: set[str] = set()
        label_seen = False
        success = 0
        profiles = load_document_profiles()
        expected_fields: set[str] = set()
        for index, source in enumerate(images):
            if cancel.is_set():
                raise Cancelled()

            def stage(fraction: float, message: str, image_index: int = index) -> None:
                # Percentages describe completed image/stage units only; OCR
                # itself has no measurable percentage and the UI may pulse.
                progress(
                    min(99, int((image_index + fraction) * 100 / len(images))),
                    f"Фото {image_index + 1} из {len(images)} · {message}",
                )

            stage(0, "Подготовка изображения")
            try:
                prepared = self.preprocessor.prepare(
                    source.selected_path,
                    "original" if roi else self.settings.preprocessing,
                    source.selected_rotation,
                )
                if cancel.is_set():
                    raise Cancelled()
                notes.extend(prepared.quality.warnings)
                stage(
                    0.15,
                    "Поиск документа · "
                    + (
                        "область найдена"
                        if prepared.document_region and prepared.document_region.detected
                        else "используется оригинал"
                    ),
                )
                geometry_warnings = (
                    set(prepared.document_region.warnings) if prepared.document_region else set()
                )
                low_quality = bool(set(prepared.quality.warnings) - geometry_warnings)
                # Prefer the final processed variant. One primary call and at
                # most one sparse-text alternative keep the work bounded.
                inputs = prepared.ocr_variants
                variant = next(reversed(inputs))
                color_input = getattr(self.engine, "preferred_input", "") == "color"
                if color_input and self.settings.preprocessing == "auto":
                    variant = "document" if "document" in inputs else "original"
                    inputs = prepared.variants
                array = inputs[variant]
                roi_offset = (0, 0)
                if roi:
                    x, y, width, height = roi
                    left, top = max(0, x), max(0, y)
                    right, bottom = (
                        min(prepared.original.shape[1], x + width),
                        min(prepared.original.shape[0], y + height),
                    )
                    array = prepared.original[top:bottom, left:right]
                    variant = "roi"
                    roi_offset = (left, top)
                    if width <= 0 or height <= 0 or right <= left or bottom <= top or array.size == 0:
                        raise AppError("Выделите область внутри изображения.")

                def extract_attempt(
                    result: OcrResult,
                    prepared_image: PreparedImage = prepared,
                    source_image: SourceImage = source,
                    quality_warning: bool = low_quality,
                    selected_variant: str = variant,
                    offset: tuple[int, int] = roi_offset,
                ) -> _RecognitionAttempt:
                    if cancel.is_set():
                        raise Cancelled()
                    kind = document_type or classify(result, profiles)
                    profile = next((p for p in profiles if p.id == kind), None)
                    detected_side = DocumentSideClassifier().classify(result, kind)
                    side = detected_side if detected_side != DocumentSide.UNKNOWN else source_image.role
                    if source_image.role != DocumentSide.UNKNOWN and (
                        source_image.role == DocumentSide.ADDITIONAL or kind != DocumentType.KG_PHYSICAL_ID
                    ):
                        side = source_image.role
                    # The extractor omits unsafe ROI coordinates for older
                    # callers. Here a known axis crop can be mapped explicitly.
                    extraction_result = replace(result, variant="original") if roi else result
                    extracted = FieldExtractor(profile).extract(extraction_result, quality_warning)
                    for candidate in extracted:
                        candidate.image_version = (
                            "processed"
                            if source_image.processed_confirmed and source_image.processed_path
                            else "original"
                        )
                        candidate.source_side = side
                        candidate.variant = selected_variant
                        if candidate.source_bbox is not None:
                            if roi:
                                bx, by, bw, bh = candidate.source_bbox
                                candidate.source_bbox = (bx + offset[0], by + offset[1], bw, bh)
                            else:
                                candidate.source_bbox = prepared_image.to_original_bbox(
                                    selected_variant, candidate.source_bbox
                                )
                    return _RecognitionAttempt(result, extracted, kind, profile, side)

                if cancel.is_set():
                    raise Cancelled()
                stage(0.25, "Распознавание текста")
                result = self.engine.recognize(array, source.id, variant, cancel)
                stage(0.65, "Извлечение данных")
                selected = extract_attempt(result)
                identifier_evidence = list(selected.candidates)
                alternate_method = getattr(self.engine, "recognize_with_psm", None)
                if (
                    selected.needs_alternative(bool(roi))
                    and self.settings.psm != 11
                    and not color_input
                    and callable(alternate_method)
                ):
                    if cancel.is_set():
                        raise Cancelled()
                    stage(0.7, "Распознавание текста · дополнительная проверка")
                    try:
                        alternate_result = alternate_method(array, source.id, variant, cancel, 11)
                        alternate = extract_attempt(alternate_result)
                        identifier_evidence.extend(alternate.candidates)
                        if alternate.quality() > selected.quality():
                            selected = alternate
                    except Cancelled:
                        raise
                    except AppError:
                        # A failed optional pass must not discard a usable
                        # primary result or expose raw OCR data in diagnostics.
                        notes.append(f"Фото {index + 1}: дополнительная проверка OCR не завершена.")
                if cancel.is_set():
                    raise Cancelled()
                if color_input and not roi and selected.needs_alternative(False):
                    alternative_variant = next(reversed(prepared.ocr_variants))
                    if alternative_variant != variant:
                        stage(0.7, "Проверка варианта подготовки")
                        try:
                            alternative_result = self.engine.recognize(
                                prepared.variants[alternative_variant], source.id, alternative_variant, cancel
                            )
                            alternative = extract_attempt(
                                alternative_result, selected_variant=alternative_variant
                            )
                            identifier_evidence.extend(alternative.candidates)
                            if alternative.quality() > selected.quality():
                                selected = alternative
                        except Cancelled:
                            raise
                        except AppError:
                            notes.append(f"Фото {index + 1}: дополнительная проверка OCR не завершена.")
                identifier_method = getattr(self.engine, "recognize_identifier_line", None)
                if callable(identifier_method):
                    stage(0.85, "Уточнение типа и номера документа")
                    try:
                        selected.candidates = refine_identifier(
                            array, selected.result, selected.candidates, identifier_method, cancel
                        )
                    except Cancelled:
                        raise
                    except AppError:
                        notes.append(
                            f"Фото {index + 1}: уточнение номера OCR не завершено; сверьте тип и номер."
                        )
                stage(0.9, "Проверка")
                if (
                    not roi
                    and callable(getattr(self.engine, "recognize_mrz", None))
                    and (
                        selected.kind == DocumentType.KG_PHYSICAL_ID
                        or DocumentType.KG_PHYSICAL_ID in types
                        or document_type == DocumentType.KG_PHYSICAL_ID
                        or selected.quality()[0] < 3
                        and prepared.document_region is not None
                        and prepared.document_region.detected
                        and 1.25
                        < max(prepared.document_region.bbox[2:])
                        / max(1, min(prepared.document_region.bbox[2:]))
                        < 2.05
                    )
                ):
                    from document_excel.infrastructure.document_profiles import KyrgyzPhysicalIDProfile
                    from document_excel.infrastructure.mrz import MRZRecognizer
                    from document_excel.infrastructure.physical_id import (
                        layout_candidates,
                        normalize_physical_name_scripts,
                        prepare_plastic_card,
                        prepare_region,
                        regional_candidates,
                        rotate_plastic_card,
                    )

                    stage(0.9, "Пластиковая карта · области и MRZ")
                    try:
                        card = prepare_plastic_card(prepared.original)
                        card_input = card.image if color_input else prepare_region(card.image, "text")
                        card_result = self.engine.recognize(card_input, source.id, "physical-card", cancel)
                        kind = classify(card_result, profiles)
                        if kind != DocumentType.KG_PHYSICAL_ID:
                            opposite = rotate_plastic_card(card, 2)
                            opposite_input = (
                                opposite.image if color_input else prepare_region(opposite.image, "text")
                            )
                            opposite_result = self.engine.recognize(
                                opposite_input, source.id, "physical-card", cancel
                            )
                            if classify(opposite_result, profiles) == DocumentType.KG_PHYSICAL_ID:
                                card, card_result, kind = (
                                    opposite,
                                    opposite_result,
                                    DocumentType.KG_PHYSICAL_ID,
                                )
                        # A manually selected plastic profile, another side of the same
                        # case, or a fresh structural match is required before field use.
                        if (
                            kind == DocumentType.KG_PHYSICAL_ID
                            or selected.kind == DocumentType.KG_PHYSICAL_ID
                            or DocumentType.KG_PHYSICAL_ID in types
                        ):
                            side = DocumentSideClassifier().classify(card_result, DocumentType.KG_PHYSICAL_ID)
                            if side == DocumentSide.UNKNOWN:
                                side = selected.side
                            if source.role == DocumentSide.ADDITIONAL:
                                side = DocumentSide.ADDITIONAL
                            physical = FieldExtractor(KyrgyzPhysicalIDProfile()).extract(
                                card_result, low_quality
                            )
                            physical = normalize_physical_name_scripts(
                                physical, card.image.shape[1], card.image.shape[0], side
                            )
                            physical.extend(regional_candidates(card, card_result, self.engine, cancel))
                            # Keep literal numbers from the independent global
                            # pass; regional OCR must not erase correct digits.
                            from document_excel.infrastructure.region_ocr import DocumentNumberExtractor

                            global_numbers = [
                                c
                                for c in identifier_evidence
                                if c.field_name in {"passport_number", "passport_type"}
                                and (
                                    DocumentNumberExtractor.literal(c.raw_value)
                                    or c.field_name == "passport_type"
                                    and c.normalized_value == "ID"
                                )
                            ]
                            regional = layout_candidates(card, card_result, self.engine, cancel, side)
                            for key, area in KyrgyzPhysicalIDProfile().ocr_regions.items():
                                prefix, _, name = key.partition(".")
                                if prefix == side.value.lower():
                                    rx, ry, rw, rh = area
                                    ch, cw = card.image.shape[:2]
                                    field_regions[name] = {
                                        "image_id": source.id,
                                        "bbox": card.original_bbox(
                                            (round(rx * cw), round(ry * ch), round(rw * cw), round(rh * ch))
                                        ),
                                        "image_version": "processed"
                                        if source.processed_confirmed
                                        else "original",
                                    }
                            # A validated value from a dedicated region supersedes noisy
                            # label fragments on that same image; independent images/MRZ
                            # remain separate sources and can still produce conflicts.
                            regional_names = {c.field_name for c in regional}
                            physical = merge_regional_candidates(physical, regional)
                            for candidate in physical:
                                candidate.source_bbox = card.original_bbox(candidate.source_bbox)
                                candidate.image_version = (
                                    "processed" if source.processed_confirmed else "original"
                                )
                                candidate.source_side = side
                                if not candidate.variant.startswith("official-"):
                                    candidate.variant = "physical-card"
                            # A coherent region pass replaces a weaker global pass.
                            attempt = _RecognitionAttempt(
                                card_result,
                                physical,
                                DocumentType.KG_PHYSICAL_ID,
                                KyrgyzPhysicalIDProfile(),
                                side,
                            )
                            if attempt.quality() >= selected.quality():
                                selected = attempt
                            if regional_names:
                                selected.candidates = merge_regional_candidates(selected.candidates, regional)
                            selected.candidates.extend(global_numbers)
                            selected.kind = DocumentType.KG_PHYSICAL_ID
                            physical_seen.add(source.id)
                            if side != DocumentSide.UNKNOWN:
                                selected.side = side
                            if side not in {DocumentSide.FRONT, DocumentSide.ADDITIONAL}:
                                report = MRZRecognizer(self.engine).recognize(card.image, source.id, cancel)
                                report.source_bbox = card.original_bbox(report.source_bbox)
                                report.image_version = (
                                    "processed" if source.processed_confirmed else "original"
                                )
                                mrz_reports[source.id] = report
                                if report.lines:
                                    selected.side = DocumentSide.BACK
                                if not report.valid:
                                    notes.append("MRZ не прошла проверку. Сверьте оборот или повторите OCR.")
                            if not card.detected:
                                notes.append(
                                    "Проверьте границы пластиковой карты; доступно ручное кадрирование."
                                )
                    except Cancelled:
                        raise
                    except AppError:
                        notes.append(
                            "MRZ: \u0434\u043e\u043f\u043e\u043b\u043d\u0438\u0442\u0435\u043b\u044c\u043d\u0430\u044f \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0430 \u043d\u0435 \u0437\u0430\u0432\u0435\u0440\u0448\u0435\u043d\u0430; \u0441\u0432\u0435\u0440\u044c\u0442\u0435 \u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442."
                        )
                if not selected.result.blocks:
                    continue
                success += 1
                for candidate in selected.candidates:
                    candidate.engine_name = selected.result.engine_name
                    candidate.model_version = selected.result.model_version
                candidates.extend(selected.candidates)
                types.append(selected.kind)
                sides[source.id] = selected.side
                ocr_texts[source.id] = selected.result.text
                diagnostic_blocks[source.id] = [
                    {
                        "bbox": b.bbox,
                        "confidence": round(b.confidence, 3),
                        "characters": len(b.text),
                        "variant": selected.result.variant,
                        "engine": selected.result.engine_name,
                        "model": selected.result.model_version,
                        "script": b.script,
                    }
                    for b in selected.result.blocks
                ]
                if selected.profile:
                    expected_fields.update(selected.profile.expected_fields)
                label_seen |= any(
                    word in selected.result.text.upper()
                    for word in ["РЕГИСТРАЦ", "ПРОПИСК", "REGISTRATION", "КАТТАЛГАН", "КАТТОО ДАРЕГИ"]
                )
            except Cancelled:
                raise
            except AppError as exc:
                notes.append(f"Фото {index + 1}: {exc}")
                logging.getLogger("document_excel").warning("image_processing_failed")
        if cancel.is_set():
            raise Cancelled()
        if not success:
            raise AppError(
                "Текст не распознан. "
                + " ".join(dict.fromkeys(notes))
                + " Проверьте OCR и качество фотографии."
            )
        known = {kind for kind in types if kind != "unknown"}
        # A reverse-side MRZ can establish the profile after the other side was
        # already processed. Revisit only those images, once, with that policy.
        pending = [image for image in images if image.id not in physical_seen]
        if (
            not roi
            and not document_type
            and DocumentType.KG_PHYSICAL_ID in known
            and pending
            and callable(getattr(self.engine, "recognize_mrz", None))
        ):
            additional = self.run(pending, cancel, progress, document_type=DocumentType.KG_PHYSICAL_ID)
            pending_ids = {image.id for image in pending}
            candidates = [c for c in candidates if c.source_image_id not in pending_ids]
            candidates.extend(c for c in additional.candidates if c.source_method != "MRZ")
            sides.update(additional.sides)
            ocr_texts.update(additional.ocr_texts)
            mrz_reports.update(additional.mrz_reports)
            diagnostic_blocks.update(additional.diagnostic_blocks)
            field_regions.update(additional.field_regions)
            label_seen |= additional.label_seen
        from document_excel.infrastructure.mrz import mrz_candidates

        if DocumentType.KG_PHYSICAL_ID in known:
            from document_excel.infrastructure.region_ocr import DocumentNumberExtractor

            candidates = [
                c
                for c in candidates
                if (
                    c.field_name not in {"passport_number", "passport_type"}
                    or c.field_name == "passport_number"
                    and DocumentNumberExtractor.literal(c.raw_value)
                    or c.field_name == "passport_type"
                    and c.normalized_value == "ID"
                )
            ]

        for report in mrz_reports.values():
            report.warnings = [warning for warning in report.warnings if "столетие MRZ" not in warning]
            report.conflicts.clear()
            candidates.extend(mrz_candidates(report, candidates))
        missing = expected_fields - {candidate.field_name for candidate in candidates}
        if missing:
            from document_excel.domain.models import FIELD_LABELS

            notes.append(
                "Не найдены ожидаемые поля: "
                + ", ".join(FIELD_LABELS.get(name, name) for name in sorted(missing))
            )
        document_type = (
            DocumentType.KG_PHYSICAL_ID
            if DocumentType.KG_PHYSICAL_ID in known
            else DocumentType.KG_DIGITAL_ID
            if DocumentType.KG_DIGITAL_ID in known
            else "kg_id"
            if "kg_id" in known
            else next(iter(known))
            if len(known) == 1
            else "unknown"
        )
        progress(100, "Распознавание завершено. Сверьте данные.")
        return RecognitionOutput(
            candidates,
            document_type,
            list(dict.fromkeys(notes)),
            label_seen,
            time.monotonic() - start,
            sides,
            ocr_texts,
            mrz_reports,
            diagnostic_blocks=diagnostic_blocks,
            field_regions=field_regions,
        )


class CaseService:
    def add_images(self, case: ClientCase, paths: list[Path]) -> list[str]:
        images, errors = collect_images(paths)
        existing = {image.path.resolve() for image in case.images}
        for path in images:
            if path not in existing:
                if len(case.images) >= 100:
                    errors.append("Не более 100 фото на клиента. Создайте следующую запись.")
                    break
                try:
                    with path.open("rb") as stream:
                        digest = hashlib.file_digest(stream, "sha256").hexdigest()
                except OSError:
                    errors.append("Не удалось прочитать фотографию. Добавьте её заново.")
                    continue
                case.images.append(SourceImage(path=path, photo_digest=digest, case_id=case.id))
                existing.add(path)
        case.invalidate()
        return errors

    def merge(self, target: ClientCase, source: ClientCase) -> None:
        existing = {image.path.resolve() for image in target.images}
        target.images.extend(image for image in source.images if image.path.resolve() not in existing)
        for image in target.images:
            image.case_id = target.id
        target.invalidate()
        # Data are not silently combined across people. OCR re-run preserves target's manual edits.
        target.errors.append("Группы объединены. Проверьте принадлежность фото и повторите OCR.")

    def move_image(self, source: ClientCase, target: ClientCase, image: SourceImage) -> None:
        source.images = [item for item in source.images if item.id != image.id]
        if not any(item.path.resolve() == image.path.resolve() for item in target.images):
            target.images.append(image.model_copy(update={"id": new_id(), "case_id": target.id}))
        self.remove_evidence(source, image.id)
        source.invalidate()
        target.invalidate()

    @staticmethod
    def remove_evidence(case: ClientCase, image_id: str) -> None:
        from document_excel.domain.models import ClientData
        from document_excel.domain.resolution import manual
        from document_excel.infrastructure.extraction import apply_candidates

        case.mrz_reports.pop(image_id, None)
        case.side_proposals.pop(image_id, None)
        case.diagnostic_blocks.pop(image_id, None)
        case.processing_metrics.pop(image_id, None)
        case.field_regions = {
            name: hint for name, hint in case.field_regions.items() if hint.get("image_id") != image_id
        }
        case.pending_identity = None
        case.identity_approved = ""
        case.identity_status = "not_checked"
        # Clear OCR-derived fields referencing a removed/rotated photo, never manual entries.
        for field, evidence in list(case.evidence.items()):
            if any(c.source_image_id == image_id for c in evidence.candidates):
                evidence.candidates = [c for c in evidence.candidates if c.source_image_id != image_id]
                if not manual(evidence):
                    if field in type(case.data).model_fields:
                        setattr(case.data, field, getattr(ClientData(), field))
                    if not evidence.candidates:
                        del case.evidence[field]
        composed_name = case.evidence.get("full_name")
        for name in ("surname", "first_name", "patronymic"):
            if name not in case.evidence:
                for script in ("cyrillic", "latin"):
                    target = f"{name}_{script}"
                    if target in type(case.data).model_fields and not case.evidence.get(target):
                        setattr(case.data, target, None)
        if composed_name and composed_name.source == "OCR: составлено из подписанных полей":
            case.data.full_name = None
            del case.evidence["full_name"]
        case.invalidate()
        remaining = [c for e in case.evidence.values() for c in e.candidates]
        apply_candidates(case, remaining, case.data.document_type)
