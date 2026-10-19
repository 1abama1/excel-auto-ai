"""One neural document pass per photograph, followed by targeted recovery.

Cache entries contain personal data in RAM only and are isolated by case/image
identity, bytes, rotation, versions, runtime settings and profile contents.
"""

from __future__ import annotations

import copy
import hashlib
import threading
import time
from collections import OrderedDict
from typing import TYPE_CHECKING

import cv2

if TYPE_CHECKING:
    from document_excel.application.workflows import RecognitionOutput

from document_excel.domain.errors import AppError, Cancelled, OcrFailure
from document_excel.domain.models import DocumentSide, DocumentType, FieldStatus, validate_field
from document_excel.infrastructure.digital_recovery import recover_digital_fields
from document_excel.infrastructure.digital_region import DigitalIDRegion, DigitalIDRegionDetector
from document_excel.infrastructure.document_profiles import KyrgyzPhysicalIDProfile
from document_excel.infrastructure.extraction import (
    DocumentSideClassifier,
    FieldExtractor,
    classify,
    load_document_profiles,
)
from document_excel.infrastructure.full_name import NAMES, FullNameExtractor
from document_excel.infrastructure.identifier_ocr import refine_identifier, refine_mkk_authority
from document_excel.infrastructure.mrz import MRZRecognizer, mrz_candidates
from document_excel.infrastructure.physical_id import (
    layout_candidates,
    prepare_plastic_card,
    rotate_plastic_card,
)
from document_excel.infrastructure.preprocessing import assess_quality

PREPROCESSING_VERSION = "digital-ppocrv5-fields-7"
_CACHE: OrderedDict[tuple, RecognitionOutput] = OrderedDict()
_LOCK = threading.Lock()


def add_inference_metrics(metric, engine, before):
    for name, value in getattr(engine, "request_totals", {}).items():
        metric[name] = metric.get(name, 0) + value - before.get(name, 0)


def clear_cache():
    with _LOCK:
        _CACHE.clear()


def document_pixels(image, mode):
    """Photometric modes retain the card's coordinate transform and full size."""
    if mode in {"auto", "original", "perspective", "deskewed"}:
        # The card homography already corrects its perspective and baseline;
        # applying a second untracked geometric transform would spoil evidence.
        return image
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    if mode == "grayscale":
        return gray
    if mode == "adaptive":
        return cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 35, 12)
    if mode == "denoised":
        return cv2.bilateralFilter(gray, 7, 50, 50)
    if mode == "sharpened":
        return cv2.addWeighted(gray, 1.6, cv2.GaussianBlur(gray, (0, 0), 2), -0.6, 0)
    if mode == "contrast":
        return cv2.createCLAHE(clipLimit=2, tileGridSize=(8, 8)).apply(gray)
    return image


def cache_key(service, images, forced, profiles):
    identity = []
    for source in images:
        try:
            digest = hashlib.sha256(source.selected_path.read_bytes()).digest()
        except OSError:
            digest = None
        identity.append(
            (
                source.case_id,
                source.id,
                digest,
                source.selected_rotation,
                source.processed_confirmed,
                str(source.role),
            )
        )
    model_manifest = getattr(service.engine, "root", None)
    model_version = None
    if model_manifest is not None:
        model_version = (
            str(model_manifest),
            tuple(
                (str(p.relative_to(model_manifest)), p.stat().st_size, p.stat().st_mtime_ns)
                for p in sorted(model_manifest.rglob("*"))
                if p.is_file()
            ),
        )
    return (
        PREPROCESSING_VERSION,
        tuple(identity),
        service.settings.model_dump_json(),
        model_version,
        repr(profiles),
        forced,
    )


def recognize(service, images, cancel, progress, forced=None):
    from document_excel.application.workflows import RecognitionOutput

    started = time.perf_counter()
    profiles = load_document_profiles()
    if cancel.is_set():
        raise Cancelled()
    key = cache_key(service, images, forced, profiles)
    with _LOCK:
        cached = _CACHE.get(key)
        if cached is not None:
            _CACHE.move_to_end(key)
            output = copy.deepcopy(cached)
            if cancel.is_set():
                raise Cancelled()
            output.elapsed = time.perf_counter() - started
            for metric in output.processing_metrics.values():
                metric["result_cache_hit"] = True
                for field in (
                    "ocr_requests",
                    "request_seconds",
                    "text_detection_seconds",
                    "text_recognition_seconds",
                ):
                    metric[field] = 0
            progress(100, "Данные из памяти · сверьте с документом")
            return output
    records = []
    processing_metrics = {}
    warnings = []
    cacheable = True
    # Read both sides before deciding the document profile. A reverse side can
    # establish the profile without forcing a second FRONT inference pass.
    for index, source in enumerate(images):
        if cancel.is_set():
            raise Cancelled()
        progress(int(index * 65 / len(images)), f"Фото {index + 1} из {len(images)} · подготовка")
        metric = {"image_id": source.id, "status": "RUNNING", "stages": {}}
        processing_metrics[source.id] = metric
        try:
            service.engine.cache_context = (source.case_id, source.id, PREPROCESSING_VERSION, repr(profiles))
            phase_started = time.perf_counter()
            prepared = service.preprocessor.prepare(
                source.selected_path, "original", source.selected_rotation
            )
            metric["stages"]["load_preprocessing_seconds"] = time.perf_counter() - phase_started
            metric["original_size"] = (
                getattr(prepared, "source_size", None) or prepared.original.shape[:2][::-1]
            )
            metric["working_original_size"] = prepared.original.shape[:2][::-1]
            phase_started = time.perf_counter()
            card = None
            if forced != DocumentType.KG_PHYSICAL_ID:
                card = DigitalIDRegionDetector().prepare(
                    prepared.original,
                    forced=forced == DocumentType.KG_DIGITAL_ID,
                    detection=getattr(prepared, "document_region", None),
                )
            if card is None:
                card = prepare_plastic_card(prepared.original)
            metric["stages"]["region_detection_seconds"] = time.perf_counter() - phase_started
            quality = prepared.quality
            if isinstance(card, DigitalIDRegion):
                metric.update(card.metrics())
                x, y, rw, rh = card.bbox
                quality = assess_quality(prepared.original[y : y + rh, x : x + rw])
            else:
                metric.update({"inference_size": card.image.shape[:2][::-1], "roi_detected": card.detected})
            metric["quality"] = {
                name: getattr(quality, name, None) for name in ("blur", "contrast", "brightness")
            }
            progress(
                int((index + 0.2) * 65 / len(images)),
                f"Фото {index + 1} из {len(images)} · обнаружение и чтение текста",
            )
            inference_before = dict(getattr(service.engine, "request_totals", {}))
            result = service.engine.recognize(
                document_pixels(card.image, service.settings.preprocessing),
                source.id,
                "digital-card" if isinstance(card, DigitalIDRegion) else "physical-card",
                cancel,
            )
            if isinstance(card, DigitalIDRegion) and card.detected and len(result.blocks) < 3:
                # A confident boundary can still enclose a non-document panel.
                # Retry the untouched screenshot once; never discard useful
                # primary blocks in favour of an empty fallback.
                oh, ow = prepared.original.shape[:2]
                alternate = service.engine.recognize(
                    prepared.original, source.id, "digital-original-fallback", cancel
                )
                metric["original_fallback_attempted"] = True
                if len(alternate.blocks) > len(result.blocks):
                    card = DigitalIDRegion(prepared.original, (ow, oh), (0, 0, ow, oh), False, 0, True)
                    result = alternate
                    metric.update(card.metrics())
            metric.update(
                {
                    "engine": result.engine_name,
                    "model": result.model_version,
                    "text_blocks": len(result.blocks),
                    "valid_boxes": sum(b.bbox[2] > 0 and b.bbox[3] > 0 for b in result.blocks),
                }
            )
            heights = sorted(b.bbox[3] for b in result.blocks)
            scale = metric.get("scale_y", 1.0)
            metric["median_text_height_source"] = heights[len(heights) // 2] / scale if heights else 0
            phase_started = time.perf_counter()
            kind = forced or classify(
                result, profiles, digital_screen=isinstance(card, DigitalIDRegion) and card.mobile_context
            )
            side = DocumentSideClassifier().classify(result, kind)
            metric["stages"]["classification_seconds"] = time.perf_counter() - phase_started
            # A genuinely unreadable orientation is the only automatic second
            # document pass. BACK with readable labels/MRZ is not rotated.
            confidence = sum(b.confidence for b in result.blocks) / max(1, len(result.blocks))
            if (
                not isinstance(card, DigitalIDRegion)
                and side == DocumentSide.UNKNOWN
                and (len(result.blocks) < 5 or confidence < 0.65)
            ):
                opposite = rotate_plastic_card(card, 2)
                alternate = service.engine.recognize(
                    document_pixels(opposite.image, service.settings.preprocessing),
                    source.id,
                    "orientation-180",
                    cancel,
                )
                alternate_kind = forced or classify(alternate, profiles)
                alternate_side = DocumentSideClassifier().classify(alternate, alternate_kind)
                if alternate_side != DocumentSide.UNKNOWN or len(alternate.blocks) > len(result.blocks):
                    card, result = opposite, alternate
                    kind = forced or classify(result, profiles)
                    side = DocumentSideClassifier().classify(result, kind)
            records.append((source, prepared, card, result, kind, side))
            add_inference_metrics(metric, service.engine, inference_before)
            notes = list(quality.warnings)
            if isinstance(card, DigitalIDRegion):
                notes = [note for note in notes if not note.startswith("Приблизьте документ:")]
                if heights and metric["median_text_height_source"] < 12:
                    notes.append(
                        "Мелкий текст исходного скриншота. Использовано увеличение; для надёжности загрузите исходник большего разрешения."
                    )
            warnings.extend(notes)
            metric["quality_warnings"] = len(notes)
            metric["quality_status"] = "DOCUMENT_IMAGE_LOW_QUALITY" if notes else "OK"
        except Cancelled:
            raise
        except AppError as exc:
            metric["status"] = "OCR_ENGINE_ERROR"
            cacheable = False
            warnings.append(f"Фото {index + 1}: {exc}")
    if not any(r[3].blocks for r in records):
        from document_excel.infrastructure.ocr_diagnostics import MESSAGES, OcrProblem

        code = OcrProblem.OCR_ENGINE_ERROR if not records else OcrProblem.NO_TEXT_DETECTED
        for metric in processing_metrics.values():
            metric["status"] = code.value
        raise OcrFailure(MESSAGES[code], code.value, processing_metrics)
    physical = any(r[4] == DocumentType.KG_PHYSICAL_ID for r in records)
    digital = not physical and any(r[4] == DocumentType.KG_DIGITAL_ID for r in records)
    candidates, sides, texts, diagnostics, regions, reports = [], {}, {}, {}, {}, {}
    kinds = set()
    label_seen = False
    for index, (source, prepared, card, result, kind, side) in enumerate(records):
        if cancel.is_set():
            raise Cancelled()
        metric = processing_metrics[source.id]
        recovery_before = dict(getattr(service.engine, "request_totals", {}))
        service.engine.cache_context = (source.case_id, source.id, PREPROCESSING_VERSION, repr(profiles))
        progress(65 + int(index * 30 / len(records)), f"Фото {index + 1} · извлечение и проверка полей")
        if physical:
            kind = DocumentType.KG_PHYSICAL_ID
            side = DocumentSideClassifier().classify(result, kind)
        elif digital and source.role != DocumentSide.ADDITIONAL:
            kind = DocumentType.KG_DIGITAL_ID
            side = DocumentSideClassifier().classify(result, kind)
        prechecked_mrz = None
        # Some recognizers read the reverse-side labels poorly. Three optical
        # row bands only propose a search; only a checked TD1 establishes BACK.
        # Neither file order nor the currently selected tab is evidence.
        if (
            physical
            and side == DocumentSide.UNKNOWN
            and source.role != DocumentSide.ADDITIONAL
            and callable(getattr(service.engine, "recognize_mrz_line", None))
            and len(MRZRecognizer.line_boxes(card.image)) == 3
        ):
            try:
                prechecked_mrz = MRZRecognizer(service.engine).recognize(card.image, source.id, cancel)
                if prechecked_mrz.valid:
                    side = DocumentSide.BACK
            except Cancelled:
                raise
            except AppError:
                cacheable = False
                warnings.append(
                    "Не удалось проверить сторону по MRZ. Укажите сторону после сверки фотографии."
                )
        if side == DocumentSide.UNKNOWN:
            side = source.role
        if source.role == DocumentSide.ADDITIONAL:
            side = source.role
        profile = next((p for p in profiles if p.id == kind), None)
        phase_started = time.perf_counter()
        extracted = FieldExtractor(profile).extract(result, bool(prepared.quality.warnings))
        metric["stages"]["field_extraction_seconds"] = time.perf_counter() - phase_started
        metric["initial_candidates"] = len(extracted)
        version = "processed" if source.processed_confirmed and source.processed_path else "original"
        if physical:
            profile = KyrgyzPhysicalIDProfile()
            # Generic word/label pairing may mix scripts or incomplete fragments.
            # Only geometrically assembled FRONT groups contribute visual names.
            extracted = [c for c in extracted if c.field_name not in {*NAMES, "full_name"}]
            names = FullNameExtractor()
            groups = names.extract(result, side)
            try:
                groups = names.recover(card.image, result, groups, service.engine, cancel)
            except Cancelled:
                raise
            except AppError:
                cacheable = False
                warnings.append("Дополнительное чтение ФИО не завершено. Сверьте лицевую сторону.")
            extracted.extend(groups.candidates)
            expected = set()
            for area_name, area in profile.ocr_regions.items():
                prefix, _, field = area_name.partition(".")
                if prefix != side.value.lower() or field == "mrz":
                    continue
                expected.add(field)
                rx, ry, rw, rh = area
                h, w = card.image.shape[:2]
                bbox = groups.regions.get(field, (round(rx * w), round(ry * h), round(rw * w), round(rh * h)))
                regions[field] = {
                    "image_id": source.id,
                    "bbox": card.original_bbox(bbox),
                    "image_version": version,
                }
            usable = {
                c.field_name
                for c in extracted
                if c.confidence >= 0.85
                and c.validation_status != FieldStatus.INVALID
                and not validate_field(c.field_name, c.normalized_value)
            }
            missing = expected - usable - set(NAMES)
            if missing:
                try:
                    recovered = layout_candidates(card, result, service.engine, cancel, side, fields=missing)
                except Cancelled:
                    raise
                except AppError:
                    cacheable = False
                    recovered = []
                    warnings.append("Дополнительное чтение области не завершено. Сверьте поля документа.")
                recovered_fields = {c.field_name for c in recovered}
                extracted = [c for c in extracted if c.field_name not in recovered_fields] + recovered
            if side == DocumentSide.BACK:
                progress(80 + int(index * 15 / len(records)), f"Фото {index + 1} · проверка MRZ")
                try:
                    report = prechecked_mrz or MRZRecognizer(service.engine).recognize(
                        card.image, source.id, cancel
                    )
                    report.source_bbox = card.original_bbox(report.source_bbox)
                    report.image_version = version
                    reports[source.id] = report
                    if not report.valid:
                        warnings.append("MRZ не прошла проверку. Сверьте оборот документа.")
                except Cancelled:
                    raise
                except AppError:
                    cacheable = False
                    warnings.append("Дополнительное чтение MRZ не завершено. Сверьте оборот документа.")
        elif digital:
            try:
                extracted, digital_layout, recovery_stats = recover_digital_fields(
                    card.image, result, profile, extracted, service.engine, cancel
                )
                if recovery_stats["failed"]:
                    cacheable = False
                    warnings.append(
                        "Уточнение цифрового поля не завершено. Найденные данные сохранены для проверки."
                    )
                for field, band in digital_layout.bands.items():
                    if field.startswith("_"):
                        continue
                    regions[field] = {
                        "image_id": source.id,
                        "bbox": card.original_bbox(band.bbox),
                        "image_version": version,
                    }
            except Cancelled:
                raise
            except AppError:
                cacheable = False
                warnings.append(
                    "Дополнительное чтение цифрового поля не завершено. Найденные данные сохранены."
                )
        elif callable(reader := getattr(service.engine, "recognize_identifier_line", None)):
            try:
                extracted = refine_identifier(card.image, result, extracted, reader, cancel)
                extracted = refine_mkk_authority(card.image, result, extracted, reader, cancel)
            except Cancelled:
                raise
            except AppError:
                cacheable = False
                warnings.append("Дополнительное чтение кода не завершено. Сверьте строку документа.")
        for candidate in extracted:
            candidate.source_bbox = card.original_bbox(candidate.source_bbox)
            candidate.source_side = side
            candidate.image_version = version
            if candidate.engine_name == "unknown":
                candidate.engine_name = result.engine_name
            if not candidate.model_version:
                candidate.model_version = result.model_version
            if candidate.variant == "original":
                candidate.variant = result.variant
        candidates.extend(extracted)
        add_inference_metrics(metric, service.engine, recovery_before)
        metric.update(
            {
                "profile": str(kind),
                "side": str(side),
                "candidates": len(extracted),
                "rejected_candidates": sum(c.validation_status == FieldStatus.INVALID for c in extracted),
            }
        )
        metric["status"] = (
            "NO_TEXT_DETECTED"
            if not result.blocks
            else "TEXT_DETECTED_BUT_FIELD_NOT_FOUND"
            if not extracted
            else "CANDIDATES_FOUND"
        )
        kinds.add(kind)
        sides[source.id] = side
        texts[source.id] = result.text
        diagnostics[source.id] = [
            {
                "bbox": card.original_bbox(b.bbox),
                "confidence": round(b.confidence, 3),
                "characters": len(b.text),
                "variant": result.variant,
                "engine": result.engine_name,
                "model": result.model_version,
                "script": b.script,
            }
            for b in result.blocks
        ]
        label_seen |= any(
            word in result.text.upper() for word in ("РЕГИСТРАЦ", "ПРОПИСК", "REGISTRATION", "КАТТОО ДАРЕГИ")
        )
    for report in reports.values():
        candidates.extend(mrz_candidates(report, candidates))
    kind = DocumentType.KG_PHYSICAL_ID if physical else next((k for k in kinds if k != "unknown"), "unknown")
    if cancel.is_set():
        raise Cancelled()
    output = RecognitionOutput(
        candidates,
        kind,
        list(dict.fromkeys(warnings)),
        label_seen,
        time.perf_counter() - started,
        sides,
        texts,
        reports,
        diagnostic_blocks=diagnostics,
        field_regions=regions,
        processing_metrics=processing_metrics,
    )
    with _LOCK:
        if not cacheable or cache_key(service, images, forced, load_document_profiles()) != key:
            progress(100, "Распознавание завершено. Сверьте данные.")
            return output
        _CACHE[key] = copy.deepcopy(output)
        # Resolution fills UNKNOWN roles from this exact OCR result. Repeating
        # the same case after that derived UI update still uses the same cache.
        inferred = tuple(
            (*i[:5], str(sides.get(i[1], i[5]))) if i[5] == str(DocumentSide.UNKNOWN) else i for i in key[1]
        )
        _CACHE[(key[0], inferred, *key[2:])] = copy.deepcopy(output)
        while len(_CACHE) > 16:
            _CACHE.popitem(last=False)
    progress(100, "Распознавание завершено. Сверьте данные.")
    return output
