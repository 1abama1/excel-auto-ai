"""Refine ambiguous official short codes from pixels, without guessing glyphs."""

from __future__ import annotations

import re
import threading
from collections.abc import Callable

import numpy as np

from document_excel.domain.models import (
    FieldCandidate,
    FieldStatus,
    split_document_identifier,
    split_mkk_authority,
)
from document_excel.infrastructure.ocr import OcrResult


def compact(value: str) -> str:
    return re.sub(r"\s+", "", value).upper()


def refine_mkk_authority(image, result, candidates, recognize_line, cancel):
    """Confirm a confusable zero in a labelled Latin MKK office code optically.

    Existing literal codes and conflicting sources are left alone. A format
    alone never supplies a zero: the independent English ROI must read it.
    """
    if any(c.field_name == "passport_department_code" for c in candidates):
        return candidates
    authorities = [c for c in candidates if c.field_name == "passport_issued_by"]
    if len({compact(c.raw_value) for c in authorities}) != 1:
        return candidates
    initial = authorities[0]
    original = compact(initial.raw_value)
    if (
        initial.validation_status == FieldStatus.INVALID
        or not re.fullmatch(r"MKK[0-9O]{6}", original)
        or "O" not in original[3:]
        or not initial.source_bbox
    ):
        return candidates
    x, y, w, h = initial.source_bbox
    left, top = max(0, x - 4), max(0, y - 3)
    right, bottom = min(image.shape[1], x + w + 4), min(image.shape[0], y + h + 3)
    if right <= left or bottom <= top:
        return candidates
    reread = recognize_line(
        image[top:bottom, left:right], result.image_id, result.variant + "-authority-line", cancel
    )
    match = re.fullmatch(r"(?:AUTHORITY:?)?(MKK[0-9]{6})", compact(reread.text))
    confidence = min((b.confidence for b in reread.blocks), default=0)
    if not match or confidence < 0.85:
        return candidates
    observed = match[1]
    if not all(a == b or a == "O" and b == "0" for a, b in zip(original, observed, strict=True)):
        return candidates
    parts = split_mkk_authority(observed)
    if parts is None:
        return candidates
    authority, code = parts
    issuer = initial.model_copy(
        update={
            "raw_value": observed,
            "normalized_value": authority,
            "confidence": min(initial.confidence, confidence),
            "validation_status": FieldStatus.REVIEW,
            "extraction_method": initial.extraction_method + "+authority-line+split-department",
            "variant": reread.variant,
            "engine_name": reread.engine_name,
            "model_version": reread.model_version,
        }
    )
    department = issuer.model_copy(
        update={
            "field_name": "passport_department_code",
            "raw_value": code,
            "normalized_value": code,
        }
    )
    return [c for c in candidates if c.field_name != "passport_issued_by"] + [issuer, department]


def refine_identifier(
    image: np.ndarray,
    result: OcrResult,
    candidates: list[FieldCandidate],
    recognize_line: Callable,
    cancel: threading.Event,
) -> list[FieldCandidate]:
    if any(c.field_name == "passport_type" for c in candidates):
        return candidates
    ambiguous = [
        c
        for c in candidates
        if c.field_name == "passport_number"
        and c.validation_status != FieldStatus.INVALID
        and re.fullmatch(r"[1IL|][0OD][0-9]{4,12}", compact(c.raw_value))
    ]
    if not ambiguous or len({compact(c.raw_value) for c in ambiguous}) != 1:
        return candidates
    # Conflicting labelled numbers must remain a conflict for human review.
    if len({compact(c.raw_value) for c in candidates if c.field_name == "passport_number"}) != 1:
        return candidates
    initial = ambiguous[0]
    original = compact(initial.raw_value)
    boxes = []
    for line in result.lines():
        words = sorted((b for b in result.blocks if b.line_key == line.line_key), key=lambda b: b.bbox[0])
        for start in range(len(words)):
            for count in range(1, min(4, len(words) - start) + 1):
                group = words[start : start + count]
                if compact("".join(b.text for b in group)) == original:
                    x = min(b.bbox[0] for b in group)
                    y = min(b.bbox[1] for b in group)
                    right = max(b.bbox[0] + b.bbox[2] for b in group)
                    bottom = max(b.bbox[1] + b.bbox[3] for b in group)
                    boxes.append((x, y, right, bottom))
    if len(set(boxes)) != 1:
        return candidates
    x, y, right, bottom = boxes[0]
    left, top = max(0, x - 4), max(0, y - 3)
    right, bottom = min(image.shape[1], right + 4), min(image.shape[0], bottom + 3)
    if right <= left or bottom <= top:
        return candidates
    reread = recognize_line(image[top:bottom, left:right], result.image_id, result.variant, cancel)
    observed = compact(reread.text)
    parts = split_document_identifier(observed)
    # Accept only a literal ID read and the exact same suffix. A numeric 10
    # alone never supplies a type, and a conflicting digit is never corrected
    # by taking the more convenient part of a second OCR attempt.
    if not parts or parts[0] != "ID" or original[2:] != parts[1]:
        return candidates
    if not reread.blocks:
        return candidates
    confidence = min(initial.confidence, min(b.confidence for b in reread.blocks))
    number = initial.model_copy(
        update={
            "raw_value": observed,
            "normalized_value": parts[1],
            "confidence": confidence,
            "extraction_method": initial.extraction_method + "+identifier-line+split-type",
            "validation_status": FieldStatus.REVIEW,
        }
    )
    prefix = number.model_copy(
        update={"field_name": "passport_type", "raw_value": parts[0], "normalized_value": parts[0]}
    )
    return [c for c in candidates if c.field_name != "passport_number"] + [prefix, number]
