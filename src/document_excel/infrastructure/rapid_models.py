"""Pinned PP-OCRv5 ONNX files. No downloading during recognition."""

from __future__ import annotations

import hashlib
import json
import os
from functools import lru_cache
from pathlib import Path

from document_excel.domain.errors import AppError

RAPID_VERSION = "3.8.1"
ORT_VERSION = "1.23.2"
MODEL_FILES = {
    "det": (
        "ch_PP-OCRv5_det_mobile.onnx",
        "4d97c44a20d30a81aad087d6a396b08f786c4635742afc391f6621f5c6ae78ae",
    ),
    "cyrillic": (
        "cyrillic_PP-OCRv5_rec_mobile.onnx",
        "90f761b4bfcce0c8c561c0cb5c887b0971d3ec01c32164bdf7374a35b0982711",
    ),
    "latin": (
        "en_PP-OCRv5_rec_mobile.onnx",
        "c3461add59bb4323ecba96a492ab75e06dda42467c9e3d0c18db5d1d21924be8",
    ),
}


def model_root(configured=""):
    if configured:
        return Path(configured).resolve()
    override = os.environ.get("DOCUMENT_EXCEL_RAPID_MODELS")
    if override:
        return Path(override).resolve()
    return Path(__file__).resolve().parents[3] / "models" / "rapidocr"


def validate_bundle(root):
    try:
        manifest = json.loads((root / "manifest.json").read_text(encoding="utf-8"))
        if manifest.get("backend") != "rapidocr" or manifest.get("ocr_version") != "PP-OCRv5":
            raise ValueError("Wrong backend/model format")
        stamps = tuple(
            (name, digest, (root / name).stat().st_size, (root / name).stat().st_mtime_ns)
            for name, digest in MODEL_FILES.values()
        )
        _validate(str(root), stamps)
    except (OSError, ValueError, KeyError, TypeError) as exc:
        raise AppError(
            "Локальные ONNX-модели PP-OCRv5 отсутствуют или повреждены. Выберите установленную папку RapidOCR."
        ) from exc


@lru_cache(maxsize=2)
def _validate(root, stamps):
    for filename, digest, _, _ in stamps:
        with (Path(root) / filename).open("rb") as stream:
            if hashlib.file_digest(stream, "sha256").hexdigest() != digest:
                raise ValueError("ONNX checksum mismatch")
