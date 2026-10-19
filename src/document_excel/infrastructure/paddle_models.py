"""Explicit offline PP-OCRv5 bundle. Downloading is a separate installer action."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import sys
import tempfile
from functools import lru_cache
from pathlib import Path

from document_excel.domain.errors import AppError

DETECTION = "PP-OCRv5_mobile_det"
CYRILLIC = "cyrillic_PP-OCRv5_mobile_rec"
ENGLISH = "en_PP-OCRv5_mobile_rec"
MODELS = (DETECTION, CYRILLIC, ENGLISH)
PADDLEOCR_VERSION = "3.7.0"
PADDLE_VERSION = "3.3.1"


def model_root(configured: str = "") -> Path:
    if configured:
        return Path(configured).resolve()
    override = os.environ.get("DOCUMENT_EXCEL_PADDLE_MODELS")
    if override:
        return Path(override).resolve()
    base = (
        Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[3]
    )
    return base / "models" / "paddle"


def sha256(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def _windows_short_path(path: str) -> str:
    import ctypes

    buffer = ctypes.create_unicode_buffer(32768)
    get_path = ctypes.windll.kernel32.GetShortPathNameW
    get_path.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_uint32]
    get_path.restype = ctypes.c_uint32
    size = get_path(path, buffer, len(buffer))
    return buffer.value if 0 < size < len(buffer) else path


def native_model_root(root: str) -> str:
    """Paddle's C++ JSON reader needs an ASCII model path on Windows.

    Python still validates the original Unicode path. Prefer its existing
    8.3 alias; when 8.3 names are disabled, copy only public model weights
    into an ASCII cache. Identity images and results never enter that cache.
    """
    if sys.platform != "win32" or root.isascii():
        return root
    short = _windows_short_path(root)
    if short.isascii():
        return short
    source = Path(root)
    validate_bundle(source)
    parent = Path(os.environ.get("PUBLIC", "C:/Users/Public")) / "Documents/DocumentExcel/ocr-models"
    if not str(parent).isascii():
        raise AppError("Для PaddleOCR выберите папку моделей с латинскими буквами в пути.")
    target = parent / sha256(source / "manifest.json")[:16]
    try:
        if target.is_dir():
            try:
                validate_bundle(target)
                return str(target)
            except AppError:
                pass
        parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="prepare-", dir=parent) as temporary:
            stage = Path(temporary)
            files = [Path("manifest.json")] + [
                Path(model) / name
                for model in MODELS
                for name in ("inference.json", "inference.pdiparams", "inference.yml")
            ]
            for relative in files:
                (stage / relative).parent.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(source / relative, stage / relative)
            validate_bundle(stage)
            for relative in files[1:] + files[:1]:
                (target / relative).parent.mkdir(parents=True, exist_ok=True)
                os.replace(stage / relative, target / relative)
        validate_bundle(target)
        return str(target)
    except OSError as exc:
        raise AppError("Не удалось подготовить модели PaddleOCR. Выберите папку моделей с латинским путём.") from exc


def validate_bundle(root: Path) -> dict:
    try:
        manifest = root / "manifest.json"
        data = json.loads(manifest.read_text(encoding="utf-8"))
        if data["ocr_version"] != "PP-OCRv5" or set(data["models"]) != set(MODELS):
            raise ValueError("Wrong models")
        stamps = []
        for model in MODELS:
            for filename in ("inference.json", "inference.pdiparams", "inference.yml"):
                path = root / model / filename
                stat = path.stat()
                stamps.append(
                    (model, filename, stat.st_size, stat.st_mtime_ns, data["models"][model][filename])
                )
        return _validate(str(root), manifest.stat().st_mtime_ns, tuple(stamps))
    except (OSError, ValueError, KeyError, TypeError) as exc:
        raise AppError(
            "Локальные модели PP-OCRv5 отсутствуют или повреждены. Выполните scripts/setup_paddle.ps1 или выберите папку моделей в настройках."
        ) from exc


@lru_cache(maxsize=2)
def _validate(root: str, manifest_stamp: int, stamps: tuple) -> dict:
    import yaml

    directory = Path(root)
    for model, filename, _size, _stamp, digest in stamps:
        if sha256(directory / model / filename) != digest:
            raise ValueError("Model checksum mismatch")
    alphabets = {}
    for model in MODELS:
        config = yaml.safe_load((directory / model / "inference.yml").read_text(encoding="utf-8"))
        if config["Global"]["model_name"] != model:
            raise ValueError("Model identity mismatch")
        if model != DETECTION:
            alphabets[model] = set(config["PostProcess"]["character_dict"])
    if not set("АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯҮӨҢүөң0123456789") <= alphabets[CYRILLIC]:
        raise ValueError("Cyrillic model lacks Russian/Kyrgyz alphabet")
    if not set("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<") <= alphabets[ENGLISH]:
        raise ValueError("English model lacks TD1 alphabet")
    return {
        "version": "PP-OCRv5",
        "languages": ["ru", "ky", "en", "digits"],
        "manifest_stamp": manifest_stamp,
    }
