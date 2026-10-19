from __future__ import annotations

import json
import logging
import os
import re
import shutil
import tempfile
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Literal
from uuid import uuid4

from pydantic import BaseModel, Field, ValidationError, field_validator

from document_excel.domain.errors import AppError

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}


class Settings(BaseModel):
    ocr_engine: Literal["paddle", "rapidocr", "tesseract"] = "rapidocr"
    rapid_model_directory: str = ""
    paddle_model_directory: str = ""
    paddle_device: Literal["cpu"] = "cpu"
    paddle_cpu_threads: int = Field(default=2, ge=1, le=8)
    auto_ocr: bool = False  # Legacy setting; imports always wait for explicit OCR.
    tesseract_path: str = ""
    tessdata_dir: str = ""
    languages: list[str] = Field(default_factory=lambda: ["rus", "eng"])
    psm: int = Field(default=6, ge=3, le=13)
    timeout: int = Field(default=30, ge=5, le=120)
    preprocessing: str = "auto"
    max_parallel: int = Field(default=1, ge=1, le=2)
    output_directory: str = ""
    image_directory: str = ""
    template_directory: str = ""
    temp_directory: str = ""
    default_template: str = ""
    default_profile: str = ""
    theme: str = "light"
    ui_scale: int = Field(default=100, ge=80, le=150)
    history_enabled: bool = True
    copy_images: bool = False
    window_width: int = 1366
    window_height: int = 820

    @field_validator("auto_ocr", mode="before")
    @classmethod
    def manual_ocr_only(cls, value):
        return False

    @field_validator("paddle_device", mode="before")
    @classmethod
    def cpu_only(cls, value):
        # Upgrade legacy gpu:0 preferences without discarding other settings.
        return "cpu"


def app_data_directory() -> Path:
    override = os.environ.get("DOCUMENT_EXCEL_DATA_DIR")
    if override:
        return Path(override)
    root = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / ".local" / "share")))
    return root / "DocumentExcel"


def atomic_json(path: Path, data: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix=".settings-", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as stream:
            json.dump(data, stream, ensure_ascii=False, indent=2)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(name, path)
    finally:
        Path(name).unlink(missing_ok=True)


class SettingsRepository:
    def __init__(self, root: Path | None = None) -> None:
        self.root = root or app_data_directory()
        self.path = self.root / "settings.json"
        self.warning = ""

    def load(self) -> Settings:
        if not self.path.exists():
            return Settings()
        try:
            return Settings.model_validate_json(self.path.read_text(encoding="utf-8"))
        except (OSError, ValidationError, ValueError):
            self.warning = "Не удалось прочитать настройки. Используются стандартные значения."
            return Settings()

    def save(self, settings: Settings) -> None:
        try:
            atomic_json(self.path, settings.model_dump())
        except OSError as exc:
            raise AppError("Не удалось сохранить настройки. Проверьте доступ к папке приложения.") from exc


class HistoryRepository:
    """Only time, operation and outcome; no names, file paths or document values."""

    def __init__(self, root: Path) -> None:
        self.path = root / "history.json"

    def read(self) -> list[dict[str, str]]:
        if not self.path.exists():
            return []
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            return data if isinstance(data, list) else []
        except (OSError, ValueError):
            logging.getLogger("document_excel").warning("history_read_failed")
            return []

    def add(self, operation: str, result: str) -> None:
        records = self.read()
        records.append(
            {
                "time": datetime.now().astimezone().isoformat(timespec="seconds"),
                "operation": operation,
                "result": result,
            }
        )
        atomic_json(self.path, records[-200:])

    def clear(self) -> None:
        self.path.unlink(missing_ok=True)


def configure_logging(root: Path) -> None:
    root.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("document_excel")
    logger.setLevel(logging.INFO)
    handler = RotatingFileHandler(root / "diagnostics.log", maxBytes=256_000, backupCount=2, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(handler)


def safe_filename(value: str, fallback: str = "document") -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip(" .")[:100]
    reserved = {
        "CON",
        "PRN",
        "AUX",
        "NUL",
        *(f"COM{i}" for i in range(1, 10)),
        *(f"LPT{i}" for i in range(1, 10)),
    }
    if not value or value.split(".")[0].upper() in reserved:
        return fallback
    return value


def collect_images(paths: list[Path]) -> tuple[list[Path], list[str]]:
    found: list[Path] = []
    errors: list[str] = []
    seen: set[Path] = set()
    for path in paths:
        try:
            # Folder import is shallow; it never follows directory links into unrelated trees.
            children = sorted(path.iterdir()) if path.is_dir() else [path]
            for child in children:
                if child.is_symlink() or not child.is_file():
                    continue
                if child.suffix.lower() not in IMAGE_EXTENSIONS:
                    if not path.is_dir():
                        errors.append("Выбран неподдерживаемый формат изображения.")
                    continue
                resolved = child.resolve()
                if resolved not in seen:
                    seen.add(resolved)
                    found.append(resolved)
        except OSError:
            errors.append("Нет доступа к выбранному файлу или папке.")
    return found, errors


def copy_originals(paths: list[Path], destination: Path, case_id: str) -> list[Path]:
    folder = destination / safe_filename(case_id)
    folder.mkdir(parents=True, exist_ok=True)
    copied: list[Path] = []
    for source in paths:
        target = folder / f"image-{uuid4().hex[:12]}{source.suffix.lower()}"
        try:
            with source.open("rb") as src, target.open("xb") as dst:
                shutil.copyfileobj(src, dst)
            copied.append(target)
        except OSError as exc:
            target.unlink(missing_ok=True)
            raise AppError("Не удалось скопировать фотографии. Проверьте права и свободное место.") from exc
    return copied
