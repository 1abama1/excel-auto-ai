from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
from math import ceil, floor
from pathlib import Path

import cv2
import numpy as np
import pytesseract

from document_excel.domain.errors import AppError, Cancelled
from document_excel.infrastructure.ocr_contract import OcrDiagnostic as OcrDiagnostic
from document_excel.infrastructure.ocr_contract import OcrResult as OcrResult
from document_excel.infrastructure.ocr_contract import TextBlock as TextBlock
from document_excel.infrastructure.storage import Settings

_ENGINE_LOCK = threading.Lock()


def find_tesseract(configured: str = "") -> str:
    if configured:
        path = Path(configured)
        return str(path) if path.is_file() else ""
    from_path = shutil.which("tesseract")
    if from_path:
        return from_path
    from_environment = os.environ.get("DOCUMENT_EXCEL_TESSERACT", "")
    if from_environment and Path(from_environment).is_file():
        return from_environment
    candidates = [
        Path(os.environ.get("ProgramFiles", "C:/Program Files")) / "Tesseract-OCR/tesseract.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Programs/Tesseract-OCR/tesseract.exe",
    ]
    if not getattr(sys, "frozen", False):
        candidates.append(Path(__file__).resolve().parents[3] / "tools/tesseract/tesseract.exe")
    return next((str(path) for path in candidates if path.is_file()), "")


class TesseractEngine:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings.model_copy(deep=True)

    def diagnose(self) -> OcrDiagnostic:
        executable = find_tesseract(self.settings.tesseract_path)
        if not executable:
            return OcrDiagnostic(
                "",
                "",
                [],
                "Tesseract не найден. Установите OCR-движок или выберите tesseract.exe в настройках.",
            )
        flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        try:
            version = subprocess.run(
                [executable, "--version"],
                capture_output=True,
                text=True,
                timeout=8,
                creationflags=flags,
                check=True,
            ).stdout.splitlines()[0]
            if not re.search(r"tesseract\s+v?5\.", version):
                return OcrDiagnostic(executable, version, [], "Требуется Tesseract версии 5.x.")
            command = [executable, "--list-langs"]
            if self.settings.tessdata_dir:
                command += ["--tessdata-dir", self.settings.tessdata_dir]
            result = subprocess.run(
                command, capture_output=True, text=True, timeout=8, creationflags=flags, check=True
            )
            languages = [
                line.strip() for line in result.stdout.splitlines() if re.fullmatch(r"[\w/]+", line.strip())
            ]
            missing = set(self.settings.languages) - set(languages)
            error = "Не установлены языковые модели: " + ", ".join(sorted(missing)) if missing else ""
            if not self.settings.languages:
                error = "Выберите хотя бы один язык OCR."
            return OcrDiagnostic(executable, version, languages, error)
        except (OSError, subprocess.SubprocessError, IndexError):
            return OcrDiagnostic(
                executable,
                "",
                [],
                "Не удалось запустить Tesseract. Проверьте путь и доступность traineddata.",
            )

    def recognize(self, image: np.ndarray, image_id: str, variant: str, cancel: threading.Event) -> OcrResult:
        return self.recognize_with_psm(image, image_id, variant, cancel, self.settings.psm)

    def recognize_with_psm(
        self, image: np.ndarray, image_id: str, variant: str, cancel: threading.Event, psm: int = 11
    ) -> OcrResult:
        """Run one bounded alternative without mutating the shared settings.

        RecognitionService decides whether an alternative is useful using field
        validation and coverage. Existing four-argument ``recognize`` callers
        keep their configured segmentation mode.
        """
        return self._recognize(image, image_id, variant, cancel, psm)

    def recognize_identifier_line(
        self, image: np.ndarray, image_id: str, variant: str, cancel: threading.Event
    ) -> OcrResult:
        """One local English line pass for letters confused with leading digits.

        No character whitelist or forced ID prefix: letters must be read from
        the pixels. The caller compares the numeric suffix before using it.
        """
        return self._recognize(
            image,
            image_id,
            variant,
            cancel,
            7,
            languages=["eng"],
            scale_override=3.0,
            border=10,
            extra_config=" -c load_system_dawg=0 -c load_freq_dawg=0",
        )

    def recognize_mrz(self, image, image_id, cancel):
        return self._recognize(
            image,
            image_id,
            "mrz",
            cancel,
            6,
            languages=["eng"],
            border=12,
            scale_override=min(3.0, max(1.0, 1800 / image.shape[1])),
            extra_config=" -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789< -c load_system_dawg=0 -c load_freq_dawg=0",
        )

    def recognize_mrz_line(self, image, image_id, cancel):
        return self._recognize(
            image,
            image_id,
            "mrz",
            cancel,
            7,
            languages=["eng"],
            border=14,
            scale_override=min(4.0, max(1.0, 2200 / image.shape[1])),
            extra_config=" -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789< -c load_system_dawg=0 -c load_freq_dawg=0",
        )

    def recognize_mrz_rawline(self, image, image_id, cancel):
        return self._recognize(
            image,
            image_id,
            "mrz",
            cancel,
            13,
            languages=["eng"],
            border=14,
            scale_override=min(4.0, max(1.0, 2200 / image.shape[1])),
            extra_config=" -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789< -c load_system_dawg=0 -c load_freq_dawg=0",
        )

    def recognize_mrz_span(self, image, image_id, cancel, kind):
        alphabet = (
            "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            if kind.startswith("letters")
            else "0123456789"
            if kind.startswith("digits")
            else "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
        )
        return self._recognize(
            image,
            image_id,
            "mrz",
            cancel,
            10 if kind.endswith("-char") else 7,
            languages=["eng"],
            border=40,
            scale_override=3.0 if kind.endswith("-char") else min(4.0, max(1.0, 55 / image.shape[0])),
            extra_config=f" -c tessedit_char_whitelist={alphabet} -c load_system_dawg=0 -c load_freq_dawg=0",
        )

    def recognize_physical_value(self, image, image_id, cancel, language="eng", kind="text"):
        languages = [language]
        if language == "rus" and "kir" in self.settings.languages:
            languages.append("kir")
        config = " -c load_system_dawg=0 -c load_freq_dawg=0"
        if kind == "name" and language == "eng":
            config += " -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-"
        if kind == "date":
            config += " -c tessedit_char_whitelist=0123456789./- -c preserve_interword_spaces=1"
        if kind == "digits":
            config += " -c tessedit_char_whitelist=0123456789"
        return self._recognize(
            image,
            image_id,
            "physical-region",
            cancel,
            7,
            languages=languages,
            scale_override=min(4.0, max(1.0, 80 / image.shape[0])),
            border=12,
            extra_config=config,
        )

    def recognize_cyrillic_name(self, image, image_id, cancel):
        # Windows command-line encodings can corrupt a Cyrillic -c value.
        # A UTF-8 config contains only the generic alphabet, never personal data.
        alphabet = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯҮӨҢабвгдеёжзийклмнопрстуфхцчшщъыьэюяүөң-"
        languages = ["rus"] + (["kir"] if "kir" in self.settings.languages else [])
        gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
        prepared = cv2.createCLAHE(clipLimit=1.5, tileGridSize=(4, 2)).apply(gray)
        with tempfile.TemporaryDirectory(prefix="document-excel-cyr-") as directory:
            config_path = Path(directory) / "alphabet.cfg"
            config_path.write_text("tessedit_char_whitelist " + alphabet + "\n", encoding="utf-8")
            result = None
            for psm in (7, 13):
                result = self._recognize(
                    prepared,
                    image_id,
                    "physical-name-cyr",
                    cancel,
                    psm,
                    languages=languages,
                    border=30,
                    scale_override=min(4.0, max(1.0, 50 / image.shape[0])),
                    extra_config=f' -c load_system_dawg=0 -c load_freq_dawg=0 "{config_path}"',
                )
                if re.fullmatch(r"[А-ЯЁҮӨҢа-яёүөң -]{2,}", result.text.strip()):
                    break
            return result

    def recognize_card_region(self, image, image_id, cancel, kind):
        config = " -c load_system_dawg=0 -c load_freq_dawg=0"
        if kind == "date":
            config += " -c tessedit_char_whitelist=0123456789./- "
        return self._recognize(
            image,
            image_id,
            "physical-region",
            cancel,
            7,
            languages=["eng"],
            scale_override=3.0,
            border=10,
            extra_config=config,
        )

    def recognize_official_region(self, image, image_id, cancel, field, variant):
        config = " -c load_system_dawg=0 -c load_freq_dawg=0"
        number = field == "passport_number"
        if number:
            config += (
                " -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
            )
        with tempfile.TemporaryDirectory(prefix="document-excel-official-") as directory:
            if not number:
                # This printed field is Cyrillic. Constrain the alphabet, not
                # the issuer code: both Е and О remain equally admissible.
                alphabet = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯҮӨҢ0123456789"
                path = Path(directory) / "alphabet.cfg"
                path.write_text("tessedit_char_whitelist " + alphabet + "\n", encoding="utf-8")
                config += f' "{path}"'
            return self._recognize(
                image,
                image_id,
                f"official-{field}-{variant}",
                cancel,
                13 if variant == "upscale" else 7,
                languages=["eng"] if number else ["rus"],
                scale_override=3.0 if variant == "upscale" else 2.0,
                border=20,
                extra_config=config,
            )

    def _recognize(
        self,
        image: np.ndarray,
        image_id: str,
        variant: str,
        cancel: threading.Event,
        psm: int,
        *,
        languages: list[str] | None = None,
        scale_override: float | None = None,
        border: int = 0,
        extra_config: str = "",
    ) -> OcrResult:
        if cancel.is_set():
            raise Cancelled()
        if psm not in range(3, 14):
            raise AppError("Недопустимый режим сегментации OCR.")
        diagnostic = self.diagnose()
        if diagnostic.error:
            raise AppError(diagnostic.error)
        selected_languages = languages if languages is not None else self.settings.languages
        if set(selected_languages) - set(diagnostic.languages):
            raise AppError("Не установлена языковая модель для уточнения номера документа.")
        config = f"--oem 1 --psm {psm}" + extra_config
        if self.settings.tessdata_dir:
            if '"' in self.settings.tessdata_dir:
                raise AppError("Недопустимые кавычки в пути traineddata.")
            config += f' --tessdata-dir "{self.settings.tessdata_dir}"'
        # pytesseract's executable setting is process-global. Serialize the call so settings
        # changes cannot redirect an already running job to a different executable.
        while not _ENGINE_LOCK.acquire(timeout=0.1):
            if cancel.is_set():
                raise Cancelled()
        try:
            if cancel.is_set():
                raise Cancelled()
            pytesseract.pytesseract.tesseract_cmd = diagnostic.executable
            # Phone cards often contain 10–14-pixel glyphs. A bounded local
            # enlargement helps Tesseract without changing the coordinate
            # contract exposed to callers or modifying the original pixels.
            scale = min(3.0, 1500 / max(image.shape[:2])) if max(image.shape[:2]) < 1200 else 1.0
            if scale_override is not None:
                scale = scale_override
            ocr_image = (
                cv2.resize(image, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
                if scale > 1
                else image
            )
            if border:
                ocr_image = cv2.copyMakeBorder(
                    ocr_image,
                    border,
                    border,
                    border,
                    border,
                    cv2.BORDER_CONSTANT,
                    value=255 if ocr_image.ndim == 2 else (255, 255, 255),
                )
            result = pytesseract.image_to_data(
                ocr_image,
                lang="+".join(selected_languages),
                config=config,
                output_type=pytesseract.Output.DICT,
                timeout=self.settings.timeout,
            )
        except RuntimeError as exc:
            raise AppError(
                "OCR превысил время ожидания. Переснимите документ или увеличьте timeout."
            ) from exc
        except (pytesseract.TesseractError, OSError) as exc:
            raise AppError("Ошибка OCR. Проверьте языковые модели и параметры Tesseract.") from exc
        finally:
            _ENGINE_LOCK.release()
        if cancel.is_set():
            raise Cancelled()
        blocks = []
        for i, text in enumerate(result["text"]):
            if not text.strip() or float(result["conf"][i]) < 0:
                continue
            raw_left, raw_top, raw_width, raw_height = (
                int(result[k][i]) for k in ("left", "top", "width", "height")
            )
            left = max(0, min(image.shape[1], floor((raw_left - border) / scale)))
            top = max(0, min(image.shape[0], floor((raw_top - border) / scale)))
            width = min(image.shape[1], ceil((raw_left + raw_width - border) / scale)) - left
            height = min(image.shape[0], ceil((raw_top + raw_height - border) / scale)) - top
            if width <= 0 or height <= 0:
                continue
            block, paragraph, line = (int(result[k][i]) for k in ("block_num", "par_num", "line_num"))
            blocks.append(
                TextBlock(
                    text.strip(),
                    float(result["conf"][i]) / 100,
                    (left, top, width, height),
                    (block, paragraph, line),
                )
            )
        return OcrResult(
            blocks, image_id, variant, image.shape[1], image.shape[0], "tesseract", diagnostic.version
        )
