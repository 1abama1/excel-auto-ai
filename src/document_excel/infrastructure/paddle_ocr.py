"""PaddleOCR 3.x adapter with one reusable, cancellable inference process.

No network URLs, defaults, or model downloads are passed to Paddle. Only an
installer-verified local bundle can reach the backend constructor.
"""

from __future__ import annotations

import atexit
import hashlib
import importlib.metadata
import json
import multiprocessing as mp
import os
import re
import socket
import tempfile
import threading
import time
from collections import OrderedDict
from dataclasses import replace
from math import ceil, floor
from typing import Any

import cv2
import numpy as np

from document_excel.domain.errors import AppError, Cancelled
from document_excel.infrastructure.ocr_contract import OcrDiagnostic, OcrResult, TextBlock
from document_excel.infrastructure.paddle_models import (
    CYRILLIC,
    DETECTION,
    ENGLISH,
    PADDLE_VERSION,
    PADDLEOCR_VERSION,
    model_root,
    native_model_root,
    validate_bundle,
)
from document_excel.infrastructure.storage import Settings

cv2.setNumThreads(2)


def script_of(text: str) -> str:
    cyrillic = bool(re.search(r"[\u0400-\u052f]", text))
    latin = bool(re.search(r"[A-Za-z]", text))
    return (
        "mixed"
        if cyrillic and latin
        else "cyrillic"
        if cyrillic
        else "latin"
        if latin
        else "digits"
        if text.isdigit()
        else "common"
    )


def decode_prediction(prediction: Any, image_id: str, variant: str, width: int, height: int) -> OcrResult:
    """3.x rec_polys are aligned with rec_texts; dt_polys are NOT filtered."""
    texts, scores, polys = (prediction[key] for key in ("rec_texts", "rec_scores", "rec_polys"))
    if not len(texts) == len(scores) == len(polys):
        raise ValueError("Mismatched PaddleOCR result arrays")
    blocks = []
    for index, (text, score, polygon) in enumerate(zip(texts, scores, polys, strict=True)):
        points = np.asarray(polygon, dtype=float)
        if points.shape != (4, 2) or not np.isfinite(points).all() or not np.isfinite(score):
            raise ValueError("Invalid PaddleOCR geometry/score")
        left, top = max(0, floor(points[:, 0].min())), max(0, floor(points[:, 1].min()))
        right, bottom = min(width, ceil(points[:, 0].max())), min(height, ceil(points[:, 1].max()))
        if text.strip() and right > left and bottom > top:
            blocks.append(
                TextBlock(
                    text.strip(),
                    min(1.0, max(0.0, float(score))),
                    (left, top, right - left, bottom - top),
                    (1, 1, index),
                    script_of(text),
                )
            )
    return OcrResult(
        blocks, image_id, variant, width, height, "paddle", "PP-OCRv5 / PaddleOCR " + PADDLEOCR_VERSION
    )


class _TimedPredictor:
    """Measure generator advancement, excluding time spent by its caller."""

    def __init__(self, predictor, times, key):
        self.predictor, self.times, self.key = predictor, times, key

    def __getattr__(self, name):
        return getattr(self.predictor, name)

    def _iterate(self, call, *args, **kwargs):
        start = time.perf_counter()
        iterator = iter(call(*args, **kwargs))
        self.times[self.key] = self.times.get(self.key, 0.0) + time.perf_counter() - start
        while True:
            start = time.perf_counter()
            try:
                value = next(iterator)
            except StopIteration:
                break
            finally:
                self.times[self.key] = self.times.get(self.key, 0.0) + time.perf_counter() - start
            yield value

    def __call__(self, *args, **kwargs):
        yield from self._iterate(self.predictor, *args, **kwargs)

    def predict(self, *args, **kwargs):
        yield from self._iterate(self.predictor.predict, *args, **kwargs)


class _Backend:
    """Created only in the child, never in Qt's thread. RGB -> Paddle BGR."""

    def __init__(self, root: str, device: str, threads: int):
        from paddleocr import PaddleOCR, TextRecognition

        root = native_model_root(root)
        self.root, self.device, self.threads = root, device, threads
        self.pipeline = PaddleOCR(
            text_detection_model_name=DETECTION,
            text_detection_model_dir=f"{root}/{DETECTION}",
            text_recognition_model_name=CYRILLIC,
            text_recognition_model_dir=f"{root}/{CYRILLIC}",
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=False,
            text_recognition_batch_size=4,
            text_rec_score_thresh=0.0,
            text_det_limit_side_len=1600,
            text_det_limit_type="max",
            device=device,
            cpu_threads=threads,
            enable_mkldnn=False,
        )
        self.stage_times: dict[str, float] = {}
        outer_pipeline = self.pipeline.paddlex_pipeline
        # PaddleX 3.7.2 forwards reads with __getattr__, but assignments on
        # that facade do not reach the CPU inference pipeline.
        pipeline = getattr(outer_pipeline, "_pipeline", outer_pipeline)
        pipeline.text_det_model = _TimedPredictor(
            pipeline.text_det_model, self.stage_times, "text_detection_seconds"
        )
        pipeline.text_rec_model = _TimedPredictor(
            pipeline.text_rec_model, self.stage_times, "text_recognition_seconds"
        )
        # PaddleX 3.7.2 forwards this public predictor from its single-device
        # pipeline. Reuse it for crops: do not load Cyrillic weights twice.
        self.recognizers: dict[str, Any] = {CYRILLIC: self.pipeline.paddlex_pipeline.text_rec_model}
        self.recognition_class = TextRecognition
        self.cache: OrderedDict[tuple, OcrResult] = OrderedDict()

    def line(self, pixels, model, image_id, variant):
        h, w = pixels.shape[:2]
        # Large empty margins compress short values during neural resizing.
        # Trim only the recognition copy and retain honest crop coordinates.
        gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
        mask = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
        ys, xs = np.where(mask > 0)
        if not len(xs) or float(gray.std()) < 1:
            return OcrResult([], image_id, variant, w, h, "paddle", "PP-OCRv5 / " + model)
        left, top = max(0, int(xs.min()) - 3), max(0, int(ys.min()) - 3)
        right, bottom = min(w, int(xs.max()) + 4), min(h, int(ys.max()) + 4)
        tight = pixels[top:bottom, left:right]
        initialization_started = time.perf_counter()
        if model not in self.recognizers:
            self.recognizers[model] = self.recognition_class(
                model_name=model,
                model_dir=f"{self.root}/{model}",
                device=self.device,
                cpu_threads=self.threads,
                enable_mkldnn=False,
            )
        initialization_seconds = time.perf_counter() - initialization_started
        prediction_started = time.perf_counter()
        result = next(iter(self.recognizers[model].predict(tight)))
        recognition_seconds = time.perf_counter() - prediction_started
        text = result["rec_text"].strip()
        score = float(result["rec_score"])
        if not np.isfinite(score):
            raise ValueError("Invalid recognition confidence")
        blocks = (
            [
                TextBlock(
                    text,
                    min(1.0, max(0.0, score)),
                    (left, top, right - left, bottom - top),
                    (1, 1, 1),
                    script_of(text),
                )
            ]
            if text
            else []
        )
        return OcrResult(
            blocks,
            image_id,
            variant,
            w,
            h,
            "paddle",
            "PP-OCRv5 / " + model,
            timings={
                "text_recognition_seconds": recognition_seconds,
                "recognizer_initialization_seconds": initialization_seconds,
            },
        )

    def recognize(self, pixels, image_id, variant, mode, context=""):
        if mode == "clear":
            self.cache.clear()
            return OcrResult([], image_id, variant, 1, 1)
        key = (
            context,
            image_id,
            variant,
            mode,
            pixels.shape,
            hashlib.sha256(memoryview(np.ascontiguousarray(pixels))).digest(),
        )
        if key in self.cache:
            self.cache.move_to_end(key)
            return replace(
                self.cache[key],
                image_id=image_id,
                variant=variant,
                timings={"cache_hit": True, "text_detection_seconds": 0.0, "text_recognition_seconds": 0.0},
            )
        pixels = (
            cv2.cvtColor(pixels, cv2.COLOR_GRAY2BGR)
            if pixels.ndim == 2
            else cv2.cvtColor(pixels, cv2.COLOR_RGB2BGR)
        )
        if mode != "document":
            started = time.perf_counter()
            result = self.line(pixels, ENGLISH if mode == "latin" else CYRILLIC, image_id, variant)
            result.timings.setdefault("text_recognition_seconds", time.perf_counter() - started)
            result.timings["line_total_seconds"] = time.perf_counter() - started
        else:
            result = self.document(pixels, image_id, variant)
        self.cache[key] = result
        if len(self.cache) > 64:
            self.cache.popitem(last=False)
        return result

    def document(self, pixels, image_id, variant):
        self.stage_times.clear()
        prediction = next(iter(self.pipeline.predict(pixels)))
        result = decode_prediction(prediction, image_id, variant, pixels.shape[1], pixels.shape[0])
        result.timings = dict(self.stage_times)
        return result


def _create_backend(root, device, threads):
    if device != "cpu":
        raise AppError("Эта поставка OCR работает на CPU; CUDA не требуется.")
    from pathlib import Path

    manifest = json.loads((Path(root) / "manifest.json").read_text(encoding="utf-8"))
    if manifest.get("backend") == "rapidocr":
        from document_excel.infrastructure.rapid_ocr import _RapidBackend

        return _RapidBackend(root, device, threads)
    return _Backend(root, device, threads)


def _serve(connection, root, device, threads):
    # Disables PaddleX's startup source probe. Every enabled module has an
    # explicit local path; optional download-capable modules are disabled.
    os.environ["PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK"] = "True"
    os.environ["HF_HUB_OFFLINE"] = "1"
    os.environ.setdefault(
        "PADDLE_PDX_CACHE_HOME", os.path.join(tempfile.gettempdir(), "document-excel-paddle-cache")
    )

    def offline(*_args, **_kwargs):
        raise OSError("OCR worker has no network access")

    socket.socket.connect = offline  # type: ignore[method-assign]
    socket.create_connection = offline
    try:
        backend = None
        while True:
            connection.send(("ready", None))
            request = connection.recv()
            if request is None:
                break
            pixels, image_id, variant, mode, *context = request
            initialization_seconds = 0.0
            if backend is None:
                initialization_started = time.perf_counter()
                backend = _create_backend(root, device, threads)
                initialization_seconds = time.perf_counter() - initialization_started
            result = backend.recognize(pixels, image_id, variant, mode, *context)
            result.timings["backend_initialization_seconds"] = initialization_seconds
            connection.send((True, result))
            del pixels, request, result
    except (EOFError, BrokenPipeError):
        pass
    except Exception as exc:
        # Backend exceptions can contain recognized identity data and paths.
        try:
            connection.send(
                (False, f"OCR: {type(exc).__name__} при локальном inference. Проверьте runtime и модели.")
            )
        except (BrokenPipeError, OSError):
            pass
    finally:
        connection.close()


class _ProcessPool:
    """At most one model set in RAM, including after settings changes."""

    def __init__(self):
        self.lock = threading.Lock()
        self.process: Any = None
        self.connection: Any = None
        self.key: tuple | None = None

    def close(self):
        if self.process is not None:
            if self.process.is_alive():
                self.process.terminate()
            self.process.join(timeout=2)
            if self.process.is_alive():
                self.process.kill()
                self.process.join(timeout=2)
            self.process.close()
            self.process = None
        if self.connection is not None:
            self.connection.close()
            self.connection = None
        self.key = None

    def request(self, key, payload, cancel, timeout):
        while not self.lock.acquire(timeout=0.05):
            if cancel.is_set():
                raise Cancelled()
        try:
            if cancel.is_set():
                raise Cancelled()
            started = time.monotonic()
            if self.key != key or self.process is None or not self.process.is_alive():
                self.close()
                context = mp.get_context("spawn")
                self.connection, child = context.Pipe()
                self.process = context.Process(target=_serve, args=(child, *key[:3]), daemon=True)
                self.process.start()
                child.close()
                self.key = key

            def receive():
                while True:
                    if cancel.is_set():
                        self.close()
                        raise Cancelled()
                    if time.monotonic() - started > timeout:
                        self.close()
                        raise AppError(
                            "OCR превысил timeout. Процесс остановлен; повторите OCR или увеличьте timeout."
                        )
                    if self.connection.poll(0.05):
                        return self.connection.recv()
                    if not self.process.is_alive():
                        self.close()
                        raise AppError("Процесс OCR завершился. Проверьте Windows CPU runtime.")

            ready, _ = receive()
            if ready != "ready":
                self.close()
                raise AppError("OCR: ошибка запуска локального процесса.")
            # The worker has completed imports and is waiting on recv. A large
            # image cannot block send behind slow model initialization.
            self.connection.send(payload)
            ok, value = receive()
            if cancel.is_set():
                self.close()
                raise Cancelled()
            if not ok:
                self.close()
                raise AppError(value)
            return value
        except (EOFError, BrokenPipeError, OSError) as exc:
            self.close()
            raise AppError("OCR недоступен. Проверьте локальные модели и Windows CPU runtime.") from exc
        finally:
            self.lock.release()


_POOL = _ProcessPool()
atexit.register(_POOL.close)


def clear_cache():
    """Forget personal inference results while retaining loaded model weights."""
    if _POOL.key is not None and _POOL.process is not None and _POOL.process.is_alive():
        try:
            _POOL.request(_POOL.key, (np.zeros((1, 1), np.uint8), "", "", "clear"), threading.Event(), 10)
        except AppError:
            _POOL.close()


class PaddleOCREngine:
    preferred_input = "color"
    regional_variants = ("original", "clahe")

    def __init__(self, settings: Settings):
        self.settings = settings.model_copy(deep=True)
        self.root = model_root(settings.paddle_model_directory)

    def diagnose(self) -> OcrDiagnostic:
        try:
            versions = {name: importlib.metadata.version(name) for name in ("paddleocr", "paddlepaddle")}
            if versions != {"paddleocr": PADDLEOCR_VERSION, "paddlepaddle": PADDLE_VERSION}:
                raise AppError(
                    f"Требуются PaddleOCR {PADDLEOCR_VERSION} и PaddlePaddle {PADDLE_VERSION}. Выполните scripts/setup_paddle.ps1 (Python 3.12)."
                )
            validate_bundle(self.root)
            if self.settings.paddle_device != "cpu":
                raise AppError("Для этой поставки выберите CPU. NVIDIA и CUDA не требуются.")
            return OcrDiagnostic(
                str(self.root), f"PaddleOCR {PADDLEOCR_VERSION} / PP-OCRv5", ["ru", "ky", "en", "digits"]
            )
        except importlib.metadata.PackageNotFoundError:
            return OcrDiagnostic(
                "",
                "",
                [],
                "PaddleOCR не установлен в этом Python. Запустите scripts/setup_paddle.ps1 и Launch-Paddle.cmd (Python 3.12).",
            )
        except AppError as exc:
            return OcrDiagnostic(str(self.root), "PP-OCRv5", [], str(exc))

    def _read(self, image, image_id, variant, cancel, mode="document"):
        if cancel.is_set():
            raise Cancelled()
        if image.dtype != np.uint8 or image.ndim not in (2, 3) or image.size == 0:
            raise AppError("Недопустимое изображение OCR.")
        if image.ndim == 3 and image.shape[2] != 3 or max(image.shape[:2]) > 10000 or image.size > 32_000_000:
            raise AppError("Слишком большое или неподдерживаемое изображение OCR.")
        diagnostic = self.diagnose()
        if diagnostic.error:
            raise AppError(diagnostic.error)
        manifest = self.root / "manifest.json"
        key = (
            str(self.root),
            self.settings.paddle_device,
            self.settings.paddle_cpu_threads,
            manifest.stat().st_mtime_ns,
        )
        context = (getattr(self, "cache_context", ""), self.settings.model_dump_json())
        started = time.perf_counter()
        result = _POOL.request(key, (image, image_id, variant, mode, context), cancel, self.settings.timeout)
        result.timings["request_seconds"] = time.perf_counter() - started
        totals = getattr(self, "request_totals", {})
        totals["ocr_requests"] = totals.get("ocr_requests", 0) + 1
        totals["inference_cache_hits"] = totals.get("inference_cache_hits", 0) + int(
            bool(result.timings.get("cache_hit"))
        )
        for name, seconds in result.timings.items():
            if name.endswith("_seconds"):
                totals[name] = totals.get(name, 0.0) + seconds
        self.request_totals = totals
        return result

    def recognize(self, image, image_id, variant, cancel):
        return self._read(image, image_id, variant, cancel)

    def recognize_with_psm(self, image, image_id, variant, cancel, psm=11):
        # Compatibility with existing card geometry orchestration; PSM has no
        # Paddle equivalent and is never passed into PaddleOCR 3.x.
        return self.recognize(image, image_id, variant, cancel)

    def recognize_identifier_line(self, image, image_id, variant, cancel):
        return self._read(image, image_id, variant, cancel, "latin")

    def recognize_mrz(self, image, image_id, cancel):
        # Three independent TD1 row crops, bypassing general text detection.
        from document_excel.infrastructure.mrz import MRZRecognizer

        boxes = MRZRecognizer.line_boxes(image)
        if len(boxes) != 3:
            h, w = image.shape[:2]
            boxes = [(0, round(i * h / 3), w, round((i + 1) * h / 3) - round(i * h / 3)) for i in range(3)]
        blocks: list[TextBlock] = []
        for index, (x, y, w, h) in enumerate(boxes):
            result = self.recognize_mrz_line(image[y : y + h, x : x + w], image_id, cancel)
            blocks.extend(
                TextBlock(b.text, b.confidence, (x, y, w, h), (1, 1, index), b.script) for b in result.blocks
            )
        return OcrResult(
            blocks, image_id, "mrz", image.shape[1], image.shape[0], result.engine_name, result.model_version
        )

    def recognize_mrz_line(self, image, image_id, cancel):
        return self._read(image, image_id, "mrz", cancel, "latin")

    def recognize_physical_value(self, image, image_id, cancel, language="eng", kind="text"):
        return self._read(
            image, image_id, "physical-region", cancel, "cyrillic" if language == "rus" else "latin"
        )

    def recognize_cyrillic_name(self, image, image_id, cancel):
        return self._read(image, image_id, "physical-name-cyr", cancel, "cyrillic")

    def recognize_text_line(self, image, image_id, variant, cancel, script="cyrillic"):
        if script not in {"cyrillic", "latin"}:
            raise AppError("Неизвестный алфавит строкового распознавания.")
        return self._read(image, image_id, variant, cancel, script)

    def recognize_card_region(self, image, image_id, cancel, kind):
        return self._read(image, image_id, "physical-region", cancel, "latin")

    def recognize_official_region(self, image, image_id, cancel, field, variant):
        return self._read(
            image,
            image_id,
            f"official-{field}-{variant}",
            cancel,
            "latin" if field == "passport_number" else "cyrillic",
        )
