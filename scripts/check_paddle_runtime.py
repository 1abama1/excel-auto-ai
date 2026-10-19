"""Real PP-OCRv5 inference against generated pixels; no identity documents."""

from __future__ import annotations

import argparse
import json
import threading
import time
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont

from document_excel.domain.errors import Cancelled
from document_excel.infrastructure.paddle_ocr import _POOL, PaddleOCREngine
from document_excel.infrastructure.storage import Settings


def check(models: Path) -> dict:
    settings = Settings(paddle_model_directory=str(models), timeout=120)
    engine = PaddleOCREngine(settings)
    diagnostic = engine.diagnose()
    if diagnostic.error:
        raise RuntimeError(diagnostic.error)
    font = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 45)
    results = {}
    pid = None
    for key, text, language in (
        ("kyrgyz", "ТЕСТ ҮӨҢ үөң 0123456789", "rus"),
        ("latin", "SYNTHETIC ID0000001", "eng"),
    ):
        image = Image.new("RGB", (1100, 85), "white")
        ImageDraw.Draw(image).text((10, 10), text, font=font, fill="black")
        read = engine.recognize_physical_value(np.asarray(image), key, threading.Event(), language)
        results[key] = {
            "expected": text,
            "observed": read.text,
            "exact": read.text == text,
            "confidence": min((b.confidence for b in read.blocks), default=0),
        }
        current_pid = _POOL.process.pid
        if pid is not None:
            results["model_process_reused"] = current_pid == pid
        pid = current_pid
    image = Image.new("RGB", (1100, 400), "white")
    for i, text in enumerate(("ФАМИЛИЯ: СИНТЕТИКОВ", "ИМЯ: ТЕСТ", "ДАТА РОЖДЕНИЯ: 07.11.1990")):
        ImageDraw.Draw(image).text((15, 10 + i * 100), text, font=font, fill="black")
    read = engine.recognize(np.asarray(image), "generated", "original", threading.Event())
    results["document"] = {"blocks": len(read.blocks), "text": read.text, "engine": read.engine_name}
    cancel = threading.Event()
    outcome = []

    def active_read():
        try:
            uncached = np.pad(np.asarray(image), ((0, 30), (0, 20), (0, 0)), constant_values=255)
            engine.recognize(uncached, "cancel-synthetic", "original", cancel)
            outcome.append("completed")
        except Cancelled:
            outcome.append("cancelled")

    worker = threading.Thread(target=active_read)
    worker.start()
    deadline = time.monotonic() + 2
    while not _POOL.lock.locked() and worker.is_alive() and time.monotonic() < deadline:
        time.sleep(0.001)
    started = time.monotonic()
    cancel.set()
    worker.join(timeout=5)
    results["cancellation"] = {
        "status": outcome,
        "model_process_released": _POOL.process is None,
        "response_seconds": round(time.monotonic() - started, 3),
    }
    _POOL.close()
    return results


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("models", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    report = json.dumps(check(args.models), ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(report, encoding="utf-8")
    print(report)
