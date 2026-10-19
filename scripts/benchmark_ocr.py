"""Generate authorized synthetic documents or use an explicitly authorized local manifest."""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path

import cv2
import numpy as np
import psutil
from PIL import Image, ImageDraw, ImageFont

from document_excel.application.benchmark import BenchmarkDocument, benchmark
from document_excel.infrastructure.paddle_ocr import _POOL
from document_excel.infrastructure.physical_verification import synthetic_plastic_images
from document_excel.infrastructure.storage import Settings


def synthetic_documents(directory: Path) -> list[BenchmarkDocument]:
    paths = synthetic_plastic_images(directory)
    expected = {
        "full_name": "СИНТЕТИКОВ ТЕСТ ПРИМЕРОВИЧ",
        "birth_date": "1990-11-07",
        "personal_identifier": "00123456789012",
        "document_number": "ID0000001",
        "passport_issued_date": "2024-11-22",
        "expiration_date": "2034-11-22",
        "passport_issued_by": "ТЕСТ ОРГАН",
        "gender": "MALE",
        "citizenship": "KGZ",
    }
    # Exact same generated source pixels for both engines, plus photographed
    # perspective and low-contrast variants. No private document is accessed.
    documents = [BenchmarkDocument("physical-clean", paths, expected)]
    for mode in ("perspective", "contrast"):
        transformed = []
        for path in paths:
            pixels = np.asarray(Image.open(path).convert("RGB"))
            h, w = pixels.shape[:2]
            if mode == "perspective":
                matrix = cv2.getPerspectiveTransform(
                    np.float32([[0, 0], [w, 0], [w, h], [0, h]]),
                    np.float32([[120, 110], [w + 20, 30], [w + 90, h + 50], [30, h + 120]]),
                )
                pixels = cv2.warpPerspective(pixels, matrix, (w + 160, h + 160), borderValue=(90, 90, 90))
            else:
                pixels = np.clip(pixels.astype(float) * 0.35 + 145, 0, 255).astype(np.uint8)
            target = directory / f"{mode}-{path.name}"
            Image.fromarray(pixels).save(target)
            transformed.append(target)
        documents.append(BenchmarkDocument("physical-" + mode, transformed, expected))
    image = Image.new("RGB", (1600, 1800), "#f3f5f7")
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 38)
    for i, text in enumerate(
        (
            "ТҮНДҮК / ЦИФРОВОЕ / KYRGYZ REPUBLIC",
            "ФАМИЛИЯ: СИНТЕТИКОВ",
            "ИМЯ: ТЕСТ",
            "ОТЧЕСТВО: ПРИМЕРОВИЧ",
            "ДАТА РОЖДЕНИЯ: 07.11.1990",
            "ПИН: 00123456789012",
            "НОМЕР ДОКУМЕНТА: ID0000001",
            "ДАТА ВЫДАЧИ: 22.11.2024",
            "СРОК ДЕЙСТВИЯ: 22.11.2034",
            "ОРГАН ВЫДАЧИ: ТЕСТ ОРГАН",
            "ПОЛ: МУЖСКОЙ",
            "ГРАЖДАНСТВО: КЫРГЫЗ РЕСПУБЛИКАСЫ",
        )
    ):
        draw.text((65, 60 + i * 125), text, font=font, fill="black")
    digital = directory / "synthetic-digital.png"
    image.save(digital)
    documents.append(BenchmarkDocument("digital-screenshot", [digital], expected))
    return documents


def rss():
    total = 0
    for process in [psutil.Process(), *psutil.Process().children(recursive=True)]:
        try:
            total += process.memory_info().rss
        except psutil.Error:
            pass
    return total


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--models", type=Path)
    parser.add_argument("--tesseract", type=Path, help="Explicit optional Tesseract executable")
    parser.add_argument(
        "--engines", nargs="+", choices=("paddle", "tesseract"), default=["tesseract", "paddle"]
    )
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--authorized-local-documents", action="store_true")
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    if args.manifest:
        if not args.authorized_local_documents:
            parser.error("A private manifest requires explicit --authorized-local-documents consent")
        records = json.loads(args.manifest.read_text(encoding="utf-8"))
        documents = [
            BenchmarkDocument(
                r["id"], [(args.manifest.parent / p).resolve() for p in r["images"]], r["expected"]
            )
            for r in records
        ]
    else:
        documents = synthetic_documents(args.output / "synthetic")
    reports = []
    for engine in args.engines:
        _POOL.close()
        settings = Settings(
            ocr_engine=engine,
            languages=["rus", "eng", "kir"],
            timeout=120,
            paddle_model_directory=str(args.models) if args.models else "",
            tesseract_path=str(args.tesseract.resolve()) if args.tesseract else "",
        )
        print("Benchmark:", engine, flush=True)
        reports.append(benchmark(documents, settings, rss))
        _POOL.close()
    data = {
        "synthetic_only": not bool(args.manifest),
        "engines": reports,
        "dangerous_error_definition": "Wrong nonempty resolved value without an INVALID/CONFLICT status; human review is still required before export.",
        "memory_definition": "Peak sampled RSS of benchmark process plus all child processes, includes inference/model load, sampled every 20ms.",
        "limitations": "Synthetic fixtures do not establish accuracy on real Kyrgyz identity documents. No automatic field or document confirmation.",
    }
    (args.output / "benchmark.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(
        json.dumps(
            {
                r["engine"]: {k: v for k, v in r.items() if k not in {"documents", "field_accuracy"}}
                for r in reports
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    os.environ.setdefault("PYTHONUTF8", "1")
    main()
