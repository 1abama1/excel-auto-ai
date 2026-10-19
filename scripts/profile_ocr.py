"""Local opt-in performance audit. Reports never include OCR text or paths."""

from __future__ import annotations

import argparse
import functools
import json
import os
import re
import threading
import time
from pathlib import Path

import psutil

STATS: dict[str, dict] = {}
DEPTH: dict[str, int] = {}


def add(name, seconds):
    item = STATS.setdefault(name, {"seconds": 0.0, "calls": 0})
    item["seconds"] += seconds
    item["calls"] += 1


def wrap(owner, name, stage, iterator=False):
    original = getattr(owner, name)
    if iterator:

        @functools.wraps(original)
        def measured(*args, **kwargs):
            outer = not DEPTH.get(stage, 0)
            DEPTH[stage] = DEPTH.get(stage, 0) + 1
            started = time.perf_counter()
            try:
                values = list(original(*args, **kwargs))
            finally:
                DEPTH[stage] -= 1
                if outer:
                    add(stage, time.perf_counter() - started)
            yield from values
    else:

        @functools.wraps(original)
        def measured(*args, **kwargs):
            started = time.perf_counter()
            try:
                return original(*args, **kwargs)
            finally:
                add(stage, time.perf_counter() - started)

    setattr(owner, name, measured)


def profiled_serve(*args):
    from document_excel.infrastructure import paddle_ocr as module

    original = module._Backend

    class ProfiledBackend(original):
        def __init__(self, *a, **kw):
            started = time.perf_counter()
            from paddleocr import TextRecognition

            if os.environ.get("OCR_PROFILE_MKLDNN") == "1":
                import paddleocr

                native = paddleocr.PaddleOCR

                def accelerated(*args, **kwargs):
                    kwargs["enable_mkldnn"] = True
                    return native(*args, **kwargs)

                paddleocr.PaddleOCR = accelerated

                def accelerated_rec(*args, **kwargs):
                    kwargs["enable_mkldnn"] = True
                    return TextRecognition(*args, **kwargs)

                paddleocr.TextRecognition = accelerated_rec

            add("paddle_import", time.perf_counter() - started)
            started = time.perf_counter()
            super().__init__(*a, **kw)
            add("initial_models", time.perf_counter() - started)
            pipeline = self.pipeline.paddlex_pipeline
            wrap(type(pipeline.text_det_model), "__call__", "text_detection", True)
            wrap(type(pipeline.text_rec_model), "__call__", "text_recognition", True)
            wrap(pipeline.text_rec_model, "predict", "text_recognition", True)

            def create(*a, **kw):
                started = time.perf_counter()
                rec = TextRecognition(*a, **kw)
                add("additional_models", time.perf_counter() - started)
                wrap(rec, "predict", "text_recognition", True)
                return rec

            self.recognition_class = create
            self.unreported = True

        def recognize(self, *a, **kw):
            before = {} if self.unreported else {k: dict(v) for k, v in STATS.items()}
            self.unreported = False
            result = super().recognize(*a, **kw)
            result.telemetry = {
                k: {
                    "seconds": v["seconds"] - before.get(k, {}).get("seconds", 0),
                    "calls": v["calls"] - before.get(k, {}).get("calls", 0),
                }
                for k, v in STATS.items()
            }
            return result

    module._Backend = ProfiledBackend
    module._serve(*args)


def main():
    from document_excel.application import workflows
    from document_excel.domain.models import ClientCase, SourceImage
    from document_excel.infrastructure import extraction, mrz, paddle_ocr, physical_id, preprocessing
    from document_excel.infrastructure.storage import Settings

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--images", nargs="+", type=Path, required=True)
    parser.add_argument("--models", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--expected-request", type=Path)
    parser.add_argument("--repeats", type=int, default=2)
    parser.add_argument("--cpu-threads", type=int, default=2)
    parser.add_argument("--mkldnn", action="store_true")
    parser.add_argument(
        "--geometry", action="store_true", help="Include masked group geometry for local diagnosis"
    )
    parser.add_argument(
        "--isolated-repeats",
        action="store_true",
        help="Warm models with an isolated client; bypass result caches",
    )
    args = parser.parse_args()
    if args.mkldnn:
        os.environ["OCR_PROFILE_MKLDNN"] = "1"
    # Keep the original target callable: the patched child entry calls _serve.
    original_serve = paddle_ocr._serve
    paddle_ocr._serve = profiled_serve
    # In a spawned process imports are fresh, so _serve there is the original.
    assert original_serve is not profiled_serve
    wrap(preprocessing, "load_image", "image_open")
    wrap(preprocessing, "detect_document_region", "document_detection")
    wrap(preprocessing.ImagePreprocessingService, "prepare", "preprocessing_total")
    wrap(physical_id, "prepare_plastic_card", "perspective")
    wrap(extraction.FieldExtractor, "extract", "field_extraction")
    try:
        from document_excel.infrastructure.full_name import FullNameExtractor
    except ImportError:
        pass
    else:
        wrap(FullNameExtractor, "extract", "full_name_extraction")
    wrap(mrz.MRZRecognizer, "recognize", "mrz_total")
    original_read = paddle_ocr.PaddleOCREngine._read

    def read(self, *a, **kw):
        started = time.perf_counter()
        result = original_read(self, *a, **kw)
        entry = {
            "variant": result.variant,
            "width": result.width,
            "height": result.height,
            "contains": {k: v in result.text for k, v in expected.items()},
            "contains_joined": {
                k: v.replace(" ", "") in result.text.replace(" ", "").replace("\n", "")
                for k, v in expected.items()
            },
        }
        if args.geometry and result.variant == "physical-card":
            from document_excel.infrastructure.full_name import FullNameExtractor

            groups = FullNameExtractor().extract(result, workflows.DocumentSide.FRONT)
            entry["name_groups"] = {
                "regions": groups.regions,
                "rows": {
                    n: [
                        {
                            "bbox": r.bbox,
                            "length": len(r.text),
                            "confidence": r.confidence,
                            "matches": {k: v == r.text for k, v in expected.items()},
                        }
                        for r in rows
                    ]
                    for n, rows in groups.rows.items()
                },
            }
            entry["labels"] = [{"field": n, "bbox": b.bbox} for b, n in physical_id.semantic_labels(result)]
        reads.append(entry)
        add("ocr_requests", time.perf_counter() - started)
        for name, metric in getattr(result, "telemetry", {}).items():
            total = STATS.setdefault("runtime." + name, {"seconds": 0.0, "calls": 0})
            total["seconds"] += metric["seconds"]
            total["calls"] += metric["calls"]
        return result

    paddle_ocr.PaddleOCREngine._read = read
    expected = {}
    if args.expected_request:
        text = args.expected_request.read_text(encoding="utf-8")
        for field, heading in (("surname", "Фамилия"), ("first_name", "Имя")):
            group = re.search(r"\*\*" + heading + r":\*\*(.*?)(?=\*\*|\Z)", text, re.S)[1]
            expected[field + "_cyrillic"] = re.search(r"Кириллица:\s*([^\r\n]+)", group)[1].strip()
            expected[field + "_latin"] = re.search(r"Латиница:\s*([^\r\n]+)", group)[1].strip()
        expected["patronymic_cyrillic"] = re.search(r"\*\*Отчество:\*\*\s*-\s*([^\r\n]+)", text)[1].strip()
    settings = Settings(
        ocr_engine="paddle",
        paddle_model_directory=str(args.models.resolve()), timeout=120, paddle_cpu_threads=args.cpu_threads
    )
    sources = [SourceImage(path=p, case_id="local-audit") for p in args.images]
    service = workflows.RecognitionService(settings)
    reports = []
    for index in range(args.repeats):
        if args.isolated_repeats:
            sources = [source.model_copy(update={"case_id": f"local-audit-{index}"}) for source in sources]
        STATS.clear()
        reads = []
        peak = [0]
        cpu = [0.0]
        initial_cpu = 0.0
        for process in [psutil.Process(), *psutil.Process().children(recursive=True)]:
            t = process.cpu_times()
            initial_cpu += t.user + t.system
        stop = threading.Event()

        def sample(stop=stop, peak=peak, cpu=cpu):
            while not stop.wait(0.05):
                rss, seconds = 0, 0.0
                for process in [psutil.Process(), *psutil.Process().children(recursive=True)]:
                    try:
                        rss += process.memory_info().rss
                        t = process.cpu_times()
                        seconds += t.user + t.system
                    except psutil.Error:
                        pass
                peak[0] = max(peak[0], rss)
                cpu[0] = max(cpu[0], seconds)

        sampler = threading.Thread(target=sample, daemon=True)
        sampler.start()
        started = time.perf_counter()
        try:
            output = service.run(sources, threading.Event(), lambda *_: None)
        except Exception as exc:
            stop.set()
            sampler.join()
            paddle_ocr._POOL.close()
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(
                json.dumps(
                    {
                        "error_type": type(exc).__name__,
                        "run": index,
                        "seconds": time.perf_counter() - started,
                        "mkldnn": args.mkldnn,
                        "threads": args.cpu_threads,
                    }
                ),
                encoding="utf-8",
            )
            raise
        case = ClientCase(images=sources, mrz_reports=output.mrz_reports)
        validation = time.perf_counter()
        extraction.apply_candidates(case, output.candidates, output.document_type, output.label_seen)
        add("field_validation", time.perf_counter() - validation)
        elapsed = time.perf_counter() - started
        stop.set()
        sampler.join()
        names = {}
        for name in ("surname", "first_name", "patronymic"):
            names[name] = [
                {
                    "length": len(c.normalized_value),
                    "side": str(c.source_side),
                    "method": c.extraction_method,
                    "bbox": c.source_bbox,
                    "matches_cyrillic": c.normalized_value == expected.get(name + "_cyrillic"),
                    "matches_latin": c.normalized_value == expected.get(name + "_latin"),
                }
                for c in output.candidates
                if c.field_name == name
            ]
        reports.append(
            {
                "run": index,
                "seconds": elapsed,
                "peak_rss_mb": peak[0] / 1024**2,
                "cpu_seconds_cumulative": cpu[0],
                "cpu_seconds": max(0.0, cpu[0] - initial_cpu),
                "cpu_core_equivalent": max(0.0, cpu[0] - initial_cpu) / elapsed,
                "stages": {k: dict(v) for k, v in STATS.items()},
                "name_checks": {k: getattr(case.data, k) == v for k, v in expected.items()},
                "fields": {
                    k: {"present": bool(getattr(case.data, k, None)), "status": str(e.status)}
                    for k, e in case.evidence.items()
                    if k in type(case.data).model_fields
                },
                "names_metadata": names,
                "reads": reads,
                "mrz_valid": any(r.valid for r in output.mrz_reports.values()),
            }
        )
        print(
            json.dumps(
                {
                    "run": index,
                    "seconds": round(elapsed, 3),
                    "ocr_requests": STATS.get("ocr_requests"),
                    "name_checks": reports[-1]["name_checks"],
                }
            ),
            flush=True,
        )
    paddle_ocr._POOL.close()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(
            {"reports": reports, "gpu": "CPU runtime; GPU not selected"}, ensure_ascii=False, indent=2
        ),
        encoding="utf-8",
    )


if __name__ == "__main__":
    os.environ.setdefault("PYTHONUTF8", "1")
    main()
