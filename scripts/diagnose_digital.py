"""Local-only stage audit. Writes counts/geometry/timing, never OCR values."""

from __future__ import annotations

import argparse
import json
import re
import sys
import threading
import time
from collections import Counter
from pathlib import Path

import psutil

from document_excel.application import fast_recognition as fast
from document_excel.application.ocr_state import OcrState
from document_excel.application.workflows import RecognitionService
from document_excel.domain.models import ClientCase, FieldStatus, SourceImage
from document_excel.infrastructure.extraction import (
    FieldExtractor,
    apply_candidates,
    classify,
    load_document_profiles,
)
from document_excel.infrastructure.paddle_ocr import _POOL
from document_excel.infrastructure.preprocessing import (
    assess_quality,
    crop_document_region,
    detect_document_region,
    load_image,
)
from document_excel.infrastructure.storage import Settings

FIELDS = (
    "full_name",
    "surname_cyrillic",
    "surname_latin",
    "first_name_cyrillic",
    "first_name_latin",
    "patronymic_cyrillic",
    "birth_date",
    "personal_identifier",
    "document_number",
    "passport_issued_by",
    "passport_department_code",
    "passport_issued_date",
    "expiration_date",
    "registration_address",
    "gender",
    "citizenship",
)


class MemorySample:
    """Observed working set of this driver and its local OCR subprocess."""

    def __init__(self):
        self.peak = 0
        self.stop = threading.Event()
        self.thread = threading.Thread(target=self.sample, daemon=True)

    def sample(self):
        while not self.stop.is_set():
            root = psutil.Process()
            total = 0
            for proc in [root, *root.children(recursive=True)]:
                try:
                    total += proc.memory_info().rss
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            self.peak = max(self.peak, total)
            self.stop.wait(0.05)

    def __enter__(self):
        self.thread.start()
        return self

    def __exit__(self, *_):
        self.stop.set()
        self.thread.join()


def presence(case):
    return {
        name: bool(value.display() if hasattr(value, "display") else value)
        for name in FIELDS
        if (value := getattr(case.data, name, None)) is not None
    } | {name: False for name in FIELDS if getattr(case.data, name, None) is None}


def run(args):
    reports = {}
    expected = json.loads(sys.stdin.readline()) if args.expected_stdin else {}
    original_extract = FieldExtractor.extract
    extraction = []

    def extract(self, result, *a, **kw):
        started = time.perf_counter()
        candidates = original_extract(self, result, *a, **kw)
        extraction.append(
            {
                "profile": self.profile.id if self.profile else None,
                "input_blocks": len(result.blocks),
                "candidates": len(candidates),
                "by_field": dict(Counter(c.field_name for c in candidates)),
                "rejected": sum(c.validation_status == FieldStatus.INVALID for c in candidates),
                "seconds": time.perf_counter() - started,
            }
        )
        return candidates

    FieldExtractor.extract = extract
    try:
        for backend in args.engines:
            _POOL.close()
            fast.clear_cache()
            service = RecognitionService(
                Settings(
                    ocr_engine=backend,
                    timeout=120,
                    rapid_model_directory=str(args.rapid_models.resolve()),
                    paddle_model_directory=str(args.paddle_models.resolve()),
                )
            )
            reads = []
            native = service.engine._read

            def read(image, image_id, variant, cancel, mode="document", native=native, reads=reads):
                start = time.perf_counter()
                result = native(image, image_id, variant, cancel, mode)
                reads.append(
                    {
                        "mode": mode,
                        "variant": variant,
                        "size": list(image.shape[:2][::-1]),
                        "blocks": len(result.blocks),
                        "seconds": time.perf_counter() - start,
                        "valid_boxes": sum(b.bbox[2] > 0 and b.bbox[3] > 0 for b in result.blocks),
                        "median_row_height": sorted(b.bbox[3] for b in result.blocks)[len(result.blocks) // 2]
                        if result.blocks
                        else 0,
                    }
                )
                return result

            service.engine._read = read
            cases = []
            modes = [None if mode == "auto" else "kg_digital_id" for mode in args.workflow_modes]
            for run_index, forced in enumerate([mode for mode in modes for _ in range(args.repeats)]):
                case = ClientCase()
                case.images = [SourceImage(path=p, case_id=case.id) for p in args.images]
                if forced:
                    case.data.document_type = forced
                    case.document_type_manual = True
                state = OcrState()
                ticket = state.begin(case)
                extraction.clear()
                reads.clear()
                with MemorySample() as memory:
                    output = service.run(
                        case.images, threading.Event(), lambda *_: None, document_type=forced
                    )
                resolved = case.model_copy(deep=True)
                apply_candidates(resolved, output.candidates, output.document_type, output.label_seen)
                committed = state.commit(case, ticket.id, output)
                preview = state.preview(case) if hasattr(state, "preview") else None
                cases.append(
                    {
                        "forced": forced,
                        "repeat": run_index,
                        "seconds": output.elapsed,
                        "peak_driver_and_worker_rss_mib": round(memory.peak / 1048576, 1),
                        "document_type": output.document_type,
                        "sides": [str(output.sides.get(i.id)) for i in case.images],
                        "reads": list(reads),
                        "extraction": list(extraction),
                        "candidates": len(output.candidates),
                        "after_resolution": presence(resolved),
                        "exact_checks": {
                            field: (
                                getattr(resolved.data, field).raw
                                if field == "registration_address"
                                else str(getattr(resolved.data, field))
                            )
                            == value
                            for field, value in expected.items()
                        },
                        "after_commit": presence(case),
                        "address_candidate_checks": [
                            {
                                "method": c.extraction_method,
                                "matches": re.sub(r"\W", "", c.normalized_value).casefold()
                                == re.sub(r"\W", "", expected["registration_address"]).casefold(),
                            }
                            for c in output.candidates
                            if c.field_name == "registration_address"
                        ]
                        if "registration_address" in expected
                        else [],
                        "preview_fields": presence(preview) if preview else {},
                        "processing": getattr(output, "processing_metrics", {}),
                        "text_checks": {
                            "registration_address": re.sub(
                                r"\W", "", resolved.data.registration_address.raw
                            ).casefold()
                            == re.sub(r"\W", "", expected["registration_address"]).casefold()
                        }
                        if "registration_address" in expected
                        else {},
                        "commit_accepted": committed,
                        "identity_status": case.identity_status,
                        "pending_identity": bool(case.pending_identity),
                        "evidence_status": {n: str(e.status) for n, e in resolved.evidence.items()},
                    }
                )
                if args.confirm_identity_for_test and case.pending_identity:
                    cases[-1]["identity_confirmation_accepted"] = state.accept_identity(case)
                    cases[-1]["after_identity_confirmation"] = presence(case)
                if args.check_cache:
                    count = len(reads)
                    cache_started = time.perf_counter()
                    cached = service.run(
                        case.images, threading.Event(), lambda *_: None, document_type=forced
                    )
                    cases[-1]["cached_seconds"] = time.perf_counter() - cache_started
                    cases[-1]["cached_new_ocr_requests"] = len(reads) - count
                    assert len(reads) == count and cached.candidates == output.candidates
            probes = []
            if args.probes:
                for index, path in enumerate(args.images):
                    original = load_image(path)
                    region = detect_document_region(original)
                    plastic = fast.prepare_plastic_card(original)
                    quality = assess_quality(original)
                    dimensions = {
                        "original": list(original.shape[:2][::-1]),
                        "plastic": list(plastic.image.shape[:2][::-1]),
                        "plastic_detected": plastic.detected,
                        "document_roi": region.bbox,
                        "roi_detected": region.detected,
                        "blur": quality.blur,
                        "contrast": quality.contrast,
                    }
                    samples = [
                        ("original", original),
                        ("existing_card_roi", crop_document_region(original, region)),
                    ]
                    image_id = case.images[index].id
                    hints = getattr(output, "field_regions", {})
                    for label, names in (
                        ("names_roi", ("surname", "first_name", "patronymic")),
                        ("birth_roi", ("birth_date",)),
                        ("pin_roi", ("personal_identifier",)),
                        ("document_number_roi", ("passport_number",)),
                    ):
                        boxes = [
                            hints[n]["bbox"] for n in names if n in hints and hints[n]["image_id"] == image_id
                        ]
                        if boxes:
                            left, top = min(b[0] for b in boxes), min(b[1] for b in boxes)
                            right, bottom = max(b[0] + b[2] for b in boxes), max(b[1] + b[3] for b in boxes)
                            samples.append((label, original[top:bottom, left:right]))
                    for variant, pixels in samples:
                        start = time.perf_counter()
                        result = service.engine.recognize(
                            pixels, f"probe-{index}", variant, threading.Event()
                        )
                        candidates = original_extract(FieldExtractor(), result)
                        probes.append(
                            {
                                "image": index,
                                "variant": variant,
                                "dimensions": dimensions,
                                "seconds": time.perf_counter() - start,
                                "blocks": len(result.blocks),
                                "document_type": classify(result, load_document_profiles()),
                                "candidates": dict(Counter(c.field_name for c in candidates)),
                            }
                        )
            reports[backend] = {"workflow": cases, "probes": probes}
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(json.dumps(reports, ensure_ascii=False, indent=2), encoding="utf-8")
            print(
                json.dumps(
                    {
                        "engine": backend,
                        "summary": [
                            {
                                "profile": r["document_type"],
                                "candidates": r["candidates"],
                                "resolved_fields": sum(r["after_resolution"].values()),
                                "committed_fields": sum(r["after_commit"].values()),
                                "commit_accepted": r["commit_accepted"],
                                "identity_status": r["identity_status"],
                                "pending_identity": r["pending_identity"],
                            }
                            for r in cases
                        ],
                    }
                ),
                flush=True,
            )
    finally:
        FieldExtractor.extract = original_extract
        _POOL.close()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(reports, ensure_ascii=False, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                engine: [
                    {
                        "profile": r["document_type"],
                        "blocks": sum(x["blocks"] for x in r["reads"]),
                        "candidates": r["candidates"],
                        "resolved_fields": sum(r["after_resolution"].values()),
                        "committed_fields": sum(r["after_commit"].values()),
                        "identity_status": r["identity_status"],
                    }
                    for r in report["workflow"]
                ]
                for engine, report in reports.items()
            }
        )
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--images", type=Path, nargs=2, required=True)
    parser.add_argument("--rapid-models", type=Path, required=True)
    parser.add_argument("--paddle-models", type=Path, required=True)
    parser.add_argument(
        "--engines", nargs="+", default=["rapidocr", "paddle"], choices=["rapidocr", "paddle"]
    )
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--probes", action="store_true")
    parser.add_argument(
        "--expected-stdin",
        action="store_true",
        help="Read private controls once from stdin; store comparisons only",
    )
    parser.add_argument(
        "--workflow-modes", nargs="+", choices=["auto", "digital"], default=["auto", "digital"]
    )
    parser.add_argument("--repeats", type=int, default=1)
    parser.add_argument("--confirm-identity-for-test", action="store_true")
    parser.add_argument("--check-cache", action="store_true")
    run(parser.parse_args())
