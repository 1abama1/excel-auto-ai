"""Field-level, local comparisons. Reports contain metrics, never OCR values."""

from __future__ import annotations

import threading
import time
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from document_excel.application.workflows import RecognitionService
from document_excel.domain.models import ClientCase, DocumentSide, FieldStatus, SourceImage
from document_excel.infrastructure.extraction import apply_candidates
from document_excel.infrastructure.ocr_factory import create_engine
from document_excel.infrastructure.storage import Settings

FIELDS = (
    "full_name",
    "birth_date",
    "personal_identifier",
    "document_number",
    "passport_issued_date",
    "expiration_date",
    "passport_issued_by",
    "gender",
    "citizenship",
)


@dataclass
class BenchmarkDocument:
    id: str
    images: list[Path]
    expected: dict[str, str]
    document_type: str | None = None


def evaluate(case: ClientCase, expected: dict[str, str]) -> dict:
    comparisons = {}
    dangerous = 0
    for name, truth in expected.items():
        if name not in FIELDS:
            raise ValueError("Unknown benchmark field")
        value = getattr(case.data, name)
        observed = str(value) if value is not None else ""
        evidence = case.evidence.get("passport_number" if name == "document_number" else name)
        unresolved = bool(evidence and evidence.status in {FieldStatus.CONFLICT, FieldStatus.INVALID})
        correct = observed == truth and not unresolved
        comparisons[name] = correct
        # A plausible but wrong populated value is dangerous even when the
        # application still requires human confirmation before Excel export.
        dangerous += bool(observed and not correct and not unresolved)
    return {"fields": comparisons, "fully_correct": all(comparisons.values()), "dangerous_errors": dangerous}


def benchmark(documents: list[BenchmarkDocument], settings: Settings, memory: Callable[[], int]) -> dict:
    engine = create_engine(settings)
    diagnostic = engine.diagnose()
    if diagnostic.error:
        return {
            "engine": settings.ocr_engine,
            "status": "unavailable",
            "documents": [],
            "reason": diagnostic.error,
        }
    service = RecognitionService(settings, engine)
    results = []
    for document in documents:
        sources = [SourceImage(path=path, role=DocumentSide.UNKNOWN) for path in document.images]
        case = ClientCase(images=sources)
        peak = [memory()]
        stop = threading.Event()

        def sample(stop=stop, peak=peak):
            while not stop.wait(0.02):
                peak[0] = max(peak[0], memory())

        sampler = threading.Thread(target=sample, daemon=True)
        sampler.start()
        started = time.perf_counter()
        try:
            output = service.run(
                sources, threading.Event(), lambda *_: None, document_type=document.document_type
            )
            apply_candidates(case, output.candidates, output.document_type, output.label_seen)
            result = evaluate(case, document.expected)
            result["mrz_valid"] = any(r.valid for r in output.mrz_reports.values())
            result["status"] = "completed"
        except Exception as exc:
            result = {
                "status": "failed",
                "error_type": type(exc).__name__,
                "fields": dict.fromkeys(document.expected, False),
                "fully_correct": False,
                "dangerous_errors": 0,
            }
        finally:
            stop.set()
            sampler.join(timeout=1)
        results.append(
            {
                "id": document.id,
                **result,
                "seconds": round(time.perf_counter() - started, 3),
                "peak_process_tree_rss_mb": round(peak[0] / 1024**2, 1),
            }
        )
    fields = {}
    for name in FIELDS:
        checks = [r["fields"][name] for r in results if name in r["fields"]]
        fields[name] = {
            "correct": sum(checks),
            "total": len(checks),
            "accuracy": sum(checks) / len(checks) if checks else None,
        }
    return {
        "engine": settings.ocr_engine,
        "version": diagnostic.version,
        "status": "completed",
        "documents": results,
        "field_accuracy": fields,
        "fully_correct_documents": sum(r["fully_correct"] for r in results),
        "total_documents": len(results),
        "dangerous_errors": sum(r["dangerous_errors"] for r in results),
    }
