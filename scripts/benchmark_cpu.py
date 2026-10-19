"""CPU comparison: identical workflow, isolated clients, exact synthetic controls.

Only booleans, timings, counters and memory are written. Real control values
remain in memory. A baseline agreement is explicitly distinct from ground truth.
"""

from __future__ import annotations

import argparse
import importlib.metadata
import json
import os
import platform
import re
import sys
import threading
import time
from pathlib import Path

import psutil


def audited_serve(*args):
    from document_excel.infrastructure import paddle_ocr as module

    create = module._create_backend
    creations = [0]

    def factory(*a):
        attempts = [0]

        def denied(*_a, **_kw):
            attempts[0] += 1
            raise OSError("Offline CPU benchmark")

        module.socket.socket.connect = denied
        module.socket.create_connection = denied
        started = time.perf_counter()
        backend = create(*a)
        creations[0] += 1
        initialization = time.perf_counter() - started
        native = backend.recognize

        def read(*a, **kw):
            result = native(*a, **kw)
            providers = ["Paddle CPU"]
            compiled_cuda = False
            if hasattr(backend, "detector"):
                import onnxruntime

                providers = backend.detector.session.session.get_providers()
                compiled_cuda = "CUDAExecutionProvider" in onnxruntime.get_available_providers()
            else:
                import paddle

                compiled_cuda = paddle.is_compiled_with_cuda()
            result.cpu_benchmark = {
                "backend_initialization_seconds": initialization,
                "model_instances": 1 + len(backend.recognizers),
                "backend_creations": creations[0],
                "providers": providers,
                "compiled_with_cuda": compiled_cuda,
                "network_attempts": attempts[0],
                "paddle_imported": "paddle" in sys.modules,
            }
            return result

        backend.recognize = read
        return backend

    module._create_backend = factory
    module._serve(*args)


def control_names(path):
    if not path:
        return {}
    text = path.read_text(encoding="utf-8")
    output = {}
    for field, heading in (("surname", "Фамилия"), ("first_name", "Имя")):
        group = re.search(r"\*\*" + heading + r":\*\*(.*?)(?=\*\*|\Z)", text, re.S)[1]
        for script, label in (("cyrillic", "Кириллица"), ("latin", "Латиница")):
            output[field + "_" + script] = re.search(label + r":\s*([^\r\n]+)", group)[1].strip()
    output["patronymic_cyrillic"] = re.search(r"\*\*Отчество:\*\*\s*-\s*([^\r\n]+)", text)[1].strip()
    return output


def synthetic_cases(directory):
    from document_excel.infrastructure.physical_verification import synthetic_plastic_images

    common = {
        "passport_number": "0000001",
        "personal_identifier": "00123456789012",
        "birth_date": "1990-11-07",
        "passport_issued_date": "2024-11-22",
        "expiration_date": "2034-11-22",
        "passport_issued_by": "МКК 012345",
    }
    variants = [
        ("short", "ЛИ", "LI", "ЯН", "YAN", "ПРИМЕРОВИЧ"),
        ("long", "КОНСТАНТИНОПОЛЬСКИЙ", "KONSTANTINOPOLSKIY", "АЛЕКСАНДРА", "ALEKSANDRA", "СЕРГЕЕВНА"),
        ("double", "ТЕСТОВА", "TESTOVA", "АННА-МАРИЯ", "ANNA-MARIA", "ПРИМЕРОВНА"),
        ("no_patronymic", "СИНТЕТИКОВ", "SINTETIKOV", "ПРИМЕР", "PRIMER", None),
    ]
    cases = []
    for index, (name, surname, surname_latin, first, first_latin, patronymic) in enumerate(variants):
        controls = common | {
            "surname_cyrillic": surname,
            "surname_latin": surname_latin,
            "first_name_cyrillic": first,
            "first_name_latin": first_latin,
            "patronymic_cyrillic": patronymic,
        }
        if index % 2:
            controls |= {
                "passport_number": "1234567",
                "personal_identifier": "00987654321098",
                "birth_date": "1987-03-15",
                "passport_issued_date": "2023-06-10",
                "expiration_date": "2033-06-10",
            }
        paths = synthetic_plastic_images(directory / name, controls)
        # The application deliberately separates legacy MKK authority/office
        # code. Core accuracy below is exact FIO, PIN, number and full dates.
        controls.pop("passport_issued_by")
        cases.append((name, paths, controls))
    return cases


def measured_run(service, paths, identity, expected, baseline):
    from document_excel.domain.models import ClientCase, SourceImage
    from document_excel.infrastructure.extraction import apply_candidates
    from document_excel.infrastructure.paddle_ocr import _POOL

    sources = [SourceImage(path=p, case_id=identity) for p in paths]
    reads, telemetry = [], []
    native = service.engine._read

    def read(*a, **kw):
        result = native(*a, **kw)
        reads.append(a[4] if len(a) > 4 else kw.get("mode", "document"))
        telemetry.append(getattr(result, "cpu_benchmark", {}))
        return result

    service.engine._read = read
    peak, cpu = [0], [0.0]
    stop = threading.Event()
    initial = sum(
        p.cpu_times().user + p.cpu_times().system
        for p in [psutil.Process(), *psutil.Process().children(recursive=True)]
    )

    def sample():
        while not stop.wait(0.01):
            rss, seconds = 0, 0.0
            for p in [psutil.Process(), *psutil.Process().children(recursive=True)]:
                try:
                    rss += p.memory_info().rss
                    t = p.cpu_times()
                    seconds += t.user + t.system
                except psutil.Error:
                    pass
            peak[0] = max(peak[0], rss)
            cpu[0] = max(cpu[0], seconds)

    thread = threading.Thread(target=sample, daemon=True)
    thread.start()
    started = time.perf_counter()
    try:
        output = service.run(sources, threading.Event(), lambda *_: None)
        case = ClientCase(images=sources, mrz_reports=output.mrz_reports)
        apply_candidates(case, output.candidates, output.document_type)
        elapsed = time.perf_counter() - started
    finally:
        stop.set()
        thread.join()
        service.engine._read = native
    observed = {
        k: str(getattr(case.data, k)) if getattr(case.data, k) is not None else None for k in expected
    }
    exact = {k: observed[k] == (str(v) if v is not None else None) for k, v in expected.items()}
    checked_mrz = next((r for r in output.mrz_reports.values() if r.valid), None)
    consistency = {}
    if checked_mrz:
        consistency["personal_identifier"] = case.data.personal_identifier == checked_mrz.values.get(
            "personal_identifier"
        )
        consistency["passport_number"] = case.data.document_number == checked_mrz.values.get(
            "document_number"
        )
        for field in ("birth_date", "expiration_date"):
            date = getattr(case.data, field)
            consistency[field] = bool(date and date.strftime("%y%m%d") == checked_mrz.values.get(field))
    field_names = (
        "surname_cyrillic",
        "surname_latin",
        "first_name_cyrillic",
        "first_name_latin",
        "patronymic_cyrillic",
        "personal_identifier",
        "passport_number",
        "birth_date",
        "passport_issued_date",
        "expiration_date",
        "passport_issued_by",
    )
    current = {
        k: str(getattr(case.data, k)) if getattr(case.data, k) is not None else None for k in field_names
    }
    agreement = {k: current[k] == v for k, v in baseline.items()} if baseline else {}
    count = len(reads)
    started = time.perf_counter()
    service.engine._read = read
    try:
        cached = service.run(sources, threading.Event(), lambda *_: None)
    finally:
        service.engine._read = native
    cached_seconds = time.perf_counter() - started
    # No inference on an unchanged client's reopen; models stay in the worker.
    assert len(reads) == count and cached.candidates == output.candidates
    record = {
        "seconds": elapsed,
        "peak_tree_rss_mib": peak[0] / 1024**2,
        "cpu_seconds": max(0.0, cpu[0] - initial),
        "ocr_calls": count,
        "document_passes": reads.count("document"),
        "cached_seconds": cached_seconds,
        "exact_checks": exact,
        "mrz_valid": bool(checked_mrz),
        "mrz_consistency": consistency,
        "baseline_agreement_not_ground_truth": agreement,
        "backend": telemetry[-1] if telemetry else {},
        "worker_pid": _POOL.process.pid,
        "sides": [str(v) for v in output.sides.values()],
        "field_statuses": {
            k: str(case.evidence[k].status) if k in case.evidence else "missing" for k in field_names
        },
    }
    return record, current


def main():
    from document_excel.application.fast_recognition import clear_cache
    from document_excel.application.workflows import RecognitionService
    from document_excel.infrastructure import paddle_ocr
    from document_excel.infrastructure.storage import Settings

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--paddle-models", type=Path, required=True)
    parser.add_argument("--rapid-models", type=Path, required=True)
    parser.add_argument("--images", type=Path, nargs=2)
    parser.add_argument("--expected-request", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--threads", type=int, default=2)
    parser.add_argument("--real-repeats", type=int, default=3)
    args = parser.parse_args()
    paddle_ocr._serve = audited_serve
    datasets = synthetic_cases(args.output.parent / "synthetic-cpu")
    if args.images:
        datasets.insert(0, ("private", args.images, control_names(args.expected_request)))
    references, reports = {}, {}
    try:
        for backend in ("paddle", "rapidocr"):
            paddle_ocr._POOL.close()
            clear_cache()
            settings = Settings(
                ocr_engine=backend,
                paddle_model_directory=str(args.paddle_models.resolve()),
                rapid_model_directory=str(args.rapid_models.resolve()),
                paddle_cpu_threads=args.threads,
                paddle_device="cpu",
                timeout=120,
            )
            service = RecognitionService(settings)
            records = []
            for name, paths, expected in datasets:
                repeats = args.real_repeats if name == "private" else 1
                for repeat in range(repeats):
                    record, values = measured_run(
                        service, paths, f"{backend}-{name}-{repeat}", expected, references.get(name)
                    )
                    record.update({"sample": name, "repeat": repeat, "cold_models": not records})
                    records.append(record)
                    if backend == "paddle" and repeat == 0:
                        references[name] = values
                    print(
                        json.dumps(
                            {
                                "engine": backend,
                                "sample": name,
                                "repeat": repeat,
                                "seconds": round(record["seconds"], 3),
                                "exact_passed": sum(record["exact_checks"].values()),
                                "exact_total": len(record["exact_checks"]),
                                "mrz_valid": record["mrz_valid"],
                                "calls": record["ocr_calls"],
                            }
                        ),
                        flush=True,
                    )
            reports[backend] = records
    finally:
        paddle_ocr._POOL.close()
    gpu_packages = sorted(
        d.metadata["Name"]
        for d in importlib.metadata.distributions()
        if d.metadata["Name"]
        and (
            d.metadata["Name"].lower().startswith("nvidia-")
            or d.metadata["Name"].lower() in {"paddlepaddle-gpu", "onnxruntime-gpu", "tensorrt"}
        )
    )
    result = {
        "hardware": {
            "processor": platform.processor(),
            "logical_cpus": os.cpu_count(),
            "windows": platform.version(),
        },
        "threads": args.threads,
        "gpu_packages": gpu_packages,
        "reports": reports,
        "privacy": "No OCR text, paths or control values stored. Baseline agreement is not ground truth.",
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")


if __name__ == "__main__":
    main()
