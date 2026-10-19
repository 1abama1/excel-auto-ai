"""Private physical controls, masked results; no copied photographs or values."""

import argparse
import json
import re
import threading
from pathlib import Path

from benchmark_cpu import control_names

from document_excel.application.workflows import RecognitionService
from document_excel.domain.models import ClientCase, SourceImage, parse_date, split_mkk_authority
from document_excel.infrastructure.extraction import apply_candidates
from document_excel.infrastructure.paddle_ocr import _POOL
from document_excel.infrastructure.storage import Settings


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--images", type=Path, nargs=2, required=True)
    parser.add_argument("--specification", type=Path, required=True)
    parser.add_argument("--models", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    text = args.specification.read_text(encoding="utf8")
    expected = control_names(args.specification)
    for field, label in (
        ("birth_date", "Дата рождения"),
        ("passport_issued_date", "Дата выдачи"),
        ("expiration_date", "Срок действия"),
        ("document_number", "Номер документа"),
        ("passport_issued_by", "Орган выдачи"),
    ):
        match = re.search(r"\|\s*" + label + r"\s*\|\s*([^|\r\n]+)", text)
        if match:
            value = match[1].strip().split(" —")[0]
            expected[field] = str(parse_date(value)) if field.endswith("date") else value
    authority = split_mkk_authority(expected.get("passport_issued_by", ""))
    if authority:
        expected["passport_issued_by"], expected["passport_department_code"] = authority
    reports = {}
    try:
        for engine in ("rapidocr", "paddle"):
            _POOL.close()
            case = ClientCase()
            case.images = [SourceImage(path=p, case_id=case.id) for p in args.images]
            settings = Settings(
                ocr_engine=engine,
                rapid_model_directory=str((args.models / "rapidocr").resolve()),
                paddle_model_directory=str((args.models / "paddle").resolve()),
                timeout=120,
            )
            output = RecognitionService(settings).run(case.images, threading.Event(), lambda *_: None)
            apply_candidates(case, output.candidates, output.document_type, output.label_seen)
            checks = {field: str(getattr(case.data, field)) == value for field, value in expected.items()}
            reports[engine] = {
                "checks": checks,
                "mrz_valid": any(r.valid for r in output.mrz_reports.values()),
                "registration_not_present": str(case.registration_status) == "NOT_PRESENT",
                "profile": output.document_type,
                "seconds": output.elapsed,
                "pin_present": bool(case.data.personal_identifier),
            }
            assert all(checks.values()) and reports[engine]["mrz_valid"], {"engine": engine, "checks": checks}
    finally:
        _POOL.close()
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(reports, indent=2), encoding="utf8")
    print(json.dumps(reports))


if __name__ == "__main__":
    main()
