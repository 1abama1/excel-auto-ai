"""Local private-photo check: only field/status metadata leaves process memory."""

import argparse
import json
import threading
from pathlib import Path

from document_excel.application.workflows import RecognitionService
from document_excel.domain.models import ClientCase, SourceImage
from document_excel.infrastructure.extraction import apply_candidates
from document_excel.infrastructure.storage import Settings


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("directory", type=Path)
    args = parser.parse_args()
    sources = [SourceImage(path=args.directory / f"{index}.png") for index in (1, 2)]
    result = RecognitionService(Settings()).run(sources, threading.Event(), lambda *_: None)
    case = ClientCase(images=sources, mrz_reports=result.mrz_reports)
    for source in sources:
        source.role = result.sides.get(source.id, source.role)
    apply_candidates(case, result.candidates, result.document_type, result.label_seen)
    report = {
        "document_type": result.document_type,
        "sides": [source.role.value for source in sources],
        "registration_status": case.registration_status.value,
        "fields": {
            name: {
                "status": evidence.status.value,
                "candidates": len(evidence.candidates),
                "methods": sorted({c.source_method for c in evidence.candidates}),
                "has_cyrillic": bool(evidence.original_cyrillic),
                "has_latin": bool(evidence.original_latin),
            }
            for name, evidence in case.evidence.items()
        },
        "mrz": [
            {
                "valid": r.valid,
                "line_lengths": [len(s) for s in r.lines],
                "checks": r.checks,
                "conflicts": r.conflicts,
                "errors": r.errors,
                "optional_2_filler_count": r.lines[1][18:29].count("<")
                if len(r.lines) == 3 and len(r.lines[1]) == 30
                else None,
                "pin_matches_visual": any(
                    c.field_name == "personal_identifier"
                    and c.normalized_value == r.values.get("personal_identifier")
                    for c in result.candidates
                ),
            }
            for r in case.mrz_reports.values()
        ],
        "seconds": round(result.elapsed, 2),
        "private_data_saved": False,
    }
    print(json.dumps(report, ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()
