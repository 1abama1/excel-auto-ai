"""Compare local photos with user-supplied expectations; output booleans only."""

import argparse
import json
import re
import threading
from pathlib import Path

from document_excel.application.workflows import RecognitionService
from document_excel.domain.models import ClientCase, SourceImage
from document_excel.infrastructure.extraction import apply_candidates
from document_excel.infrastructure.ocr import TesseractEngine
from document_excel.infrastructure.storage import Settings


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("directory", type=Path)
    parser.add_argument("expectations", type=Path)
    args = parser.parse_args()
    specification = args.expectations.read_text(encoding="utf-8")
    expected = {
        key: re.search(rf'{key}\s*=\s*"([^"\n]+)"', specification).group(1)
        for key in ("passport_issued_by", "document_number")
    }
    observed = []

    class Engine(TesseractEngine):
        def _recognize(self, *params, **kwargs):
            result = super()._recognize(*params, **kwargs)
            observed.append(
                {
                    "variant": result.variant,
                    "authority_literal_found": expected["passport_issued_by"] in result.text,
                    "authority_casefold_matches": " ".join(result.text.split()).casefold()
                    == expected["passport_issued_by"].casefold(),
                    "number_literal_found": expected["document_number"] in result.text.replace(" ", ""),
                    "input_shape": list(params[0].shape),
                    "authority_pattern": "".join(
                        "9"
                        if c.isdigit()
                        else "C"
                        if "\u0400" <= c <= "\u04ff"
                        else "L"
                        if c.isalpha()
                        else c
                        for c in result.text
                    )
                    if "official-passport_issued_by" in result.variant
                    else None,
                }
            )
            return result

    sources = [SourceImage(path=args.directory / f"{index}.png") for index in (1, 2)]
    output = RecognitionService(Settings(), Engine(Settings())).run(
        sources, threading.Event(), lambda *_: None
    )
    case = ClientCase(images=sources, mrz_reports=output.mrz_reports)
    apply_candidates(case, output.candidates, output.document_type, output.label_seen)
    fields = {}
    for name, target in (
        ("passport_issued_by", expected["passport_issued_by"]),
        ("passport_number", expected["document_number"][2:]),
    ):
        evidence = case.evidence.get(name)
        fields[name] = {
            "selected_matches": getattr(case.data, name) == target,
            "status": evidence.status.value if evidence else "MISSING",
            "candidates": [
                {
                    "matches": c.normalized_value == target,
                    "side": c.source_side.value,
                    "method": c.extraction_method,
                    "source": c.source_method,
                    "bbox": c.source_bbox,
                }
                for c in evidence.candidates
            ]
            if evidence
            else [],
        }
    print(json.dumps({"fields": fields, "optical_passes": observed, "private_values_saved": False}, indent=2))


if __name__ == "__main__":
    main()
