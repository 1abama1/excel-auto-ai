"""Anonymized geometry diagnostics; values are never printed or saved."""

import re
import sys
import threading
from pathlib import Path

from document_excel.domain.models import DocumentSide
from document_excel.infrastructure.document_profiles import KyrgyzPhysicalIDProfile
from document_excel.infrastructure.extraction import FieldExtractor, label_matches
from document_excel.infrastructure.mrz import MRZRecognizer
from document_excel.infrastructure.ocr import TesseractEngine
from document_excel.infrastructure.physical_id import layout_candidates, prepare_plastic_card, prepare_region
from document_excel.infrastructure.preprocessing import load_image
from document_excel.infrastructure.storage import Settings


def main():
    engine = TesseractEngine(Settings())
    cancel = threading.Event()
    for index in (1, 2):
        card = prepare_plastic_card(load_image(Path(sys.argv[1]) / f"{index}.png"))
        print("card", index, card.image.shape, card.detected)
        result = engine.recognize_with_psm(
            prepare_region(card.image, "text"), "private", "original", cancel, 11
        )
        for line in result.lines():
            fields = [name for name, _ in label_matches(line.text)]
            pattern = re.sub(r"[^<\s/.-]", lambda m: "D" if m[0].isdigit() else "A", line.text)
            print(fields, line.bbox, pattern)
        for c in FieldExtractor(KyrgyzPhysicalIDProfile()).extract(result):
            print(
                "candidate",
                c.field_name,
                c.validation_status,
                len(c.normalized_value),
                c.original_latin is not None,
            )
        values = layout_candidates(
            card, result, engine, cancel, DocumentSide.FRONT if index == 1 else DocumentSide.BACK
        )
        print(
            "layout",
            [
                (
                    c.field_name,
                    len(c.normalized_value),
                    c.original_cyrillic is not None,
                    c.original_latin is not None,
                )
                for c in values
            ],
        )
        if index == 2:
            print("boxes", MRZRecognizer.line_boxes(card.image))
            for fraction in (0.57, 0.66):
                region = card.image[round(card.image.shape[0] * fraction) :]
                result = engine.recognize_mrz(prepare_region(region, "text"), "private", cancel)
                print("mrz_region", fraction)
                for line in result.lines():
                    print(
                        len(line.text),
                        re.sub(r"[A-Z0-9]", lambda m: "D" if m[0].isdigit() else "A", line.text),
                    )


if __name__ == "__main__":
    main()
