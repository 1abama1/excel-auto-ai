"""Explicit ONLINE provisioning. Runtime inference never downloads anything."""

import argparse
import hashlib
import json
import os
import tempfile
import urllib.request
from pathlib import Path

from document_excel.infrastructure.rapid_models import MODEL_FILES, validate_bundle

BASE = "https://www.modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.8.0/onnx/PP-OCRv5/"


def prepare(destination):
    destination.mkdir(parents=True, exist_ok=True)
    records = {}
    for kind, (filename, expected) in MODEL_FILES.items():
        target = destination / filename
        digest = hashlib.sha256(target.read_bytes()).hexdigest() if target.is_file() else ""
        if digest != expected:
            url = BASE + ("det/" if kind == "det" else "rec/") + filename
            descriptor, name = tempfile.mkstemp(dir=destination, suffix=".download")
            temporary = Path(name)
            try:
                with (
                    os.fdopen(descriptor, "wb") as output,
                    urllib.request.urlopen(url, timeout=90) as response,
                ):
                    while chunk := response.read(1024 * 1024):
                        output.write(chunk)
                with temporary.open("rb") as stream:
                    if hashlib.file_digest(stream, "sha256").hexdigest() != expected:
                        raise ValueError("ONNX model checksum mismatch")
                os.replace(temporary, target)
            finally:
                temporary.unlink(missing_ok=True)
        records[kind] = {"file": filename, "sha256": expected}
    (destination / "manifest.json").write_text(
        json.dumps({"backend": "rapidocr", "ocr_version": "PP-OCRv5", "models": records}, indent=2),
        encoding="utf-8",
    )
    validate_bundle(destination)
    print("Verified local CPU ONNX bundle", flush=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("destination", type=Path)
    prepare(parser.parse_args().destination.resolve())
