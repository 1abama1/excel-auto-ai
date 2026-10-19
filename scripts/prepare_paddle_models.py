"""Explicit ONLINE provisioning. The application itself never downloads models."""

from __future__ import annotations

import argparse
import json
import shutil
import tarfile
import tempfile
import urllib.request
from pathlib import Path

from document_excel.infrastructure.paddle_models import MODELS, sha256, validate_bundle

BASE = "https://paddle-model-ecology.bj.bcebos.com/paddlex/official_inference_model/paddle3.0.0/"


def prepare(destination: Path) -> None:
    destination.mkdir(parents=True, exist_ok=True)
    records = {}
    for model in MODELS:
        target = destination / model
        required = ("inference.json", "inference.pdiparams", "inference.yml")
        if not all((target / name).is_file() for name in required):
            with tempfile.TemporaryDirectory(prefix="model-", dir=destination) as temporary:
                staging = Path(temporary)
                archive = staging / "model.tar"
                url = BASE + model + "_infer.tar"
                print("Download:", model, flush=True)
                with urllib.request.urlopen(url, timeout=90) as source, archive.open("wb") as output:
                    shutil.copyfileobj(source, output)
                unpacked = staging / "unpacked"
                with tarfile.open(archive) as package:
                    package.extractall(unpacked, filter="data")
                configs = list(unpacked.rglob("inference.yml"))
                if len(configs) != 1:
                    raise ValueError("Unexpected model archive layout")
                source_dir = configs[0].parent
                if not all((source_dir / name).is_file() for name in required):
                    raise ValueError("Incomplete model archive")
                target.mkdir(exist_ok=True)
                for name in required:
                    shutil.copy2(source_dir / name, target / name)
        records[model] = {name: sha256(target / name) for name in required}
        records[model]["source"] = BASE + model + "_infer.tar"
    (destination / "manifest.json").write_text(
        json.dumps({"ocr_version": "PP-OCRv5", "models": records}, indent=2), encoding="utf-8"
    )
    validate_bundle(destination)
    print("Verified local bundle:", destination, flush=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("destination", type=Path)
    prepare(parser.parse_args().destination.resolve())
