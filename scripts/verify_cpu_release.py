"""CPU bundle integrity and private-control leakage audit; masked output only."""

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path

import onnxruntime
import paddle

from document_excel.infrastructure.paddle_models import validate_bundle as validate_paddle
from document_excel.infrastructure.rapid_models import validate_bundle as validate_rapid


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("release", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--expected-stdin", action="store_true")
    parser.add_argument("--physical-control", type=Path)
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    release = args.release.resolve()
    metadata = json.loads((release / "build.json").read_text(encoding="utf8"))
    hashes = metadata["source_hashes"]
    matched = all(
        hashlib.sha256((release / "app/src" / name).read_bytes()).hexdigest()
        == digest
        == hashlib.sha256((root / "src" / name).read_bytes()).hexdigest()
        for name, digest in hashes.items()
    )
    validate_rapid(release / "models/rapidocr")
    validate_paddle(release / "models/paddle")
    packages = (release / "requirements-locked.txt").read_text(encoding="utf8").casefold()
    gpu = any(name in packages for name in ("nvidia-", "paddlepaddle-gpu", "onnxruntime-gpu", "tensorrt"))
    controls = json.loads(sys.stdin.readline()) if args.expected_stdin else {}
    if args.physical_control:
        from benchmark_cpu import control_names

        controls.update({"physical_" + k: v for k, v in control_names(args.physical_control).items()})
    sensitive = [
        str(v)
        for k, v in controls.items()
        if len(str(v)) >= 7 and ("name" in k or k in {"personal_identifier", "document_number"})
    ]
    if "document_number" in controls:
        sensitive.append(re.sub(r"^[A-Za-z]+", "", controls["document_number"]))
    patterns = [re.compile(r"(?<!\w)" + re.escape(v) + r"(?!\w)", re.I) for v in sensitive]
    files = [
        p
        for area in ("src", "tests", "scripts", "docs")
        for p in (root / area).rglob("*")
        if p.is_file() and p.suffix in {".py", ".json", ".md"}
    ]
    files += list(root.glob("*.md"))
    files += list((root / "test-artifacts/digital-bug").glob("*.json"))
    matches = sum(any(pattern.search(p.read_text(encoding="utf8")) for pattern in patterns) for p in files)
    report = {
        "source_files_verified": len(hashes) if matched else 0,
        "both_model_bundles_verified": True,
        "gpu_packages": gpu,
        "paddle_cuda_build": paddle.is_compiled_with_cuda(),
        "ort_cpu_available": "CPUExecutionProvider" in onnxruntime.get_available_providers(),
        "ort_cuda_available": "CUDAExecutionProvider" in onnxruntime.get_available_providers(),
        "privacy_files_checked": len(files),
        "control_literal_matches": matches,
        "private_work_directories_excluded": not any(
            (release / p).exists() for p in ("local-data", "test-artifacts")
        ),
    }
    report["ok"] = (
        matched
        and not gpu
        and not report["paddle_cuda_build"]
        and report["ort_cpu_available"]
        and not report["ort_cuda_available"]
        and not matches
        and report["private_work_directories_excluded"]
    )
    args.output.write_text(json.dumps(report, indent=2), encoding="utf8")
    print(json.dumps(report))
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
