"""Build a standalone CPU release from a verified local Windows runtime.

Identical files may use hard links on the build volume to avoid duplicate disk
usage. The output contains every file: copying/zipping it is self-contained.
Overlays use atomic replacement so earlier releases cannot be modified via links.
"""

from __future__ import annotations

import argparse
import compileall
import hashlib
import importlib.metadata
import json
import os
import shutil
import sys
from pathlib import Path

from document_excel.infrastructure.paddle_models import validate_bundle as validate_paddle
from document_excel.infrastructure.rapid_models import validate_bundle as validate_rapid


def link_or_copy(source, target):
    source, target = Path(source), Path(target)
    if target.exists():
        if source.read_bytes() == target.read_bytes():
            return str(target)
        temporary = target.with_name(target.name + ".cpu-overlay")
        assert not temporary.exists()
        link_or_copy(source, temporary)
        os.replace(temporary, target)
        return str(target)
    try:
        os.link(source, target)
    except OSError:
        shutil.copy2(source, target)
    return str(target)


def package(base, output, packages, models, release="cpu-ppocrv5-2026-10-19"):
    if sys.platform != "win32" or sys.version_info[:2] != (3, 12):
        raise RuntimeError("Package with Windows x64 Python 3.12")
    if output.exists():
        raise FileExistsError("Choose a new release directory")
    validate_rapid(models)
    validate_paddle(base / "models" / "paddle")
    root = Path(__file__).resolve().parents[1]
    output.mkdir(parents=True)
    print("Clone standalone CPU runtime", flush=True)
    shutil.copytree(base / "runtime", output / "runtime", copy_function=link_or_copy)
    site = output / "runtime" / "Lib" / "site-packages"
    selected = []
    for source in packages.iterdir():
        if any(
            source.name == name or source.name.startswith(name + "-")
            for name in ("rapidocr", "onnxruntime", "pyclipper", "omegaconf")
        ):
            if source.is_dir():
                shutil.copytree(source, site / source.name, dirs_exist_ok=True, copy_function=link_or_copy)
            else:
                link_or_copy(source, site / source.name)
            selected.append(source.name)
    if not any(name.startswith("rapidocr-") for name in selected):
        raise RuntimeError("Verified RapidOCR runtime packages are missing")
    print("Copy current application and local models", flush=True)
    shutil.copytree(
        root / "src" / "document_excel",
        output / "app" / "src" / "document_excel",
        ignore=shutil.ignore_patterns("__pycache__", "*.pyc"),
    )
    shutil.copy2(root / "run.py", output / "app" / "run.py")
    shutil.copytree(models, output / "models" / "rapidocr", copy_function=link_or_copy)
    shutil.copytree(base / "models" / "paddle", output / "models" / "paddle", copy_function=link_or_copy)
    validate_rapid(output / "models" / "rapidocr")
    validate_paddle(output / "models" / "paddle")
    launcher = r"""@echo off
setlocal
cd /d "%~dp0"
set "DOCUMENT_EXCEL_RAPID_MODELS=%~dp0models\rapidocr"
set "DOCUMENT_EXCEL_PADDLE_MODELS=%~dp0models\paddle"
set "PYTHONPATH=%~dp0app\src"
set "PYTHONNOUSERSITE=1"
set "PYTHONUTF8=1"
start "" "%~dp0runtime\pythonw.exe" "%~dp0app\run.py" %*
"""
    (output / "DocumentExcel.cmd").write_bytes(launcher.replace("\n", "\r\n").encode("ascii"))
    for name in ("README.md", "USER_GUIDE.md", "THIRD_PARTY_NOTICES.md", "TEST_REPORT.md"):
        shutil.copy2(root / name, output / name)
    shutil.copytree(root / "docs", output / "docs")
    shutil.copytree(base / "licenses", output / "licenses", copy_function=link_or_copy)
    shutil.copytree(
        root / "licenses" / "cpu", output / "licenses" / "cpu", dirs_exist_ok=True, copy_function=link_or_copy
    )
    for name in selected:
        if name.endswith(".dist-info"):
            shutil.copytree(
                site / name,
                output / "licenses" / "cpu-dependencies" / name,
                dirs_exist_ok=True,
                copy_function=link_or_copy,
            )
    versions = {
        d.metadata["Name"]: d.version
        for d in importlib.metadata.distributions(path=[str(site)])
        if d.metadata["Name"]
    }
    forbidden = [
        name
        for name in versions
        if name.lower().startswith("nvidia-")
        or name.lower() in {"paddlepaddle-gpu", "onnxruntime-gpu", "tensorrt"}
    ]
    if forbidden:
        raise RuntimeError("GPU packages are not permitted in the CPU release")
    (output / "requirements-locked.txt").write_text(
        "\n".join(f"{n}=={v}" for n, v in sorted(versions.items())) + "\n", encoding="utf-8"
    )
    print("Precompile application and CPU adapter dependencies", flush=True)
    for path in (
        output / "app" / "src",
        *(site / n for n in ("rapidocr", "onnxruntime", "pyclipper", "omegaconf")),
    ):
        if not compileall.compile_dir(path, quiet=2, stripdir=str(output), prependdir="."):
            raise RuntimeError("CPU bytecode compilation failed")
    hashes = {
        str(p.relative_to(root / "src")): hashlib.sha256(p.read_bytes()).hexdigest()
        for p in (root / "src" / "document_excel").rglob("*")
        if p.is_file() and "__pycache__" not in p.parts and p.suffix != ".pyc"
    }
    (output / "build.json").write_text(
        json.dumps(
            {
                "release": release,
                "python": sys.version,
                "primary_engine": "rapidocr",
                "provider": "CPUExecutionProvider",
                "cpu_threads": 2,
                "gpu_dependencies": [],
                "offline": True,
                "source_hashes": hashes,
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    print("Standalone CPU release ready", flush=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ("base", "output", "packages", "models"):
        parser.add_argument("--" + name, type=Path, required=True)
    parser.add_argument("--release", default="cpu-ppocrv5-2026-10-19")
    args = parser.parse_args()
    package(
        args.base.resolve(),
        args.output.resolve(),
        args.packages.resolve(),
        args.models.resolve(),
        args.release,
    )
