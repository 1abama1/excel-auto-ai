"""Create a portable Windows CPU distribution with Python, runtime and models.

Uses a private embedded runtime instead of freezing Paddle's dynamic imports.
Never overwrites an existing distribution; no dependency on the source venv
or a system Python remains in the resulting launcher.
"""

from __future__ import annotations

import argparse
import compileall
import json
import re
import shutil
import sys
import sysconfig
from importlib.metadata import distributions
from pathlib import Path

from document_excel.infrastructure.paddle_models import validate_bundle


def prepare_bytecode(output: Path):
    """Compile the copied runtime during packaging, instead of the first OCR."""
    excluded = re.compile(r"(?:^|[\\/])(?:tests?|examples?)(?:[\\/]|$)")
    for directory in (output / "runtime" / "Lib", output / "app" / "src"):
        if not compileall.compile_dir(directory, quiet=2, rx=excluded, stripdir=str(output), prependdir="."):
            raise RuntimeError("Portable runtime bytecode compilation failed")


def package(output: Path, models: Path):
    if sys.platform != "win32" or sys.version_info[:2] != (3, 12):
        raise RuntimeError("Build with the tested Windows x64 Python 3.12 Paddle environment")
    if output.exists():
        raise FileExistsError("Choose a new output directory")
    validate_bundle(models)
    root = Path(__file__).resolve().parents[1]
    output.mkdir(parents=True)
    runtime = output / "runtime"
    print("Copy standalone Python", flush=True)
    shutil.copytree(Path(sys.base_prefix), runtime, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
    site = runtime / "Lib" / "site-packages"
    site.mkdir(parents=True, exist_ok=True)
    print("Copy installed inference dependencies", flush=True)
    # Dependency helpers such as numpy._pytesttester/shapely._coverage are
    # runtime imports. Broad "test"/"coverage" exclusions break inference.
    exclusions = shutil.ignore_patterns("__pycache__", "*.pyc", "__editable__*")
    shutil.copytree(Path(sysconfig.get_path("purelib")), site, dirs_exist_ok=True, ignore=exclusions)
    # Editable installation helpers must never refer to the build workstation.
    for pth in site.glob("*.pth"):
        if str(root).lower() in pth.read_text(encoding="utf-8", errors="ignore").lower():
            pth.unlink()
    (site / "document_excel_portable.pth").write_text("../../../app/src\n", encoding="utf-8")
    print("Copy application and verified offline models", flush=True)
    app = output / "app"
    shutil.copytree(
        root / "src" / "document_excel",
        app / "src" / "document_excel",
        ignore=shutil.ignore_patterns("__pycache__", "*.pyc"),
    )
    shutil.copy2(root / "run.py", app / "run.py")
    shutil.copytree(models, output / "models" / "paddle")
    validate_bundle(output / "models" / "paddle")
    print("Precompile portable Python modules", flush=True)
    prepare_bytecode(output)
    launcher = """@echo off
setlocal
cd /d "%~dp0"
set "DOCUMENT_EXCEL_PADDLE_MODELS=%~dp0models\\paddle"
set "PYTHONPATH=%~dp0app\\src"
set "PYTHONNOUSERSITE=1"
set "PYTHONUTF8=1"
start "" "%~dp0runtime\\pythonw.exe" "%~dp0app\\run.py" %*
"""
    (output / "DocumentExcel.cmd").write_bytes(launcher.replace("\n", "\r\n").encode("ascii"))
    for name in ("README.md", "USER_GUIDE.md", "THIRD_PARTY_NOTICES.md"):
        shutil.copy2(root / name, output / name)
    licenses = output / "licenses"
    licenses.mkdir()
    locked = []
    for dist in distributions():
        name = dist.metadata["Name"] or "unknown"
        locked.append(f"{name}=={dist.version}")
        for relative in dist.files or []:
            if "license" in str(relative).lower() or "copying" in relative.name.lower():
                source = Path(dist.locate_file(relative))
                if source.is_file():
                    target = licenses / name / str(relative)
                    target.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(source, target)
    (output / "requirements-locked.txt").write_text("\n".join(sorted(locked)) + "\n", encoding="utf-8")
    (output / "build.json").write_text(
        json.dumps(
            {"python": sys.version, "platform": sys.platform, "model_bundle": "PP-OCRv5", "offline": True},
            indent=2,
        )
    )
    print("Portable Windows distribution:", output, flush=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--models", required=True, type=Path)
    args = parser.parse_args()
    package(args.output.resolve(), args.models.resolve())
