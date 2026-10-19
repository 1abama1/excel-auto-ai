"""Build one Windows EXE: tested core + automatic pinned Paddle CPU installer."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import struct
import subprocess
import tomllib
import zipfile
from pathlib import Path

MAGIC = b"DOCEXCEL-ZIP-V1!"
PADDLE = {"paddleocr": "paddleocr", "paddlepaddle": "paddle", "paddlex": "paddlex"}


def digest(path):
    with path.open("rb") as source:
        return hashlib.file_digest(source, "sha256").hexdigest()


def build(base, output, work):
    root = Path(__file__).resolve().parents[1]
    work.mkdir(parents=True, exist_ok=True)
    output.parent.mkdir(parents=True, exist_ok=True)
    lock = tomllib.loads((root / "uv.lock").read_text(encoding="utf8"))
    packages = []
    for package in lock["package"]:
        if package["name"] not in PADDLE:
            continue
        wheel = next(
            w
            for w in package["wheels"]
            if w["url"].endswith(("cp312-cp312-win_amd64.whl", "py3-none-any.whl"))
        )
        packages.append(
            {
                "name": package["name"],
                "version": package["version"],
                "url": wheel["url"],
                "filename": wheel["url"].rsplit("/", 1)[-1],
                "sha256": wheel["hash"].split(":")[1],
                "size": wheel["size"],
                "module": PADDLE[package["name"]],
            }
        )
    manifest = {
        "packages": packages,
        "paddle_models": json.loads((base / "models/paddle/manifest.json").read_text(encoding="utf8")),
    }
    (work / "components.json").write_text(json.dumps(manifest, indent=2), encoding="utf8")
    archive = work / "core.zip"
    count, unpacked = 0, 0
    print("Compress tested application, Python, CPU dependencies and RapidOCR models", flush=True)
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED, compresslevel=6, allowZip64=True) as zip:
        for source in sorted(base.rglob("*")):
            if not source.is_file():
                continue
            relative = source.relative_to(base)
            parts = relative.parts
            if parts[:2] == ("app", "src"):
                continue
            if len(parts) >= 4 and parts[:3] == ("runtime", "Lib", "site-packages"):
                name = parts[3]
                if name in set(PADDLE.values()) or any(
                    name.startswith(n + "-") and name.endswith(".dist-info") for n in PADDLE
                ):
                    continue
            if parts[0] == "runtime" and len(parts) > 1 and parts[1] == "Scripts":
                # Wheel-installed command wrappers embed the build interpreter.
                # The desktop program uses Python modules exclusively.
                continue
            if relative.suffix == ".pyc" or "__pycache__" in parts:
                continue
            zip.write(source, relative.as_posix())
            count += 1
            unpacked += source.stat().st_size
        # Package current application sources, including fixes verified after
        # the reusable native runtime was assembled.
        for source in sorted((root / "src").rglob("*")):
            if source.is_file() and source.suffix != ".pyc" and "__pycache__" not in source.parts:
                zip.write(source, "app/src/" + source.relative_to(root / "src").as_posix())
        for source, name in (
            (root / "installer/provision.py", "installer/provision.py"),
            (work / "components.json", "installer/components.json"),
            (root / "assets/lermontov.xlsx", "templates/lermontov.xlsx"),
            (root / "docs/ONEFILE_WINDOWS.md", "docs/ONEFILE_WINDOWS.md"),
        ):
            zip.write(source, name)
        # App-local Microsoft runtime files shipped in the licensed Qt wheel;
        # no system-wide Visual C++ installer or admin access is necessary.
        for source in (base / "runtime/Lib/site-packages/PySide6").glob("*.dll"):
            if source.name.startswith(("msvcp140", "concrt140", "vcomp140")):
                zip.write(source, "runtime/" + source.name)
        count = len(zip.infolist())
        unpacked = sum(item.file_size for item in zip.infolist())
    finish(root, output, work, archive, count, unpacked)


def finish(root, output, work, archive, count, unpacked):
    packages = json.loads((work / "components.json").read_text(encoding="utf8"))["packages"]
    package_hash = digest(archive)
    identity = "v1-" + package_hash[:12]
    info = work / "PackageInfo.cs"
    info.write_text(
        'static class PackageInfo { public const string Id="'
        + identity
        + '"; public const string Sha256="'
        + package_hash
        + '"; }',
        encoding="utf8",
    )
    app_manifest = work / "app.manifest"
    app_manifest.write_text(
        """<?xml version="1.0" encoding="utf-8"?>
<assembly manifestVersion="1.0" xmlns="urn:schemas-microsoft-com:asm.v1">
<assemblyIdentity version="1.0.0.0" name="DocumentExcel.Bootstrap"/>
<trustInfo xmlns="urn:schemas-microsoft-com:asm.v3"><security><requestedPrivileges><requestedExecutionLevel level="asInvoker" uiAccess="false"/></requestedPrivileges></security></trustInfo>
<application xmlns="urn:schemas-microsoft-com:asm.v3"><windowsSettings><longPathAware xmlns="http://schemas.microsoft.com/SMI/2016/WindowsSettings">true</longPathAware></windowsSettings></application>
</assembly>""",
        encoding="utf8",
    )
    compiler = Path(os.environ.get("WINDIR", "C:/Windows")) / "Microsoft.NET/Framework64/v4.0.30319/csc.exe"
    stub = work / "launcher.exe"
    command = [
        str(compiler),
        "/nologo",
        "/target:winexe",
        "/platform:x64",
        "/optimize+",
        "/codepage:65001",
        "/out:" + str(stub),
        "/win32manifest:" + str(app_manifest),
        "/win32icon:" + str(root / "src/document_excel/resources/app.ico"),
        "/reference:System.Windows.Forms.dll",
        "/reference:System.Drawing.dll",
        "/reference:System.IO.Compression.dll",
        "/reference:System.Core.dll",
        str(root / "installer/Bootstrap.cs"),
        str(info),
    ]
    subprocess.run(command, check=True)
    temporary = output.with_suffix(".building")
    with temporary.open("wb") as destination:
        with stub.open("rb") as source:
            shutil.copyfileobj(source, destination)
        with archive.open("rb") as source:
            shutil.copyfileobj(source, destination)
        destination.write(MAGIC + struct.pack("<q", archive.stat().st_size))
    os.replace(temporary, output)
    report = {
        "exe": output.name,
        "sha256": digest(output),
        "payload_sha256": package_hash,
        "install_id": identity,
        "bytes": output.stat().st_size,
        "unpacked_core_bytes": unpacked,
        "core_files": count,
        "paddle_download_bytes": sum(p["size"] for p in packages),
        "cpu_only": True,
        "automatic_paddle_installation": True,
        "python_included": True,
        "primary_engine": "rapidocr",
        "ocr_after_import": False,
        "native_models_included": True,
        "low_disk_auto_fallback": True,
        "temporary_files_on_install_drive": True,
    }
    output.with_suffix(".build.json").write_text(json.dumps(report, indent=2), encoding="utf8")
    print(json.dumps(report), flush=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base", type=Path, default=Path("dist-digital/DocumentExcel"))
    parser.add_argument("--output", type=Path, default=Path("dist-onefile/DocumentExcel.exe"))
    parser.add_argument("--work", type=Path, default=Path("build/onefile"))
    parser.add_argument("--repack-launcher", action="store_true")
    args = parser.parse_args()
    if args.repack_launcher:
        archive = args.work.resolve() / "core.zip"
        root = Path(__file__).resolve().parents[1]
        with zipfile.ZipFile(archive) as package:
            assert package.read("installer/provision.py") == (root / "installer/provision.py").read_bytes(), (
                "Provisioner changed; rebuild the core"
            )
            count, unpacked = len(package.infolist()), sum(i.file_size for i in package.infolist())
        finish(root, args.output.resolve(), args.work.resolve(), archive, count, unpacked)
    else:
        build(args.base.resolve(), args.output.resolve(), args.work.resolve())
