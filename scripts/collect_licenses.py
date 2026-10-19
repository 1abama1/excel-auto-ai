from __future__ import annotations

import argparse
import importlib.metadata
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, default=ROOT / "dist/DocumentExcel")
    target = parser.parse_args().output / "LICENSES"
    target.mkdir(parents=True, exist_ok=True)
    for distribution in importlib.metadata.distributions():
        name = distribution.metadata.get("Name", "unknown")
        for file in distribution.files or []:
            if any(token in file.name.lower() for token in ("license", "copying", "notice")):
                source = Path(distribution.locate_file(file))
                if (
                    source.is_file()
                    and source.stat().st_size < 2_000_000
                    and source.suffix.lower() not in {".py", ".pyc", ".dll", ".pyd"}
                ):
                    folder = target / name
                    folder.mkdir(exist_ok=True)
                    # Preserve each path to avoid collisions between nested third-party notices.
                    filename = str(file).replace("\\", "_").replace("/", "_").replace(":", "_")
                    shutil.copyfile(source, folder / filename[-180:])
    for file in (
        "THIRD_PARTY_NOTICES.md",
        "README.md",
        "USER_GUIDE.md",
        "BUILD.md",
        "TEST_REPORT.md",
        "ARCHITECTURE.md",
    ):
        if (ROOT / file).exists():
            shutil.copyfile(ROOT / file, target.parent / file)
    for name in (
        "MAPPING.md",
        "REDESIGN.md",
        "REVIEW.md",
        "CAMERA_PHOTOS.md",
        "KG_PHYSICAL_ID.md",
        "ID_FIELDS_FIX.md",
        "OCR_STATE_AUDIT.md",
    ):
        if (ROOT / "docs" / name).exists():
            (target.parent / "docs").mkdir(exist_ok=True)
            shutil.copyfile(ROOT / "docs" / name, target.parent / "docs" / name)


if __name__ == "__main__":
    main()
