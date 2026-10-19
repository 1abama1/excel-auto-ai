#!/usr/bin/env sh
set -eu
cd "$(dirname "$0")/.."
export PYINSTALLER_CONFIG_DIR="$PWD/.pyinstaller-cache"
.venv/bin/python scripts/build_assets.py
.venv/bin/python -m PyInstaller --noconfirm DocumentExcel.spec
.venv/bin/python scripts/collect_licenses.py
printf '%s\n' "dist/DocumentExcel/DocumentExcel"
