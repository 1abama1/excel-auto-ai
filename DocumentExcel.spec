from pathlib import Path
import sys
from PyInstaller.utils.hooks import collect_data_files

root = Path(SPECPATH)
a = Analysis(
    [str(root / 'run.py')],
    pathex=[str(root / 'src')],
    binaries=[],
    datas=collect_data_files('document_excel'),
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=['PySide6.QtWebEngineCore', 'PySide6.QtWebEngineWidgets', 'PySide6.QtQml', 'PySide6.QtQuick', 'PySide6.QtPdf', 'matplotlib', 'pandas', 'tkinter', 'pytest'],
    noarchive=False,
)
if sys.platform == 'win32':
    # Qt 6.12 uses the unversioned Windows 10 ICU API. A different icuuc.dll
    # from tools on PATH has incompatible exports; use the operating system DLL.
    a.binaries = [entry for entry in a.binaries if Path(entry[0]).name.lower() != 'icuuc.dll']
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name='DocumentExcel',
          debug=False, bootloader_ignore_signals=False, strip=False, upx=False,
          console=False, icon=str(root / 'src/document_excel/resources/app.ico'))
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name='DocumentExcel')
