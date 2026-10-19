@echo off
setlocal
cd /d "%~dp0"
if exist "%~dp0dist-paddle\DocumentExcel\DocumentExcel.cmd" (
    start "" "%~dp0dist-paddle\DocumentExcel\DocumentExcel.cmd"
    exit /b
)
if not defined DOCUMENT_EXCEL_PADDLE_PYTHON set "DOCUMENT_EXCEL_PADDLE_PYTHON=%~dp0.venv-paddle\Scripts\pythonw.exe"
if not exist "%DOCUMENT_EXCEL_PADDLE_PYTHON%" (
    echo Run scripts\setup_paddle.ps1, or set DOCUMENT_EXCEL_PADDLE_PYTHON to the installed Python 3.12 runtime.
    pause
    exit /b 1
)
if exist "%~dp0tools\tesseract\tesseract.exe" set "DOCUMENT_EXCEL_TESSERACT=%~dp0tools\tesseract\tesseract.exe"
start "" "%DOCUMENT_EXCEL_PADDLE_PYTHON%" "%~dp0run.py"
