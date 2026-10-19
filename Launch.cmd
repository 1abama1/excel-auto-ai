@echo off
setlocal
cd /d "%~dp0"
if exist "%~dp0tools\tesseract\tesseract.exe" set "DOCUMENT_EXCEL_TESSERACT=%~dp0tools\tesseract\tesseract.exe"
set "DOCUMENT_EXCEL_BUILD="
if exist "%~dp0current-build.txt" set /p "DOCUMENT_EXCEL_BUILD="<"%~dp0current-build.txt"
if defined DOCUMENT_EXCEL_BUILD if exist "%~dp0%DOCUMENT_EXCEL_BUILD%" (
    start "" "%~dp0%DOCUMENT_EXCEL_BUILD%"
    exit /b
)
if exist "%~dp0dist\DocumentExcel\DocumentExcel.exe" (
    start "" "%~dp0dist\DocumentExcel\DocumentExcel.exe"
    exit /b
)
if exist "%~dp0.venv\Scripts\pythonw.exe" (
    start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0run.py"
    exit /b
)
echo Install the application dependencies first. See README.md.
pause
