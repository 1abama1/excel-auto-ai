$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath ([System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')))
$env:PYTHONUTF8 = '1'
& .venv\Scripts\ruff.exe check src tests scripts
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& .venv\Scripts\mypy.exe src\document_excel\domain src\document_excel\infrastructure src\document_excel\application
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& .venv\Scripts\python.exe -m pytest --cov=document_excel --cov-report=term-missing --junitxml=test-artifacts/pytest.xml
exit $LASTEXITCODE
