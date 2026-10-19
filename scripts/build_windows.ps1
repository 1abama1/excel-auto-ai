param([string]$OutputDirectory = 'dist')
$ErrorActionPreference = 'Stop'
$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if ($OutputDirectory -notmatch '^[A-Za-z0-9][A-Za-z0-9_/\\-]*$') { throw 'Use a relative build directory inside the project.' }
$outputRoot = [System.IO.Path]::GetFullPath((Join-Path $projectRoot $OutputDirectory))
if (-not $outputRoot.StartsWith($projectRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) { throw 'Build output must stay inside the project.' }
$runningBuild = Get-Process -Name DocumentExcel -ErrorAction SilentlyContinue | Where-Object { $_.Path -and $_.Path.StartsWith($outputRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase) }
if ($runningBuild) { throw 'This build is running. Use -OutputDirectory with a different new directory; keep the current client window open.' }
Set-Location -LiteralPath $projectRoot
$python = Join-Path $projectRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $python)) { throw 'First install the development dependencies; see BUILD.md.' }
$env:PYINSTALLER_CONFIG_DIR = Join-Path $projectRoot '.pyinstaller-cache'
& $python 'scripts\build_assets.py'
if ($LASTEXITCODE -ne 0) { throw 'Asset preparation failed.' }
$previousBuildPath = $env:PATH
try {
    $basePython = & $python -c 'import sys; print(sys.base_prefix)'
    $env:PATH = @($basePython, (Split-Path -Parent $python), (Join-Path $env:SystemRoot 'System32'), $env:SystemRoot) -join [System.IO.Path]::PathSeparator
    & $python -m PyInstaller --clean --noconfirm --distpath $outputRoot 'DocumentExcel.spec'
    if ($LASTEXITCODE -ne 0) { throw 'PyInstaller failed.' }
} finally {
    $env:PATH = $previousBuildPath
}
$applicationRoot = Join-Path $outputRoot 'DocumentExcel'
& $python 'scripts\collect_licenses.py' --output $applicationRoot
if ($LASTEXITCODE -ne 0) { throw 'License collection failed.' }
[System.IO.File]::WriteAllText((Join-Path $projectRoot 'current-build.txt'), ($OutputDirectory.TrimEnd('\', '/') + '\DocumentExcel\DocumentExcel.exe'), [System.Text.Encoding]::ASCII)
Write-Output (Join-Path $applicationRoot 'DocumentExcel.exe')
