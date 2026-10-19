param(
    [string]$RuntimeDirectory = '',
    [string]$ModelDirectory = '',
    [string]$Python = '3.12'
)
$ErrorActionPreference = 'Stop'
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if (-not $RuntimeDirectory) { $RuntimeDirectory = Join-Path $projectRoot '.venv-paddle' }
if (-not $ModelDirectory) { $ModelDirectory = Join-Path $projectRoot 'models\paddle' }
$RuntimeDirectory = [IO.Path]::GetFullPath($RuntimeDirectory)
$ModelDirectory = [IO.Path]::GetFullPath($ModelDirectory)
Set-Location -LiteralPath $projectRoot
uv venv $RuntimeDirectory --python $Python
if ($LASTEXITCODE -ne 0) { throw 'Python 3.12 x64 required. Install it with uv python install 3.12.' }
$runtimePython = Join-Path $RuntimeDirectory 'Scripts\python.exe'
uv pip install --python $runtimePython -e ($projectRoot + '[paddle,dev]')
if ($LASTEXITCODE -ne 0) { throw 'Paddle runtime installation failed.' }
& $runtimePython (Join-Path $PSScriptRoot 'prepare_paddle_models.py') $ModelDirectory
if ($LASTEXITCODE -ne 0) { throw 'Offline models verification failed.' }
Write-Output "Runtime: $runtimePython"
Write-Output "Models: $ModelDirectory"
Write-Output 'Install Microsoft Visual C++ Redistributable 2015-2022 x64 on the target Windows machine.'
