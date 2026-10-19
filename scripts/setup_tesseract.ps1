param([string]$Destination = (Join-Path $PSScriptRoot '..\tools\tesseract'))
$ErrorActionPreference = 'Stop'
$resolvedDestination = [System.IO.Path]::GetFullPath($Destination)
if (Test-Path -LiteralPath $resolvedDestination) {
    throw 'Choose a new, empty Tesseract installation directory.'
}
$installerDirectory = Join-Path $PSScriptRoot '..\tools\downloads'
New-Item -ItemType Directory -Force -Path $installerDirectory | Out-Null
$installer = Join-Path $installerDirectory 'tesseract-5.5.3-setup.exe'
$url = 'https://github.com/tesseract-ocr/tesseract/releases/download/5.5.3/tesseract-ocr-w64-setup-5.5.3.20260724.exe'
Invoke-WebRequest -Uri $url -OutFile $installer
$signature = Get-AuthenticodeSignature -LiteralPath $installer
if ($signature.Status -ne 'Valid') {
    $release = Invoke-RestMethod -Uri 'https://api.github.com/repos/tesseract-ocr/tesseract/releases/tags/5.5.3'
    $asset = $release.assets | Where-Object { $_.name -eq 'tesseract-ocr-w64-setup-5.5.3.20260724.exe' }
    $actualHash = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash.ToLowerInvariant()
    if (-not $asset -or $asset.digest -ne "sha256:$actualHash") {
        throw "Neither the signature nor the official release SHA256 could be verified. No installer was executed."
    }
    Write-Output "Verified SHA256 against the official GitHub release: $actualHash"
} else {
    Write-Output "Verified publisher: $($signature.SignerCertificate.Subject)"
}
$process = Start-Process -FilePath $installer -ArgumentList @('/S', "/D=$resolvedDestination") -WindowStyle Hidden -Wait -PassThru
if ($process.ExitCode -ne 0) { throw "Tesseract installer failed: $($process.ExitCode)" }
$tessdata = Join-Path $resolvedDestination 'tessdata'
foreach ($language in @('rus', 'eng', 'kir')) {
    $modelUrl = "https://raw.githubusercontent.com/tesseract-ocr/tessdata_best/main/$language.traineddata"
    Invoke-WebRequest -Uri $modelUrl -OutFile (Join-Path $tessdata "$language.traineddata")
}
& (Join-Path $resolvedDestination 'tesseract.exe') --version
& (Join-Path $resolvedDestination 'tesseract.exe') --list-langs
