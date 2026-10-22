param([Parameter(Mandatory=$true)][string]$Directory, [switch]$InPlace)
$ErrorActionPreference = 'Stop'
$checkedDirectory = [System.IO.Path]::GetFullPath($Directory)
$sourcePath = Join-Path $checkedDirectory 'synthetic-input.xlsx'
$resultPath = if ($InPlace) { $sourcePath } else { Join-Path $checkedDirectory 'synthetic-result.xlsx' }
if ($InPlace -and (Test-Path -LiteralPath (Join-Path $checkedDirectory 'report.json'))) {
    $report = Get-Content -LiteralPath (Join-Path $checkedDirectory 'report.json') -Raw | ConvertFrom-Json
    if ($report.output) { $resultPath = [System.IO.Path]::GetFullPath($report.output); $sourcePath = $resultPath }
}
$before = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($resultPath, 0, $true)
    $excel.CalculateFull()
    $main = $book.Worksheets.Item(1)
    $tools = $book.Worksheets.Item(2)
    $calc = $book.Worksheets.Item(3)
    $photos = $book.Worksheets.Item(6)
    $checks = [ordered]@{
        opened_read_only = [bool]$book.ReadOnly
        synthetic_name_written = $main.Range('A18').Value2 -ceq 'Синтетиков Пример Тестович'
        missing_pin_empty = $null -eq $main.Range('C22').Value2
        missing_registration_empty = $null -eq $main.Range('J20').Value2
        price_preserved = $tools.Range('N20').Value2 -eq 350
        deposit_preserved = $tools.Range('P20').Value2 -eq 5000
        formula_preserved = $calc.Range('L21').Formula -ceq '=G21-F21'
        one_photo_inserted = $photos.Shapes.Count -eq 1
        source_unchanged = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash -ceq $before
    }
    foreach ($key in @($checks.Keys)) {
        if (-not $checks[$key]) { throw "Native Excel check failed: $key" }
    }
    $checks | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $checkedDirectory 'native-excel.json') -Encoding utf8
    $checks | ConvertTo-Json -Compress
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
