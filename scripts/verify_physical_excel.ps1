param([Parameter(Mandatory=$true)][string]$Result, [string]$ExpectedIssuer = '')
$ErrorActionPreference = 'Stop'
$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$resultPath = (Resolve-Path -LiteralPath $Result).Path
if (-not $resultPath.StartsWith((Join-Path $projectRoot 'test-artifacts\'), [StringComparison]::OrdinalIgnoreCase)) { throw 'Expected a workspace test result' }
$sourcePath = Join-Path $projectRoot 'assets\lermontov.xlsx'
$sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
$excel = $null
$book = $null
$report = [ordered]@{}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($resultPath, 0, $true)
    $sheet = $book.Worksheets.Item('Пасп Доп Данные')
    $data = $book.Worksheets.Item('Дог.')
    if ($ExpectedIssuer) { $report.issuer = $data.Range('J19').Value2 -eq $ExpectedIssuer }
    $report.full_name = $data.Range('A18').Value2 -eq 'СИНТЕТИКОВ ТЕСТ ПРИМЕРОВИЧ'
    $report.document_type = $data.Range('E19').Value2 -eq 'ID'
    $report.number_leading_zeros = $data.Range('G19').Value2 -eq '0000001'
    $report.pin_leading_zeros = $data.Range('C22').Value2 -eq '00123456789012'
    $report.photo_sheet_shapes = $sheet.Shapes.Count -eq 3
    $owned = 0
    foreach ($shape in $sheet.Shapes) {
        if ($shape.Name -notlike 'DocumentExcel|*') { continue }
        $owned++
        $side = if ($shape.Name -like '*|FRONT') { 'FRONT' } else { 'BACK' }
        $range = $sheet.Range($(if ($side -eq 'FRONT') { 'A1:E24' } else { 'F1:K24' }))
        $report["${side}_inside"] = $shape.Left -ge $range.Left - 0.75 -and $shape.Top -ge $range.Top - 0.75 -and $shape.Left + $shape.Width -le $range.Left + $range.Width + 0.75 -and $shape.Top + $shape.Height -le $range.Top + $range.Height + 0.75
        $report["${side}_ratio"] = [Math]::Abs($shape.Width / $shape.Height - (1900.0 / 1200.0)) -lt 0.01
    }
    $report.two_owned_photos = $owned -eq 2
    $pdfPath = Join-Path ([System.IO.Path]::GetDirectoryName($resultPath)) 'physical-photos-native.pdf'
    $sheet.Range('A1:K24').ExportAsFixedFormat(0, $pdfPath)
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
$report.source_unchanged = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash -eq $sourceHash
$reportPath = Join-Path ([System.IO.Path]::GetDirectoryName($resultPath)) 'native-excel-report.json'
$report | ConvertTo-Json | Set-Content -LiteralPath $reportPath -Encoding UTF8
$report | ConvertTo-Json
if (@($report.Values | Where-Object { $_ -ne $true }).Count) { throw 'Native Excel verification failed' }
