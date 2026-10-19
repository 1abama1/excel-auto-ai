param([Parameter(Mandatory=$true)][string]$ResultPath, [Parameter(Mandatory=$true)][string]$ReportDirectory)
$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
$sourceBook = $null
$root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$source = Join-Path $root 'assets\lermontov.xlsx'
$before = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
try {
    # Own isolated Excel instance. Never attach to an existing user workbook.
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $sourceBook = $excel.Workbooks.Open($source, 0, $true)
    $sourceSheets = $sourceBook.Worksheets.Count
    $sourceShapes = 0
    foreach ($sheet in $sourceBook.Worksheets) { $sourceShapes += $sheet.Shapes.Count }
    $sourceBook.Close($false)
    $sourceBook = $null
    $book = $excel.Workbooks.Open([System.IO.Path]::GetFullPath($ResultPath), 0, $true)
    $excel.CalculateFull()
    $main = $book.Worksheets.Item(1)
    $page2 = $book.Worksheets.Item(19)
    $act2 = $book.Worksheets.Item(20)
    $calc2 = $book.Worksheets.Item(21)
    $checks = [ordered]@{
        source_sheet_count = $sourceSheets
        result_sheet_count = $book.Worksheets.Count
        read_only = [bool]$book.ReadOnly
        name_correct = $main.Range('A18').Value2 -eq 'SYNTHETIC TEST CLIENT'
        pin_keeps_zeros = $main.Range('C22').Value2 -ceq '00123456789012'
        number_keeps_zeros = $main.Range('G19').Value2 -ceq '0000001'
        department_keeps_zeros = $main.Range('K19').Value2 -ceq '000123'
        original_shape_count = $sourceShapes
        second_page_name = $page2.Range('B20').Value2 -ceq 'TEST 6'
        linked_act_recalculated = $act2.Range('B21').Value2 -ceq 'TEST 6'
        linked_act_inventory = $act2.Range('L21').Value2 -ceq '0006'
        linked_daily_price_recalculated = $calc2.Range('N20').Value2 -eq 105
        logo_in_second_page = $page2.Shapes.Count -eq 1
        logo_in_second_act = $act2.Shapes.Count -eq 1
        deposit_not_repeated = $null -eq $page2.Range('P20').Value2
    }
    $photos = $book.Worksheets.Item(6)
    $checks.photo_sheet_images = $photos.Shapes.Count -eq 3
    foreach ($side in @('FRONT', 'BACK')) {
        $photo = @($photos.Shapes | Where-Object { $_.Name.StartsWith('DocumentExcel|') -and $_.Name.EndsWith('|' + $side) })
        $checks[('photo_' + $side.ToLower() + '_one')] = $photo.Count -eq 1
        if ($photo.Count -eq 1) {
            $area = $photos.Range($(if ($side -eq 'FRONT') { 'A1:E24' } else { 'F1:K24' }))
            $shape = $photo[0]
            $checks[('photo_' + $side.ToLower() + '_inside')] = $shape.Left -ge $area.Left - 0.5 -and $shape.Top -ge $area.Top - 0.5 -and $shape.Left + $shape.Width -le $area.Left + $area.Width + 0.5 -and $shape.Top + $shape.Height -le $area.Top + $area.Height + 0.5
            $ratio = if ($side -eq 'FRONT') { 1200.0/750.0 } else { 900.0/1200.0 }
            Write-Output ('PhotoGeometry=' + (@{ side=$side; width=$shape.Width; height=$shape.Height; ratio=($shape.Width/$shape.Height); area_width=$area.Width; area_height=$area.Height; normal_font=$book.Styles.Item('Normal').Font.Name; normal_size=$book.Styles.Item('Normal').Font.Size } | ConvertTo-Json -Compress))
            $checks[('photo_' + $side.ToLower() + '_proportions')] = [Math]::Abs(($shape.Width/$shape.Height)/$ratio - 1) -lt 0.01
        }
    }
    foreach ($key in @($checks.Keys)) {
        if ($checks[$key] -is [bool] -and -not $checks[$key]) { throw "Native Excel check failed: $key" }
    }
    # Render a synthetic section through Excel's actual print engine. Only this
    # owned read-only instance's temporary page settings are changed in memory.
    $page = $book.Worksheets.Item(2)
    $page.PageSetup.PrintArea = '$A$18:$Q$30'
    $page.PageSetup.Zoom = $false
    $page.PageSetup.FitToPagesWide = 1
    $page.PageSetup.FitToPagesTall = 1
    $pdf = Join-Path $ReportDirectory 'tools-native.pdf'
    $page.ExportAsFixedFormat(0, $pdf)
    $photos.PageSetup.PrintArea = '$A$1:$K$24'
    $photos.PageSetup.Zoom = $false
    $photos.PageSetup.FitToPagesWide = 1
    $photos.PageSetup.FitToPagesTall = 1
    $photos.ExportAsFixedFormat(0, (Join-Path $ReportDirectory 'photos-native.pdf'))
    $checks.source_unchanged = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ceq $before
    $checks | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $ReportDirectory 'native-excel.json') -Encoding utf8
    $checks | ConvertTo-Json -Compress
} finally {
    if ($book) { $book.Close($false) }
    if ($sourceBook) { $sourceBook.Close($false) }
    if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
