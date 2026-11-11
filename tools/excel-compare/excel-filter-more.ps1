# Excel's other AutoFilters, for filter-more.test.js: text and number filters (Range.AutoFilter with criteria and
# xlAnd/xlOr), Top 10 (items and percent, top and bottom), above and below average and the date filters (xlFilterDynamic),
# and filtering by a fill or font colour (xlFilterCellColor, xlFilterFontColor). Each case on a new workbook of the same
# made-up table (its dates =TODAY() and days either side, so a case reads the same on any day; today's serial is
# recorded), filtered once; the rows hidden, and from the workbook Excel then saves, its autoFilter element and its
# dxfs. An invisible instance, macros disabled for it only; no Excel setting changed, no keys sent.
# usage: pwsh tools/excel-compare/excel-filter-more.ps1 <out.json> <folder for the saved workbooks>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$TABLE = [ordered]@{
  A1 = 'Name'; B1 = 'Qty'; C1 = 'Team'; D1 = 'Date'
  A2 = 'apple'; B2 = '1'; C2 = 'red'; D2 = '=TODAY()-1'
  A3 = 'pear'; B3 = '2'; C3 = 'blue'; D3 = '=TODAY()'
  A4 = 'Apple'; B4 = '3'; C4 = 'red'; D4 = '=TODAY()+1'
  A5 = 'fig'; B5 = '1'; C5 = ''; D5 = '=TODAY()-40'
  A6 = 'plum'; B6 = '2.5'; C6 = 'green'; D6 = '=TODAY()+40'
  A7 = 'kiwi'; B7 = '10'; C7 = 'blue'; D7 = '=TODAY()-400'
  A8 = 'banana'; B8 = '7'; C8 = 'red'; D8 = '=TODAY()-3'
}
# Each case: field (1-based), then how: [criteria1, operator, criteria2] (xlAnd 1, xlOr 2, xlTop10Items 3, xlBottom10Items
# 4, xlTop10Percent 5, xlBottom10Percent 6, xlFilterDynamic 11, xlFilterCellColor 8, xlFilterFontColor 9).
$CASES = [ordered]@{
  'text equals' = @(1, '=pear'); 'begins with' = @(1, '=a*'); 'ends with' = @(1, '=*e'); 'contains' = @(1, '=*an*')
  'does not contain' = @(1, '<>*an*'); 'does not equal' = @(3, '<>red'); 'greater than' = @(2, '>2'); 'less than or equal' = @(2, '<=2')
  'between' = @(2, '>=2', 1, '<=7'); 'below or above' = @(2, '<2', 2, '>7'); 'text or' = @(1, '=a*', 2, '=k*')
  'top 3 items' = @(2, '3', 3); 'bottom 2 items' = @(2, '2', 4); 'top 50 percent' = @(2, '50', 5); 'bottom 30 percent' = @(2, '30', 6)
  'above average' = @(2, 33, 11); 'below average' = @(2, 34, 11)
  'today' = @(4, 1, 11); 'yesterday' = @(4, 2, 11); 'tomorrow' = @(4, 3, 11); 'this week' = @(4, 4, 11); 'last week' = @(4, 5, 11)
  'this month' = @(4, 7, 11); 'last month' = @(4, 8, 11); 'next month' = @(4, 9, 11); 'this year' = @(4, 13, 11); 'last year' = @(4, 14, 11)
  'year to date' = @(4, 16, 11); 'this quarter' = @(4, 10, 11)
  'fill yellow' = @(1, 0x00FFFF, 8); 'font red' = @(1, 0x0000FF, 9)
}
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$answers = New-Object System.Collections.Generic.List[object]
$today = $null
try {
  $i = 0
  foreach ($name in $CASES.Keys) {
    $i++
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    foreach ($k in $TABLE.Keys) { if ($TABLE[$k] -ne '') { $ws.Range($k).Formula = $TABLE[$k] } }
    $ws.Range('D2:D8').NumberFormat = 'm/d/yyyy'
    $ws.Range('A3').Interior.Color = 0x00FFFF; $ws.Range('A5').Interior.Color = 0x00FFFF; $ws.Range('A7').Interior.Color = 0x50D092
    $ws.Range('A4').Font.Color = 0x0000FF
    if ($null -eq $today) { $today = [double]$ws.Range('D3').Value2 }
    $c = $CASES[$name]
    $rg = $ws.Range('A1:D8')
    if ($c.Count -eq 2) { [void]$rg.AutoFilter($c[0], $c[1]) }
    elseif ($c[2] -eq 11) { [void]$rg.AutoFilter($c[0], $c[1], 11) }
    elseif ($c[2] -eq 8 -or $c[2] -eq 9) { [void]$rg.AutoFilter($c[0], $c[1], $c[2]) }
    elseif ($c.Count -eq 3) { [void]$rg.AutoFilter($c[0], $c[1], $c[2]) }
    else { [void]$rg.AutoFilter($c[0], $c[1], $c[2], $c[3]) }
    $hidden = New-Object System.Collections.Generic.List[int]
    for ($r = 2; $r -le 8; $r++) { if ($ws.Rows.Item($r).Hidden) { $hidden.Add($r) } }
    $file = Join-Path $Books "filter-more-$i.xlsx"
    if (Test-Path $file) { Remove-Item $file }
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    $read = { param($n) $s = New-Object IO.StreamReader($zip.GetEntry($n).Open()); $t = $s.ReadToEnd(); $s.Close(); $t }
    $sheet = & $read 'xl/worksheets/sheet1.xml'
    $styles = & $read 'xl/styles.xml'
    $zip.Dispose()
    $answers.Add([ordered]@{ name = $name; how = $c; hidden = $hidden; autoFilter = [regex]::Match($sheet, '<autoFilter[^>]*?(/>|>.*?</autoFilter>)').Value; dxfs = [regex]::Match($styles, '<dxfs\b[^>]*/>|<dxfs\b[\s\S]*?</dxfs>').Value })
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
[ordered]@{ source = 'tools/excel-compare/excel-filter-more.ps1'; today = $today; cases = $answers } | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 $Out
"$($answers.Count) cases"
