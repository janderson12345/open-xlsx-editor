# What Excel writes for Format > Default Width..., a row height given to the whole sheet (all selected), and AutoFit
# of wrapped and merged cells: each case on a new workbook in an invisible Excel, saved to a folder outside the
# repository, its sheetFormatPr, cols and rows read back (made-up cells; no user name in them). No keys sent, no
# setting changed.
# usage: pwsh tools/excel-compare/excel-sizes2.ps1 <out.json> <folder for the saved workbooks>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
trap { Write-Host "LINE $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; break }
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$long = 'The quick brown fox jumps over the lazy dog again and again'
$CASES = [ordered]@{
  'default width 12' = { param($ws) $ws.Range('A1').Value2 = 'x'; $ws.Columns.Item(3).ColumnWidth = 20; $ws.StandardWidth = 12 }
  'default width 4.5' = { param($ws) $ws.Range('A1').Value2 = 'x'; $ws.StandardWidth = 4.5 }
  'whole sheet row height 21' = { param($ws) $ws.Range('A1').Value2 = 'x'; $ws.Range('A3').Value2 = 'y'; $ws.Cells.RowHeight = 21 }
  'whole sheet row height then one row' = { param($ws) $ws.Range('A1').Value2 = 'x'; $ws.Cells.RowHeight = 21; $ws.Rows.Item(2).RowHeight = 30 }
  'autofit column with a merged cell' = { param($ws) $ws.Range('A1').Value2 = $long; $ws.Range('A1:C1').Merge(); $ws.Range('A2').Value2 = 'short'; [void]$ws.Columns.Item(1).AutoFit() }
  'autofit column with wrapped text' = { param($ws) $ws.Range('A1').Value2 = $long; $ws.Range('A1').WrapText = $true; $ws.Range('A2').Value2 = 'short'; [void]$ws.Columns.Item(1).AutoFit() }
  'autofit column with a wrapped line break' = { param($ws) $ws.Range('A1').Value2 = "first line`nsecond longer line here"; $ws.Range('A1').WrapText = $true; [void]$ws.Columns.Item(1).AutoFit() }
  'autofit row with wrapped text' = { param($ws) $ws.Range('A1').Value2 = $long; $ws.Range('A1').WrapText = $true; $ws.Rows.Item(1).RowHeight = 15; [void]$ws.Rows.Item(1).AutoFit() }
  'autofit row with a merged wrapped cell' = { param($ws) $ws.Range('A1').Value2 = $long; $ws.Range('A1:B1').Merge(); $ws.Range('A1').WrapText = $true; $ws.Rows.Item(1).RowHeight = 15; [void]$ws.Rows.Item(1).AutoFit() }
  'autofit row with a merged wrapped cell and a wrapped one' = { param($ws) $ws.Range('A1').Value2 = $long; $ws.Range('A1:B1').Merge(); $ws.Range('A1').WrapText = $true; $ws.Range('D1').Value2 = 'one two three four'; $ws.Range('D1').WrapText = $true; $ws.Rows.Item(1).RowHeight = 15; [void]$ws.Rows.Item(1).AutoFit() }
}
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$results = @()
try {
  $k = 0
  foreach ($name in $CASES.Keys) {
    $k++
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    & $CASES[$name] $ws
    $seen = [ordered]@{ A = $ws.Columns.Item(1).ColumnWidth; B = $ws.Columns.Item(2).ColumnWidth; D = $ws.Columns.Item(4).ColumnWidth; row1 = $ws.Rows.Item(1).RowHeight; row2 = $ws.Rows.Item(2).RowHeight; row5 = $ws.Rows.Item(5).RowHeight; standard = $ws.StandardWidth; standardHeight = $ws.StandardHeight }
    $file = Join-Path $Books "sizes2-$k.xlsx"
    if (Test-Path $file) { Remove-Item $file }
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    $s = New-Object IO.StreamReader($zip.GetEntry('xl/worksheets/sheet1.xml').Open()); $sheet = $s.ReadToEnd(); $s.Close()
    $zip.Dispose()
    $fmt = ([regex]::Match($sheet, '<sheetFormatPr\b[^>]*/>')).Value
    $cols = ([regex]::Match($sheet, '<cols>[\s\S]*?</cols>')).Value
    $rows = ([regex]::Matches($sheet, '<row\b[^>]*>') | ForEach-Object { $_.Value }) -join ''
    $results += [pscustomobject]@{ name = $name; sheetFormatPr = $fmt; cols = $cols; rows = $rows; seen = $seen }
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x); [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
$results | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 $Out
"$($results.Count) cases"
