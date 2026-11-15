# Excel's data tools, for datatools.test.js: Fill > Series (Range.DataSeries), Remove Duplicates (Range.RemoveDuplicates),
# Text to Columns (Range.TextToColumns), Flash Fill (Range.FlashFill) and Move or Copy > Create a copy (Worksheet.Copy).
# Each case on a new workbook of made-up cells; afterwards the cells A1:H12 as Excel shows them and holds them (value
# and type), and for the sheet copy, the workbook's sheet names, defined names and the two sheets' XML from the saved
# workbook. An invisible instance; no keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-datatools.ps1 <out.json> <folder for the saved workbooks>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
function Put($ws, $cells) { foreach ($k in $cells.Keys) { $ws.Range($k).Formula = [string]$cells[$k] } }
# Each case: cells to put first, then a script block given the sheet.
$m = [Type]::Missing
$CASES = [ordered]@{
  'series linear down' = @(@{ A1 = '1' }, { param($ws) [void]$ws.Range('A1:A6').DataSeries(2, -4132, 1, 2) })
  'series linear stop' = @(@{ A1 = '5' }, { param($ws) [void]$ws.Range('A1:A10').DataSeries(2, -4132, 1, 3, 15) })
  'series growth' = @(@{ A1 = '2' }, { param($ws) [void]$ws.Range('A1:A5').DataSeries(2, 2, 1, 3) })
  'series across' = @(@{ A1 = '10' }, { param($ws) [void]$ws.Range('A1:E1').DataSeries(1, -4132, 1, -2.5) })
  'series months' = @(@{ A1 = '1/31/2026' }, { param($ws) [void]$ws.Range('A1:A5').DataSeries(2, 3, 3, 1) })
  'series weekdays' = @(@{ A1 = '11/13/2026' }, { param($ws) [void]$ws.Range('A1:A6').DataSeries(2, 3, 2, 1) })
  'series days by 7' = @(@{ A1 = '1/1/2026' }, { param($ws) [void]$ws.Range('A1:A4').DataSeries(2, 3, 1, 7) })
  'series years' = @(@{ A1 = '2/29/2024' }, { param($ws) [void]$ws.Range('A1:A3').DataSeries(2, 3, 4, 1) })
  'series trend' = @(@{ A1 = '1'; A2 = '3'; A3 = '4' }, { param($ws) [void]$ws.Range('A1:A6').DataSeries(2, -4132, 1, 1, [Type]::Missing, $true) })
  'series two columns' = @(@{ A1 = '1'; B1 = '10' }, { param($ws) [void]$ws.Range('A1:B4').DataSeries(2, -4132, 1, 5) })
  'remove duplicates all columns' = @(@{ A1 = 'Name'; B1 = 'Team'; A2 = 'pear'; B2 = 'red'; A3 = 'Pear'; B3 = 'red'; A4 = 'fig'; B4 = 'blue'; A5 = 'pear'; B5 = 'blue'; A6 = 'fig'; B6 = 'blue'; D2 = 'beside' },
    { param($ws) $ws.Range('A1:B6').RemoveDuplicates([object[]]@(1, 2), 1) })
  'remove duplicates one column' = @(@{ A1 = 'Name'; B1 = 'Team'; A2 = 'pear'; B2 = 'red'; A3 = 'fig'; B3 = 'red'; A4 = 'kiwi'; B4 = 'blue'; A5 = 'plum'; B5 = 'blue'; A6 = '1'; B6 = 'red'; A7 = '1.0'; B7 = 'red' },
    { param($ws) $ws.Range('A1:B7').RemoveDuplicates([object[]]@(2), 1) })
  'remove duplicates numbers no header' = @(@{ A1 = '1'; A2 = '2'; A3 = '1'; A4 = '3'; A5 = '2'; A6 = '1' }, { param($ws) $ws.Range('A1:A6').RemoveDuplicates([object[]]@(1), 2) })
  'text to columns comma' = @(@{ A1 = 'a,b,c'; A2 = '1,2.5,x'; A3 = 'one,,three'; A4 = 'solo' }, { param($ws) [void]$ws.Range('A1:A4').TextToColumns($ws.Range('A1'), 1, 1, $false, $false, $false, $true, $false, $false) })
  'text to columns space consecutive' = @(@{ A1 = 'John  Smith'; A2 = 'Mary Ann Lee'; A3 = ' lead' }, { param($ws) [void]$ws.Range('A1:A3').TextToColumns($ws.Range('A1'), 1, 1, $true, $false, $false, $false, $true) })
  'text to columns tab and semicolon' = @(@{ A1 = "a`tb;c"; A2 = '1;2' }, { param($ws) [void]$ws.Range('A1:A2').TextToColumns($ws.Range('A1'), 1, 1, $false, $true, $true) })
  'text to columns quoted' = @(@{ A1 = '"Smith, John",42'; A2 = 'x,"y"' }, { param($ws) [void]$ws.Range('A1:A2').TextToColumns($ws.Range('A1'), 1, 1, $false, $false, $false, $true) })
  'text to columns dates and zeros' = @(@{ A1 = '007,3/4/2026,TRUE,-5%' }, { param($ws) [void]$ws.Range('A1').TextToColumns($ws.Range('A1'), 1, 1, $false, $false, $false, $true) })
  'text to columns elsewhere' = @(@{ A1 = 'a-b'; A2 = 'c-d' }, { param($ws) [void]$ws.Range('A1:A2').TextToColumns($ws.Range('C1'), 1, 1, $false, $false, $false, $false, $false, $true, '-') })
  'text to columns fixed width' = @(@{ A1 = 'ABC12345 x'; A2 = 'DE 678'; A3 = 'XYZ0000999Y'; A4 = 'ab   cd  e'; A5 = '1234567' },
    { param($ws) [void]$ws.Range('A1:A5').TextToColumns($ws.Range('A1'), 2, 1, $m, $m, $m, $m, $m, $m, $m, [object[]]@([object[]]@(0, 1), [object[]]@(3, 1), [object[]]@(8, 1))) })
  'text to columns fixed width types' = @(@{ A1 = '2026-01-05 12.5  TRUE'; A2 = '3/4/2026  007   -5%' },
    { param($ws) [void]$ws.Range('A1:A2').TextToColumns($ws.Range('A1'), 2, 1, $m, $m, $m, $m, $m, $m, $m, [object[]]@([object[]]@(0, 1), [object[]]@(10, 1), [object[]]@(16, 1))) })
  'text to columns fixed width elsewhere' = @(@{ A1 = 'abc'; A2 = 'abcdefghij' },
    { param($ws) [void]$ws.Range('A1:A2').TextToColumns($ws.Range('C1'), 2, 1, $m, $m, $m, $m, $m, $m, $m, [object[]]@([object[]]@(0, 1), [object[]]@(2, 1), [object[]]@(5, 1), [object[]]@(8, 1))) })
  'text to columns column formats' = @(@{ A1 = '007,3/4/2026,gone,12.50,4/3/26'; A2 = '0042,13/4/2026,x,1e3,2026-03-04' },
    { param($ws) [void]$ws.Range('A1:A2').TextToColumns($ws.Range('A1'), 1, 1, $false, $false, $false, $true, $false, $false, $m, [object[]]@([object[]]@(1, 2), [object[]]@(2, 4), [object[]]@(3, 9), [object[]]@(4, 1), [object[]]@(5, 3))) })
  'text to columns fixed width formats' = @(@{ A1 = '007 2026/03/04 x  1/2/03'; A2 = '12  26-3-4     y  12/31/99' },
    { param($ws) [void]$ws.Range('A1:A2').TextToColumns($ws.Range('A1'), 2, 1, $m, $m, $m, $m, $m, $m, $m, [object[]]@([object[]]@(0, 2), [object[]]@(4, 5), [object[]]@(15, 9), [object[]]@(18, 3))) })
  'flash fill first names' = @(@{ A1 = 'Full'; B1 = 'First'; A2 = 'Avery Chen'; B2 = 'Avery'; A3 = 'Sam Lee'; A4 = 'Jo Park'; A5 = 'Rin Ito' }, { param($ws) [void]$ws.Range('B2:B5').FlashFill() })
  'flash fill initials' = @(@{ A1 = 'Full'; B1 = 'Short'; A2 = 'Avery Chen'; B2 = 'A. Chen'; A3 = 'Sam Lee'; A4 = 'Jo Park' }, { param($ws) [void]$ws.Range('B2:B4').FlashFill() })
  'flash fill upper' = @(@{ A1 = 'Code'; B1 = 'Up'; A2 = 'ab-12'; B2 = 'AB-12'; A3 = 'cd-34'; A4 = 'ef-56' }, { param($ws) [void]$ws.Range('B2:B4').FlashFill() })
  'flash fill join' = @(@{ A1 = 'First'; B1 = 'Last'; C1 = 'Both'; A2 = 'Avery'; B2 = 'Chen'; C2 = 'Chen, Avery'; A3 = 'Sam'; B3 = 'Lee'; A4 = 'Jo'; B4 = 'Park' }, { param($ws) [void]$ws.Range('C2:C4').FlashFill() })
  'flash fill digits' = @(@{ A1 = 'Ref'; B1 = 'No'; A2 = 'INV-0042-X'; B2 = '0042'; A3 = 'INV-0107-Y'; A4 = 'INV-2300-Z' }, { param($ws) [void]$ws.Range('B2:B4').FlashFill() })
}
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$answers = New-Object System.Collections.Generic.List[object]
try {
  foreach ($name in $CASES.Keys) {
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $spec = $CASES[$name]
    Put $ws $spec[0]
    $err = $null
    try { & $spec[1] $ws } catch { $err = [string]$_.Exception.Message }
    $cells = [ordered]@{}
    for ($r = 1; $r -le 12; $r++) { for ($c = 1; $c -le 8; $c++) { $cell = $ws.Cells.Item($r, $c); if ($null -ne $cell.Value2 -and [string]$cell.Formula -ne '') { $cells[$cell.Address($false, $false)] = @([string]$cell.Text, [string]$cell.Value2.GetType().Name, [string]$cell.NumberFormat, [string]$cell.Value2) } } }
    $answers.Add([ordered]@{ name = $name; error = $err; cells = $cells })
    $wb.Close($false)
  }
  # Move or Copy > Create a copy: a sheet with cells, a format, a column width, a merge, a note, a local name and a
  # conditional format; copied after itself.
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'Data'
  Put $ws @{ A1 = 'Item'; B1 = 'Qty'; A2 = 'pear'; B2 = '3'; A3 = 'fig'; B3 = '=B2*2' }
  $ws.Range('A1:B1').Font.Bold = $true
  $ws.Columns.Item(1).ColumnWidth = 15
  $ws.Range('D1:E1').Merge()
  [void]$ws.Names.Add('Local', '=Data!$B$2')
  [void]$wb.Names.Add('Everywhere', '=Data!$A$1')
  [void]$ws.Range('B2:B3').FormatConditions.Add(1, 5, '4')
  $ws.Copy([Type]::Missing, $ws)
  $file = Join-Path $Books 'datatools-copy.xlsx'
  if (Test-Path $file) { Remove-Item $file }
  $wb.SaveAs($file, 51)
  $names = @(); foreach ($s in $wb.Worksheets) { $names += $s.Name }
  $wb.Close($false)
  $zip = [IO.Compression.ZipFile]::OpenRead($file)
  $read = { param($n) $s = New-Object IO.StreamReader($zip.GetEntry($n).Open()); $t = $s.ReadToEnd(); $s.Close(); $t }
  $book = & $read 'xl/workbook.xml'
  $s1 = & $read 'xl/worksheets/sheet1.xml'
  $s2 = & $read 'xl/worksheets/sheet2.xml'
  $zip.Dispose()
  $copy = [ordered]@{ sheets = $names; book = $book; sheet1 = $s1; sheet2 = $s2 }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
[ordered]@{ source = 'tools/excel-compare/excel-datatools.ps1'; cases = $answers; copy = $copy } | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 $Out
"$($answers.Count) cases"
