# What Excel writes for outlines, for outline.test.js: Group and Ungroup (Range.Group, Ungroup), Show Levels
# (Outline.ShowLevels), Show Detail and Hide Detail (Range.ShowDetail), Clear Outline (Range.ClearOutline), Auto
# Outline (Range.AutoOutline), the Settings (Outline.SummaryRow, SummaryColumn), Subtotal (Range.Subtotal) and Remove
# All (Range.RemoveSubtotal). Each case on a new workbook of made-up cells in an invisible Excel, saved to a folder
# outside the repository and its sheet's sheetPr, sheetFormatPr, cols, rows (with their cells) and rowBreaks read back.
# Then the outline bar's room: where the cells start on screen (Window.PointsToScreenPixelsX/Y) with 1 to 8 levels of
# rows and columns, against no outline. No keys sent, no setting changed; only this script's own Excel is quit.
# usage: pwsh tools/excel-compare/excel-outline.ps1 <out.json> <folder for the saved workbooks>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
trap { Write-Host "LINE $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; break }
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$m = [Type]::Missing
# (Made-up cells: numbers down A1:A12 and across A1:H1.)
function Numbers($ws) { for ($r = 1; $r -le 12; $r++) { $ws.Cells.Item($r, 1).Value2 = $r * 10 }; for ($c = 2; $c -le 8; $c++) { $ws.Cells.Item(1, $c).Value2 = $c } }
# (A list for Subtotal: Region, Item, Units, Price; sorted by Region as Subtotal expects.)
function List($ws) {
  $rows = 'Region,Item,Units,Price', 'East,pear,4,1.5', 'East,fig,2,3', 'North,kiwi,7,0.5', 'West,plum,1,2.25', 'West,pear,5,1.5', 'West,lime,3,0.75'
  $r = 0
  foreach ($line in $rows) { $r++; $c = 0; foreach ($v in $line.Split(',')) { $c++; $ws.Range("$([char](64 + $c))$r").Formula = $v } }
}
# (Totals by hand for Auto Outline: SUM rows under their details and a SUM column right of its details.)
function Sums($ws) {
  for ($r = 1; $r -le 3; $r++) { for ($c = 1; $c -le 2; $c++) { $ws.Cells.Item($r, $c).Value2 = $r + $c } }
  $ws.Range('A4').Formula = '=SUM(A1:A3)'; $ws.Range('B4').Formula = '=SUM(B1:B3)'
  for ($r = 1; $r -le 4; $r++) { $ws.Cells.Item($r, 3).Formula = "=SUM(A$($r):B$($r))" }
}
$CASES = [ordered]@{
  'group rows 2-4' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A4').EntireRow.Group() }
  'group rows twice nested' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A7').EntireRow.Group(); [void]$ws.Range('A3:A4').EntireRow.Group() }
  'group rows three levels' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A9').EntireRow.Group(); [void]$ws.Range('A3:A7').EntireRow.Group(); [void]$ws.Range('A4:A5').EntireRow.Group() }
  'group rows two groups side by side' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A3').EntireRow.Group(); [void]$ws.Range('A5:A6').EntireRow.Group() }
  'group rows touching' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A3').EntireRow.Group(); [void]$ws.Range('A4:A5').EntireRow.Group() }
  'group empty rows past the data' = { param($ws) Numbers $ws; [void]$ws.Range('A20:A22').EntireRow.Group() }
  'group rows then collapse level 1' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A4').EntireRow.Group(); [void]$ws.Outline.ShowLevels(1) }
  'nested collapse level 1' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A7').EntireRow.Group(); [void]$ws.Range('A3:A4').EntireRow.Group(); [void]$ws.Outline.ShowLevels(1) }
  'nested collapse level 2' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A7').EntireRow.Group(); [void]$ws.Range('A3:A4').EntireRow.Group(); [void]$ws.Outline.ShowLevels(2) }
  'nested level 1 then level 3' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A7').EntireRow.Group(); [void]$ws.Range('A3:A4').EntireRow.Group(); [void]$ws.Outline.ShowLevels(1); [void]$ws.Outline.ShowLevels(3) }
  'hide detail of inner group' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A7').EntireRow.Group(); [void]$ws.Range('A3:A4').EntireRow.Group(); $ws.Rows.Item(5).ShowDetail = $false }
  'hide inner then outer then show outer' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A7').EntireRow.Group(); [void]$ws.Range('A3:A4').EntireRow.Group(); $ws.Rows.Item(5).ShowDetail = $false; $ws.Rows.Item(8).ShowDetail = $false; $ws.Rows.Item(8).ShowDetail = $true }
  'hide outer then show inner' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A7').EntireRow.Group(); [void]$ws.Range('A3:A4').EntireRow.Group(); $ws.Rows.Item(8).ShowDetail = $false; $ws.Rows.Item(5).ShowDetail = $true }
  'hide detail then show detail' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A4').EntireRow.Group(); $ws.Rows.Item(5).ShowDetail = $false; $ws.Rows.Item(5).ShowDetail = $true }
  'hide detail from a detail row' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A4').EntireRow.Group(); $ws.Rows.Item(3).ShowDetail = $false }
  'ungroup inner' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A7').EntireRow.Group(); [void]$ws.Range('A3:A4').EntireRow.Group(); [void]$ws.Range('A3:A4').EntireRow.Ungroup() }
  'ungroup part of a group' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A6').EntireRow.Group(); [void]$ws.Range('A3:A4').EntireRow.Ungroup() }
  'ungroup a collapsed group' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A4').EntireRow.Group(); $ws.Rows.Item(5).ShowDetail = $false; [void]$ws.Range('A2:A4').EntireRow.Ungroup() }
  'ungroup rows not grouped' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A4').EntireRow.Group(); try { [void]$ws.Range('A8:A9').EntireRow.Ungroup() } catch { $ws.Range('J1').Value2 = 'refused' } }
  'summary above' = { param($ws) Numbers $ws; $ws.Outline.SummaryRow = 0; [void]$ws.Range('A3:A5').EntireRow.Group() }
  'summary above collapse' = { param($ws) Numbers $ws; $ws.Outline.SummaryRow = 0; [void]$ws.Range('A3:A5').EntireRow.Group(); [void]$ws.Outline.ShowLevels(1) }
  'summary left columns' = { param($ws) Numbers $ws; $ws.Outline.SummaryColumn = -4131; [void]$ws.Range('C1:D1').EntireColumn.Group(); [void]$ws.Outline.ShowLevels($m, 1) }
  'summary below and right set again' = { param($ws) Numbers $ws; $ws.Outline.SummaryRow = 0; $ws.Outline.SummaryRow = 1; [void]$ws.Range('A3:A5').EntireRow.Group() }
  'group columns B-D' = { param($ws) Numbers $ws; [void]$ws.Range('B1:D1').EntireColumn.Group() }
  'group columns collapse' = { param($ws) Numbers $ws; [void]$ws.Range('B1:D1').EntireColumn.Group(); [void]$ws.Outline.ShowLevels($m, 1) }
  'group columns nested hide inner' = { param($ws) Numbers $ws; [void]$ws.Range('B1:F1').EntireColumn.Group(); [void]$ws.Range('C1:D1').EntireColumn.Group(); $ws.Columns.Item(5).ShowDetail = $false }
  'group columns with widths' = { param($ws) Numbers $ws; $ws.Columns.Item(3).ColumnWidth = 20; [void]$ws.Range('B1:D1').EntireColumn.Group(); [void]$ws.Outline.ShowLevels($m, 1) }
  'group rows with a height' = { param($ws) Numbers $ws; $ws.Rows.Item(3).RowHeight = 30; [void]$ws.Range('A2:A4').EntireRow.Group(); [void]$ws.Outline.ShowLevels(1) }
  'group rows already hidden' = { param($ws) Numbers $ws; $ws.Rows.Item(3).Hidden = $true; [void]$ws.Range('A2:A4').EntireRow.Group() }
  'eight levels' = { param($ws) Numbers $ws; for ($k = 0; $k -lt 8; $k++) { [void]$ws.Range("A$(2 + $k):A$(20 - $k)").EntireRow.Group() } }
  'ninth level refused' = { param($ws) Numbers $ws; for ($k = 0; $k -lt 8; $k++) { [void]$ws.Range("A$(2 + $k):A$(20 - $k)").EntireRow.Group() }; try { [void]$ws.Range('A10:A11').EntireRow.Group() } catch { $ws.Range('J1').Value2 = 'refused' } }
  'clear outline' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A7').EntireRow.Group(); [void]$ws.Range('B1:C1').EntireColumn.Group(); [void]$ws.Outline.ShowLevels(1, 1); [void]$ws.Cells.ClearOutline() }
  'clear outline of part' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A7').EntireRow.Group(); [void]$ws.Range('A9:A10').EntireRow.Group(); [void]$ws.Range('A2:A7').EntireRow.ClearOutline() }
  # (Ctrl+8 is the window's DisplayOutline: with an outline, without one, and turned off then on again.)
  'outline symbols off' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A4').EntireRow.Group(); $ws.Parent.Windows.Item(1).DisplayOutline = $false }
  'outline symbols off without an outline' = { param($ws) Numbers $ws; $ws.Parent.Windows.Item(1).DisplayOutline = $false }
  'outline symbols off then on' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A4').EntireRow.Group(); $ws.Parent.Windows.Item(1).DisplayOutline = $false; $ws.Parent.Windows.Item(1).DisplayOutline = $true }
  'auto outline' = { param($ws) Sums $ws; [void]$ws.Range('A1').AutoOutline() }
  # (Rows and columns inserted (Range.Insert on whole lines) inside, at the edges of and next to a group.)
  'insert row inside a group' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A4').EntireRow.Group(); [void]$ws.Rows.Item(3).Insert() }
  'insert row at a group''s first row' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A4').EntireRow.Group(); [void]$ws.Rows.Item(2).Insert() }
  'insert row at the summary row' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A4').EntireRow.Group(); [void]$ws.Rows.Item(5).Insert() }
  'insert row under the summary row' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A4').EntireRow.Group(); [void]$ws.Rows.Item(6).Insert() }
  'insert two rows in nested groups' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A7').EntireRow.Group(); [void]$ws.Range('A3:A4').EntireRow.Group(); [void]$ws.Range('A4:A5').EntireRow.Insert() }
  'insert row at the inner summary' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A7').EntireRow.Group(); [void]$ws.Range('A3:A4').EntireRow.Group(); [void]$ws.Rows.Item(5).Insert() }
  'insert row in a collapsed group' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A4').EntireRow.Group(); [void]$ws.Outline.ShowLevels(1); [void]$ws.Rows.Item(3).Insert() }
  'insert row at a collapsed summary' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A4').EntireRow.Group(); [void]$ws.Outline.ShowLevels(1); [void]$ws.Rows.Item(5).Insert() }
  'insert row summary above' = { param($ws) Numbers $ws; $ws.Outline.SummaryRow = 0; [void]$ws.Range('A3:A5').EntireRow.Group(); [void]$ws.Rows.Item(3).Insert() }
  'insert row summary above after the group' = { param($ws) Numbers $ws; $ws.Outline.SummaryRow = 0; [void]$ws.Range('A3:A5').EntireRow.Group(); [void]$ws.Rows.Item(6).Insert() }
  'insert column inside a group' = { param($ws) Numbers $ws; [void]$ws.Range('B1:D1').EntireColumn.Group(); [void]$ws.Columns.Item(3).Insert() }
  'insert column at a group''s first column' = { param($ws) Numbers $ws; [void]$ws.Range('B1:D1').EntireColumn.Group(); [void]$ws.Columns.Item(2).Insert() }
  'insert column at the summary column' = { param($ws) Numbers $ws; [void]$ws.Range('B1:D1').EntireColumn.Group(); [void]$ws.Columns.Item(5).Insert() }
  # (The Settings box's Automatic styles (Outline.AutomaticStyles), its Apply Styles (Range.ApplyOutlineStyles) and
  # Create (AutoOutline with the styles on), and Group and Subtotal with them on.)
  'automatic styles on' = { param($ws) Numbers $ws; $ws.Outline.AutomaticStyles = $true }
  'automatic styles on then off' = { param($ws) Numbers $ws; $ws.Outline.AutomaticStyles = $true; $ws.Outline.AutomaticStyles = $false }
  'group with automatic styles' = { param($ws) Numbers $ws; $ws.Outline.AutomaticStyles = $true; [void]$ws.Range('A2:A4').EntireRow.Group() }
  'apply styles two levels' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A7').EntireRow.Group(); [void]$ws.Range('A3:A4').EntireRow.Group(); [void]$ws.Range('B1:D1').EntireColumn.Group(); [void]$ws.Range('A1:H12').ApplyOutlineStyles() }
  'apply styles three levels' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A9').EntireRow.Group(); [void]$ws.Range('A3:A7').EntireRow.Group(); [void]$ws.Range('A4:A5').EntireRow.Group(); [void]$ws.Range('A1:H12').ApplyOutlineStyles() }
  'apply styles to one cell' = { param($ws) Numbers $ws; [void]$ws.Range('A2:A4').EntireRow.Group(); [void]$ws.Range('A5').ApplyOutlineStyles() }
  'apply styles summary above' = { param($ws) Numbers $ws; $ws.Outline.SummaryRow = 0; [void]$ws.Range('A3:A5').EntireRow.Group(); [void]$ws.Range('A1:H12').ApplyOutlineStyles() }
  'auto outline with automatic styles' = { param($ws) Sums $ws; $ws.Outline.AutomaticStyles = $true; [void]$ws.Range('A1').AutoOutline() }
  'subtotal with automatic styles' = { param($ws) List $ws; $ws.Outline.AutomaticStyles = $true; [void]$ws.Range('A1:D7').Subtotal(1, -4157, [object[]]@(3), $true, $false, 1) }
  'subtotal sum units' = { param($ws) List $ws; [void]$ws.Range('A1:D7').Subtotal(1, -4157, [object[]]@(3), $true, $false, 1) }
  'subtotal count two columns' = { param($ws) List $ws; [void]$ws.Range('A1:D7').Subtotal(1, -4112, [object[]]@(3, 4), $true, $false, 1) }
  'subtotal average above' = { param($ws) List $ws; [void]$ws.Range('A1:D7').Subtotal(1, -4106, [object[]]@(3), $true, $false, 0) }
  'subtotal page breaks' = { param($ws) List $ws; [void]$ws.Range('A1:D7').Subtotal(1, -4157, [object[]]@(3), $true, $true, 1) }
  'subtotal max then add min' = { param($ws) List $ws; [void]$ws.Range('A1:D7').Subtotal(1, -4136, [object[]]@(3), $true, $false, 1); [void]$ws.Range('A1:D11').Subtotal(1, -4139, [object[]]@(3), $false, $false, 1) }
  'subtotal replace' = { param($ws) List $ws; [void]$ws.Range('A1:D7').Subtotal(1, -4136, [object[]]@(3), $true, $false, 1); [void]$ws.Range('A1:D11').Subtotal(1, -4157, [object[]]@(4), $true, $false, 1) }
  'subtotal at item' = { param($ws) List $ws; [void]$ws.Range('A1:D7').Subtotal(2, -4157, [object[]]@(3), $true, $false, 1) }
  'subtotal from one cell' = { param($ws) List $ws; [void]$ws.Range('B3').CurrentRegion.Subtotal(1, -4157, [object[]]@(3), $true, $false, 1) }
  'subtotal functions' = { param($ws) List $ws; [void]$ws.Range('A1:D7').Subtotal(1, -4149, [object[]]@(3), $true, $false, 1); [void]$ws.Range('A1:D11').Subtotal(1, -4113, [object[]]@(4), $false, $false, 1) }
  'subtotal stdev var' = { param($ws) List $ws; [void]$ws.Range('A1:D7').Subtotal(1, -4155, [object[]]@(3), $true, $false, 1); [void]$ws.Range('A1:D11').Subtotal(1, -4156, [object[]]@(3), $false, $false, 1); [void]$ws.Range('A1:D15').Subtotal(1, -4164, [object[]]@(3), $false, $false, 1); [void]$ws.Range('A1:D19').Subtotal(1, -4165, [object[]]@(3), $false, $false, 1) }
  'remove subtotals' = { param($ws) List $ws; [void]$ws.Range('A1:D7').Subtotal(1, -4157, [object[]]@(3), $true, $true, 1); [void]$ws.Range('A1:D11').RemoveSubtotal() }
  'subtotal collapse level 2' = { param($ws) List $ws; [void]$ws.Range('A1:D7').Subtotal(1, -4157, [object[]]@(3), $true, $false, 1); [void]$ws.Outline.ShowLevels(2) }
}
function Grab($file) {
  $zip = [IO.Compression.ZipFile]::OpenRead($file)
  $s = New-Object IO.StreamReader($zip.GetEntry('xl/worksheets/sheet1.xml').Open()); $sheet = $s.ReadToEnd(); $s.Close()
  $strings = @()
  $e = $zip.GetEntry('xl/sharedStrings.xml')
  if ($e) { $s = New-Object IO.StreamReader($e.Open()); $sst = $s.ReadToEnd(); $s.Close(); $strings = @([regex]::Matches($sst, '<t\b[^>]*>([^<]*)</t>') | ForEach-Object { $_.Groups[1].Value }) }
  $s = New-Object IO.StreamReader($zip.GetEntry('xl/styles.xml').Open()); $styles = $s.ReadToEnd(); $s.Close()
  $zip.Dispose()
  return [ordered]@{
    strings = $strings
    fonts = ([regex]::Match($styles, '<fonts\b[\s\S]*?</fonts>')).Value
    cellXfs = ([regex]::Match($styles, '<cellXfs\b[\s\S]*?</cellXfs>')).Value
    cellStyleXfs = ([regex]::Match($styles, '<cellStyleXfs\b[\s\S]*?</cellStyleXfs>')).Value
    cellStyles = ([regex]::Match($styles, '<cellStyles\b[\s\S]*?</cellStyles>')).Value
    sheetPr = ([regex]::Match($sheet, '<sheetPr\b[^>]*/>|<sheetPr\b[\s\S]*?</sheetPr>')).Value
    sheetFormatPr = ([regex]::Match($sheet, '<sheetFormatPr\b[^>]*/>')).Value
    sheetView = ([regex]::Match($sheet, '<sheetView\b[^>]*>')).Value
    cols = ([regex]::Match($sheet, '<cols>[\s\S]*?</cols>')).Value
    sheetData = ([regex]::Match($sheet, '<sheetData>[\s\S]*?</sheetData>|<sheetData/>')).Value
    rowBreaks = ([regex]::Match($sheet, '<rowBreaks\b[\s\S]*?</rowBreaks>')).Value
  }
}
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$results = @()
$bar = @()
try {
  $k = 0
  foreach ($name in $CASES.Keys) {
    $k++
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $err = $null
    try { & $CASES[$name] $ws } catch { $err = $_.Exception.Message }
    $file = Join-Path $Books "outline-$k.xlsx"
    if (Test-Path $file) { Remove-Item $file }
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $got = Grab $file
    $got.name = $name
    $got.error = $err
    $results += [pscustomobject]$got
  }
  # The outline bar's room: where the cells start on screen with 0 to 8 levels of rows (x) and columns (y).
  $x.ScreenUpdating = $true
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $win = $wb.Windows.Item(1)
  $win.Zoom = 100
  $base = [ordered]@{ x = $win.PointsToScreenPixelsX(0); y = $win.PointsToScreenPixelsY(0) }
  for ($k = 1; $k -le 7; $k++) {
    [void]$ws.Range("A$(2 + $k):A$(40 - $k)").EntireRow.Group()
    [void]$ws.Range("$([char](66 + $k))1:$([char](80 - $k))1").EntireColumn.Group()
    $bar += [pscustomobject]@{ levels = $k + 1; x = $win.PointsToScreenPixelsX(0) - $base.x; y = $win.PointsToScreenPixelsY(0) - $base.y }
  }
  $win.DisplayOutline = $false
  $bar += [pscustomobject]@{ levels = 'hidden'; x = $win.PointsToScreenPixelsX(0) - $base.x; y = $win.PointsToScreenPixelsY(0) - $base.y }
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x); [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
$json = [ordered]@{ cases = $results; bar = $bar } | ConvertTo-Json -Depth 5
[IO.File]::WriteAllText($Out, $json, (New-Object Text.UTF8Encoding($false)))
"$($results.Count) cases"
