# What Excel writes for Freeze Panes in the cases docs/VIEW.md had as not yet: frozen while scrolled, with A1 active,
# Freeze Top Row and First Column while scrolled, with split panes there first, over hidden rows and columns, Unfreeze
# afterwards, and the scrolled pane's place saved. Each case on a new workbook in an invisible Excel: the window
# scrolled (ScrollRow, ScrollColumn) or split (SplitRow, SplitColumn), a cell selected, then Excel's own ribbon command
# (CommandBars.ExecuteMso FreezePanes, FreezeTopRow, FreezeFirstColumn) or Window.FreezePanes; saved to a folder
# outside the repository and the sheet's sheetView element read back, with the window's rows and columns in view.
# No keys sent, no setting saved.
# usage: pwsh tools/excel-compare/excel-freeze.ps1 <out.json> <folder for the saved books>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
trap { Write-Host "LINE $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; break }
Add-Type -AssemblyName System.IO.Compression.FileSystem
function ViewXml($path) {
  $zip = [IO.Compression.ZipFile]::OpenRead($path)
  try {
    $e = $zip.Entries | Where-Object FullName -eq 'xl/worksheets/sheet1.xml'
    $text = (New-Object IO.StreamReader($e.Open())).ReadToEnd()
    return [regex]::Match($text, '<sheetViews>[\s\S]*?</sheetViews>').Value
  } finally { $zip.Dispose() }
}
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$results = @()
try {
  $cases = [ordered]@{
    'C5, not scrolled' = { param($w, $ws) $ws.Range('C5').Select(); $x.CommandBars.ExecuteMso('FreezePanes') }
    'E15, scrolled to row 10 and column C' = { param($w, $ws) $w.ScrollRow = 10; $w.ScrollColumn = 3; $ws.Range('E15').Select(); $x.CommandBars.ExecuteMso('FreezePanes') }
    'B3, scrolled to row 10 and column C (active cell out of view)' = { param($w, $ws) $ws.Range('B3').Select(); $w.ScrollRow = 10; $w.ScrollColumn = 3; $x.CommandBars.ExecuteMso('FreezePanes') }
    'A1, not scrolled' = { param($w, $ws) $ws.Range('A1').Select(); $x.CommandBars.ExecuteMso('FreezePanes') }
    'C1, not scrolled' = { param($w, $ws) $ws.Range('C1').Select(); $x.CommandBars.ExecuteMso('FreezePanes') }
    'A5, not scrolled' = { param($w, $ws) $ws.Range('A5').Select(); $x.CommandBars.ExecuteMso('FreezePanes') }
    'A10 active, scrolled to row 10' = { param($w, $ws) $ws.Range('A10').Select(); $w.ScrollRow = 10; $x.CommandBars.ExecuteMso('FreezePanes') }
    'Top Row, scrolled to row 10' = { param($w, $ws) $ws.Range('D20').Select(); $w.ScrollRow = 10; $x.CommandBars.ExecuteMso('FreezeTopRow') }
    'Top Row, scrolled to row 10 and column D' = { param($w, $ws) $ws.Range('F20').Select(); $w.ScrollRow = 10; $w.ScrollColumn = 4; $x.CommandBars.ExecuteMso('FreezeTopRow') }
    'First Column, scrolled to row 10 and column D' = { param($w, $ws) $ws.Range('F20').Select(); $w.ScrollRow = 10; $w.ScrollColumn = 4; $x.CommandBars.ExecuteMso('FreezeFirstColumn') }
    'columns C to D hidden, D5 (active column hidden)' = { param($w, $ws) $ws.Range('D5').Select(); $ws.Columns('C:D').Hidden = $true; $x.CommandBars.ExecuteMso('FreezePanes') }
    'First Column, scrolled to column D' ={ param($w, $ws) $ws.Range('F3').Select(); $w.ScrollColumn = 4; $x.CommandBars.ExecuteMso('FreezeFirstColumn') }
    'split at row 3 and column 2, then Freeze Panes (ribbon), C5 active' = { param($w, $ws) $w.SplitRow = 3; $w.SplitColumn = 2; $ws.Range('E8').Select(); $x.CommandBars.ExecuteMso('FreezePanes') }
    'split at row 3 and column 2, then Window.FreezePanes' = { param($w, $ws) $w.SplitRow = 3; $w.SplitColumn = 2; $ws.Range('E8').Select(); $w.FreezePanes = $true }
    'split at row 3 and column 2, saved split' = { param($w, $ws) $w.SplitRow = 3; $w.SplitColumn = 2; $ws.Range('E8').Select() }
    'split, then Top Row' = { param($w, $ws) $w.SplitRow = 3; $w.SplitColumn = 2; $ws.Range('E8').Select(); $x.CommandBars.ExecuteMso('FreezeTopRow') }
    'rows 2 to 3 hidden, C5' = { param($w, $ws) $ws.Rows('2:3').Hidden = $true; $ws.Range('C5').Select(); $x.CommandBars.ExecuteMso('FreezePanes') }
    'columns B to C hidden, E5' = { param($w, $ws) $ws.Columns('B:C').Hidden = $true; $ws.Range('E5').Select(); $x.CommandBars.ExecuteMso('FreezePanes') }
    'rows 4 to 6 hidden, C5 (active cell hidden)' = { param($w, $ws) $ws.Range('C5').Select(); $ws.Rows('4:6').Hidden = $true; $x.CommandBars.ExecuteMso('FreezePanes') }
    'row 1 hidden, Top Row' = { param($w, $ws) $ws.Rows('1:1').Hidden = $true; $ws.Range('C5').Select(); $x.CommandBars.ExecuteMso('FreezeTopRow') }
    'C5 frozen, then the bottom pane scrolled to row 40 and column H' = { param($w, $ws) $ws.Range('C5').Select(); $x.CommandBars.ExecuteMso('FreezePanes'); $w.ScrollRow = 40; $w.ScrollColumn = 8 }
    'E15 frozen scrolled, then unfrozen' = { param($w, $ws) $w.ScrollRow = 10; $w.ScrollColumn = 3; $ws.Range('E15').Select(); $x.CommandBars.ExecuteMso('FreezePanes'); $w.FreezePanes = $false }
    'C5 frozen, bottom pane scrolled to row 40, then unfrozen' = { param($w, $ws) $ws.Range('C5').Select(); $x.CommandBars.ExecuteMso('FreezePanes'); $w.ScrollRow = 40; $w.FreezePanes = $false }
    'C5 frozen, then Freeze Panes again (the menu turned to Unfreeze)' = { param($w, $ws) $ws.Range('C5').Select(); $x.CommandBars.ExecuteMso('FreezePanes'); $x.CommandBars.ExecuteMso('FreezePanes') }
    'C5 frozen, then Top Row' = { param($w, $ws) $ws.Range('C5').Select(); $x.CommandBars.ExecuteMso('FreezePanes'); $x.CommandBars.ExecuteMso('FreezeTopRow') }
    'C5 frozen, then E8 selected and Freeze Panes on the frozen sheet with Window.FreezePanes' = { param($w, $ws) $ws.Range('C5').Select(); $x.CommandBars.ExecuteMso('FreezePanes'); $ws.Range('E8').Select(); $w.FreezePanes = $false; $w.FreezePanes = $true }
  }
  $n = 0
  foreach ($name in $cases.Keys) {
    $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $w = $x.ActiveWindow
    Write-Host "case: $name"
    $before = $w.VisibleRange.Address($false, $false)
    [void](& $cases[$name] $w $ws)
    $n++; $path = Join-Path $Books "freeze$n.xlsx"
    $wb.SaveAs($path, 51)
    $row = [pscustomobject]@{ name = $name; inView = $before; frozen = [bool]$w.FreezePanes; splitRow = $w.SplitRow; splitColumn = $w.SplitColumn; scrollRow = $w.ScrollRow; scrollColumn = $w.ScrollColumn; active = $x.ActiveCell.Address($false, $false); sheetViews = '' }
    $wb.Close($false)
    $row.sheetViews = ViewXml $path
    $results += $row
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x); [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
$results | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 $Out
"$($results.Count) cases"
