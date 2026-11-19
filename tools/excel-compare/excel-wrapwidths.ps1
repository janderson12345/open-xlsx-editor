# Where Excel breaks wrapped text when it works out a row's height (src/editor/rowheights.js excelLineCount): each
# case of wrapwidths-cases.mjs ({ text, font, size, bold, italic, halign, indent }) alone in cell (i, i) of a new, never
# saved workbook whose Normal style is Calibri 11 (so a column of n px is (n - 5) / 7 characters), its rows AutoFit.
#   -Mode scan (default): all columns at each width from -From to -To px, every row's height read: written as
#     { excel, from, to, cases: [{ ..., heights: [[px, pt], ...] }] }, a pair where the height changes.
#   -Mode narrowest: for each case the narrowest column that keeps it on one line, by halving: written as
#     { excel, cases: [{ ..., px }] }. Of n of one character that is n x its width + Excel's margin, which gives the
#     character's width and the margin both (docs/LAYOUT.md).
# A workbook holds some 500 fonts at most, so the cases go 300 to a workbook.
# An invisible instance, macros disabled for it only, nothing opened or saved, no keys sent.
# usage: node tools/excel-compare/wrapwidths-cases.mjs <cases.json>
#        pwsh tools/excel-compare/excel-wrapwidths.ps1 <cases.json> <out.json> [-Mode narrowest] [-From 15] [-To 400]
param([Parameter(Mandatory)][string]$In, [Parameter(Mandatory)][string]$Out, [ValidateSet('scan', 'narrowest')][string]$Mode = 'scan', [int]$From = 15, [int]$To = 400, [int]$Max = 1790)
$ErrorActionPreference = 'Stop'
$all = @([IO.File]::ReadAllText((Resolve-Path $In).Path, [Text.Encoding]::UTF8) | ConvertFrom-Json)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
try {
  $res = [ordered]@{ excel = $x.Version }
  $list = @()
  # (A workbook holds some 500 fonts at most, so the cases go 300 to a workbook.)
  for ($at = 0; $at -lt $all.Count; $at += 300) {
  $cases = @($all[$at..([math]::Min($all.Count, $at + 300) - 1)])
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $normal = $wb.Styles.Item('Normal').Font
  $normal.Name = 'Calibri'; $normal.Size = 11
  $count = $cases.Count
  $cols = @($null); $rows = @($null)
  for ($i = 1; $i -le $count; $i++) {
    $c = $cases[$i - 1]
    $cell = $ws.Cells.Item($i, $i)
    $cell.NumberFormat = '@'
    $cell.Font.Name = $c.font; $cell.Font.Size = [double]$c.size
    if ($c.bold) { $cell.Font.Bold = $true }
    if ($c.italic) { $cell.Font.Italic = $true }
    $cell.WrapText = $true
    # (xlCenter, xlRight, xlLeft.)
    if ($c.halign -eq 'center') { $cell.HorizontalAlignment = -4108 }
    if ($c.halign -eq 'right') { $cell.HorizontalAlignment = -4152 }
    if ($c.halign -eq 'left') { $cell.HorizontalAlignment = -4131 }
    if ($c.indent) { $cell.IndentLevel = [int]$c.indent }
    $cell.Value2 = [string]$c.text
    $cols += $ws.Columns.Item($i); $rows += $ws.Rows.Item($i)
  }
  $allCols = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $count)).EntireColumn
  $allRows = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($count, 1)).EntireRow
  # (A column's width in characters for a width in pixels; a double, or COM takes a whole number as an Int32.)
  function Chars($px) { return [double](($px - 5) / 7) }
  if ($Mode -eq 'scan') {
    $res.from = $From; $res.to = $To
    $heights = @(); foreach ($c in $cases) { $heights += , (New-Object System.Collections.ArrayList) }
    $last = @($null) * ($count + 1)
    for ($px = $From; $px -le $To; $px++) {
      $allCols.ColumnWidth = Chars $px
      $got = [math]::Round([double]$cols[1].Width * 4 / 3)
      if ($got -ne $px) { throw "a column of $px px came out $got" }
      [void]$allRows.AutoFit()
      for ($i = 1; $i -le $count; $i++) {
        $h = [double]$rows[$i].RowHeight
        if ($h -ne $last[$i]) { [void]$heights[$i - 1].Add(@($px, $h)); $last[$i] = $h }
      }
    }
    for ($i = 1; $i -le $count; $i++) { $c = $cases[$i - 1]; $list += [ordered]@{ text = $c.text; font = $c.font; size = $c.size; bold = [bool]$c.bold; italic = [bool]$c.italic; halign = $c.halign; indent = $c.indent; heights = $heights[$i - 1] } }
  } else {
    $lo = @(0) * ($count + 1); $hi = @(0) * ($count + 1); $one = @(0.0) * ($count + 1)
    for ($i = 1; $i -le $count; $i++) { $cols[$i].ColumnWidth = Chars $Max; $lo[$i] = 12; $hi[$i] = $Max }
    [void]$allRows.AutoFit()
    for ($i = 1; $i -le $count; $i++) { $one[$i] = [double]$rows[$i].RowHeight }
    # (lo: too narrow, from 12px, the narrowest column these widths in characters come out exact at; hi: fits; until they
    # meet. A case that fits in 12px reads 13.)
    while ($true) {
      $open = 0
      for ($i = 1; $i -le $count; $i++) { if ($hi[$i] - $lo[$i] -gt 1) { $open++; $cols[$i].ColumnWidth = Chars ([math]::Floor(($lo[$i] + $hi[$i]) / 2)) } }
      if (-not $open) { break }
      [void]$allRows.AutoFit()
      for ($i = 1; $i -le $count; $i++) {
        if ($hi[$i] - $lo[$i] -le 1) { continue }
        $mid = [math]::Floor(($lo[$i] + $hi[$i]) / 2)
        $got = [math]::Round([double]$cols[$i].Width * 4 / 3)
        if ($got -ne $mid) { throw "case $i`: a column of $mid px came out $got" }
        if ([double]$rows[$i].RowHeight -le $one[$i] + 0.01) { $hi[$i] = $mid } else { $lo[$i] = $mid }
      }
    }
    for ($i = 1; $i -le $count; $i++) { $c = $cases[$i - 1]; $list += [ordered]@{ text = $c.text; font = $c.font; size = $c.size; bold = [bool]$c.bold; italic = [bool]$c.italic; px = $hi[$i] } }
  }
  $wb.Close($false)
  }
  $res.cases = $list
  # (UTF-8 without a byte order mark, which Windows PowerShell's Set-Content would add.)
  [IO.File]::WriteAllText($Out, ($res | ConvertTo-Json -Compress -Depth 6) + "`n", (New-Object Text.UTF8Encoding $false))
  Write-Output "$($all.Count) cases, $Mode, Excel $($x.Version)"
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
