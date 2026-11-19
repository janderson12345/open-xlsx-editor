# The height Excel gives a row of 1, 2, 3... lines of wrapped text, font by font (src/editor/rowheights.js): each case
# ({ font, size, bold, italic, lines }) is "Ag" on that many lines (line breaks in the cell, wrap on) alone in column A
# of a new, never saved workbook whose Normal style is -NormalFont at -NormalSize, its row AutoFit. Written as
# { excel, normal, standard, cases: [{ ..., pt }] }, `standard` the sheet's standard row height.
# A workbook holds some 500 fonts at most, so the cases go 300 to a workbook.
# An invisible instance, macros disabled for it only, nothing opened or saved, no keys sent.
# usage: pwsh tools/excel-compare/excel-wraprows.ps1 <cases.json> <out.json> [-NormalFont Calibri] [-NormalSize 11]
param([Parameter(Mandatory)][string]$In, [Parameter(Mandatory)][string]$Out, [string]$NormalFont = 'Calibri', [double]$NormalSize = 11)
$ErrorActionPreference = 'Stop'
$cases = @([IO.File]::ReadAllText((Resolve-Path $In).Path, [Text.Encoding]::UTF8) | ConvertFrom-Json)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$x.AutomationSecurity = 3
try {
  $list = New-Object System.Collections.ArrayList
  $standard = $null
  for ($from = 0; $from -lt $cases.Count; $from += 300) {
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $normal = $wb.Styles.Item('Normal').Font
    $normal.Name = $NormalFont; $normal.Size = $NormalSize
    $ws.Columns.Item(1).ColumnWidth = [double]30
    $count = [math]::Min(300, $cases.Count - $from)
    for ($i = 1; $i -le $count; $i++) {
      $c = $cases[$from + $i - 1]
      $cell = $ws.Cells.Item($i, 1)
      $cell.NumberFormat = '@'
      try { $cell.Font.Name = $c.font; $cell.Font.Size = [double]$c.size } catch { throw "case $($from + $i) ($($c.font) $($c.size)): $($_.Exception.Message)" }
      if ($c.bold) { $cell.Font.Bold = $true }
      if ($c.italic) { $cell.Font.Italic = $true }
      $cell.WrapText = $true
      $cell.Value2 = (@('Ag') * [int]$c.lines) -join "`n"
    }
    [void]$ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($count, 1)).EntireRow.AutoFit()
    for ($i = 1; $i -le $count; $i++) { $c = $cases[$from + $i - 1]; [void]$list.Add([ordered]@{ font = $c.font; size = $c.size; bold = [bool]$c.bold; italic = [bool]$c.italic; lines = $c.lines; pt = [double]$ws.Rows.Item($i).RowHeight }) }
    $standard = [double]$ws.StandardHeight
    $wb.Close($false)
  }
  $res = [ordered]@{ excel = $x.Version; normal = "$NormalFont $NormalSize"; standard = $standard; cases = $list }
  # (UTF-8 without a byte order mark, which Windows PowerShell's Set-Content would add.)
  [IO.File]::WriteAllText($Out, ($res | ConvertTo-Json -Compress -Depth 5) + "`n", (New-Object Text.UTF8Encoding $false))
  Write-Output "$($cases.Count) cases, Normal $NormalFont $NormalSize, standard height $standard"
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
