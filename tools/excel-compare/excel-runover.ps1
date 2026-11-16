# Which gridlines Excel leaves out under text that runs on into empty cells: a new workbook (Normal style as Excel
# makes it) with Page Layout > Gridlines: Print on, text left, right and centred running on over empty neighbours,
# text stopped by a neighbour, and runs of zeros 6 to 12 long (as wide as that many digits, against an 8.43
# column), exported to PDF; shapes.cjs then gives every gridline segment. Saved as .xlsx too, for print-ours.
# An invisible instance, macros disabled for it only, no keys sent.
# usage: pwsh tools/excel-compare/excel-runover.ps1 <out folder>
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $Out | Out-Null
$Out = (Resolve-Path $Out).Path
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $set = { param($r, $c, $v, $h) $cell = $ws.Cells.Item($r, $c); $cell.NumberFormat = '@'; $cell.Value2 = $v; if ($h) { $cell.HorizontalAlignment = [int]$h } }
  # xlLeft -4131, xlRight -4152, xlCenter -4108
  & $set 2 2 'Left text running on over three columns' $null
  & $set 4 5 'Right text running on to the left' -4152
  & $set 6 4 'Centred text running both ways here' -4108
  & $set 8 2 'Stopped by a neighbour, not further' $null
  & $set 8 4 'x' $null
  & $set 10 2 'Into a filled cell and past it too' $null
  $ws.Cells.Item(10, 3).Interior.Color = 0xCCFFCC
  & $set 12 3 'Short' $null
  for ($n = 6; $n -le 12; $n++) { & $set (12 + 2 * ($n - 5)) 2 ('0' * $n) $null }
  & $set 28 4 ('0' * 10) -4152
  & $set 30 4 ('0' * 10) -4108
  & $set 32 4 ('0' * 11) -4108
  $ws.PageSetup.PrintGridlines = $true
  $ws.PageSetup.PrintArea = 'A1:H33'
  $wb.SaveAs((Join-Path $Out 'runover.xlsx'), 51)
  $ws.ExportAsFixedFormat(0, (Join-Path $Out 'runover.pdf'))
  $wb.Close($false)
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
'runover.xlsx, runover.pdf'
