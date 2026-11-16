# The maximum digit width Excel lays columns out with, by the Normal style's font: in a new workbook the Normal style
# is given each font and size, then columns are set to 1, 2, 5, 10 and 20 characters and their Range.Width (points)
# read back, with the sheet's StandardWidth and the default column's width. Pixels are points x 96 / 72 at 100% zoom
# (Excel's Width is in points of the screen's pixels, so the screen's scale is read from Calibri 11, 64px for 8.43).
# An invisible instance, macros disabled for it only, the workbook closed without saving, no keys sent.
# With a folder as second argument each workbook is also saved there (to read the <col> widths Excel writes).
# usage: pwsh tools/excel-compare/excel-digitwidth.ps1 <out.json> [books folder]
param([Parameter(Mandatory)][string]$Out, [string]$Books)
$ErrorActionPreference = 'Stop'
$fonts = @(
  @('Calibri', 8), @('Calibri', 9), @('Calibri', 10), @('Calibri', 11), @('Calibri', 12), @('Calibri', 14), @('Calibri', 16), @('Calibri', 18), @('Calibri', 20), @('Calibri', 24),
  @('Aptos Narrow', 11), @('Aptos', 11), @('Aptos', 10), @('Aptos', 12),
  @('Arial', 8), @('Arial', 9), @('Arial', 10), @('Arial', 11), @('Arial', 12), @('Arial', 14), @('Arial', 16),
  @('Times New Roman', 10), @('Times New Roman', 11), @('Times New Roman', 12), @('Times New Roman', 14),
  @('Cambria', 10), @('Cambria', 11), @('Cambria', 12), @('Segoe UI', 9), @('Segoe UI', 10), @('Segoe UI', 11),
  @('Verdana', 8), @('Verdana', 10), @('Verdana', 12), @('Courier New', 10), @('Courier New', 11), @('Courier New', 12),
  @('Tahoma', 10), @('Georgia', 11), @('Consolas', 11), @('MS Sans Serif', 10)
)
$chars = @(1, 2, 5, 10, 20)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$rows = New-Object System.Collections.Generic.List[object]
try {
  foreach ($f in $fonts) {
    $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
    $st = $wb.Styles.Item('Normal')
    $st.Font.Name = $f[0]; $st.Font.Size = $f[1]
    $widths = @()
    for ($k = 0; $k -lt $chars.Count; $k++) {
      $col = $ws.Columns.Item($k + 2)
      $col.ColumnWidth = $chars[$k]
      $widths += [double]$col.Width
    }
    # AutoFit of ten zeros as text (as wide as ten digits): its padding over 10 x mdw.
    $ws.Cells.Item(1, 10).Value2 = "'0000000000"
    [void]$ws.Columns.Item(10).AutoFit()
    $rows.Add([ordered]@{ font = $f[0]; size = $f[1]; got = [string]$st.Font.Name; standardWidth = [double]$ws.StandardWidth; defaultPt = [double]$ws.Columns.Item(1).Width; chars = $chars; pt = $widths; fitPt = [double]$ws.Columns.Item(10).Width })
    if ($Books) { $wb.SaveAs((Join-Path $Books ("{0}-{1}.xlsx" -f ($f[0] -replace ' ', ''), $f[1])), 51) }
    $wb.Close($false)
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
$json = [ordered]@{ excel = 'Excel.Application, Normal style font set, Range.Width in points'; fonts = $rows } | ConvertTo-Json -Depth 4
[IO.File]::WriteAllText($Out, $json, (New-Object Text.UTF8Encoding $false))
"$($rows.Count) fonts"
