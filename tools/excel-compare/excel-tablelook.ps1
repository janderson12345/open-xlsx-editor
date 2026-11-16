# Has Excel say how it shows a table's cells in a table style: for each case, a new workbook with the same made-up
# 6-row, 5-column list (a header row, five data rows, a totals row once ShowTotals is on), made a table
# (ListObjects.Add), given a style and the Table Style Options of the case, and then each cell's Range.DisplayFormat
# read (what Excel shows: the table style under the cell's own format): fill colour and pattern, font colour, bold and
# italic, and the four edges' line style, weight and colour. Colours as RRGGBB. Cells one row and column outside the
# table are read too. Cases: [{ name, style, header, totals, rowStripes, colStripes, first, last, own: [{ cell, fill,
# font, bold, border }] (direct formats put on before the table is made) }]; with --all-styles, the case is run once for
# each built-in style. An invisible instance, macros disabled for it only; no setting changed, no keys sent, nothing
# saved.
# usage: pwsh tools/excel-compare/excel-tablelook.ps1 <cases.json> <out.json> [-AllStyles]
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out, [switch]$AllStyles)
$ErrorActionPreference = 'Stop'
$list = @(Get-Content $Cases -Raw | ConvertFrom-Json)
if ($AllStyles) {
  $names = @()
  foreach ($n in 1..21) { $names += "TableStyleLight$n" }
  foreach ($n in 1..28) { $names += "TableStyleMedium$n" }
  foreach ($n in 1..11) { $names += "TableStyleDark$n" }
  $base = $list[0]
  $list = foreach ($n in $names) { $c = $base.PSObject.Copy(); $c.name = $n; $c.style = $n; $c }
}
$hex = { param($v) if ($null -eq $v -or $v -is [DBNull]) { $null } else { $i = [int64]$v; '{0:X2}{1:X2}{2:X2}' -f ($i -band 255), (($i -shr 8) -band 255), (($i -shr 16) -band 255) } }
$rows = @(@('Region', 'Item', 'Qty', 'Price', 'Note'), @('North', 'kiwi', 30, 1.5, 'a'), @('South', 'apple', 10, 0.5, 'b'),
  @('East', 'pear', 20, 2, 'c'), @('West', 'date', 50, 3, 'd'), @('North', 'fig', 40, 2.5, 'e'))
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$answers = New-Object System.Collections.Generic.List[object]
try {
  foreach ($case in $list) {
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    for ($i = 0; $i -lt $rows.Count; $i++) { for ($j = 0; $j -lt 5; $j++) { $ws.Cells.Item($i + 2, $j + 2).Formula = [string]$rows[$i][$j] } }
    foreach ($o in @($case.own)) {
      if ($null -eq $o) { continue }
      $rg = $ws.Range($o.cell)
      if ($o.fill) { $rg.Interior.Color = [int]$o.fill }
      if ($null -ne $o.font) { $rg.Font.Color = [int]$o.font }
      if ($null -ne $o.bold) { $rg.Font.Bold = [bool]$o.bold }
      if ($o.border) { $rg.Borders.Item(9).LineStyle = 1; $rg.Borders.Item(9).Weight = -4138 }
    }
    $lo = $ws.ListObjects.Add(1, $ws.Range('B2:F7'), [Type]::Missing, 1)
    $lo.TableStyle = $case.style
    if ($null -ne $case.totals) { $lo.ShowTotals = [bool]$case.totals }
    if ($null -ne $case.header) { $lo.ShowHeaders = [bool]$case.header }
    if ($null -ne $case.rowStripes) { $lo.ShowTableStyleRowStripes = [bool]$case.rowStripes }
    if ($null -ne $case.colStripes) { $lo.ShowTableStyleColumnStripes = [bool]$case.colStripes }
    if ($null -ne $case.first) { $lo.ShowTableStyleFirstColumn = [bool]$case.first }
    if ($null -ne $case.last) { $lo.ShowTableStyleLastColumn = [bool]$case.last }
    $a = [ordered]@{ name = $case.name; style = $case.style; ref = $lo.Range.Address($false, $false); cells = [ordered]@{} }
    $area = $lo.Range
    $r0 = $area.Row - 1; $c0 = $area.Column - 1; $r1 = $area.Row + $area.Rows.Count; $c1 = $area.Column + $area.Columns.Count
    for ($r = $r0; $r -le $r1; $r++) {
      for ($c = $c0; $c -le $c1; $c++) {
        $d = $ws.Cells.Item($r, $c).DisplayFormat
        # (One string per cell: "fill:RRGGBB font:RRGGBB b i L:style/weight/RRGGBB T:... B:... R:...", COM's
        # XlLineStyle and XlBorderWeight numbers; no fill: no pattern.)
        $cell = @()
        if ($d.Interior.Pattern -ne -4142) { $cell += "fill:$(& $hex $d.Interior.Color)" }
        $cell += "font:$(& $hex $d.Font.Color)"
        if ($d.Font.Bold) { $cell += 'b' }
        if ($d.Font.Italic) { $cell += 'i' }
        foreach ($k in @(@('L', 7), @('T', 8), @('B', 9), @('R', 10))) {
          $e = $d.Borders.Item($k[1])
          if ($e.LineStyle -ne -4142) { $cell += "$($k[0]):$($e.LineStyle)/$($e.Weight)/$(& $hex $e.Color)" }
        }
        $a.cells[$ws.Cells.Item($r, $c).Address($false, $false)] = $cell -join ' '
      }
    }
    $answers.Add($a)
    $wb.Close($false)
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
$json = ConvertTo-Json -InputObject $answers.ToArray() -Depth 6 -Compress
[IO.File]::WriteAllText(([IO.Path]::GetFullPath($Out, (Get-Location).Path)), $json + "`n", (New-Object Text.UTF8Encoding($false)))
