# Where a table, a chart, a shape and notes go when Excel inserts or deletes rows, columns or cells (anchor-cases.json):
# for each case a new workbook in an invisible Excel, the sheet built, the step done, and read back: the table's range
# (or gone), the chart's and the shape's top-left and bottom-right cells and their place and size in points, each
# note's cell, or Excel's refusal. Nothing is saved (a note carries Excel's user name). No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-anchors.ps1 <anchor-cases.json> <out.json>
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$spec = Get-Content -Raw $Cases | ConvertFrom-Json
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
function Box($o) { return @{ from = $o.TopLeftCell.Address($false, $false); to = $o.BottomRightCell.Address($false, $false); left = $o.Left; top = $o.Top; width = $o.Width; height = $o.Height } }
$results = @()
try {
  foreach ($case in $spec.cases) {
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $d = $spec.sheet.data
    $grid = New-Object 'object[,]' $d.Count, 3
    for ($i = 0; $i -lt $d.Count; $i++) { $row = @($d[$i]); for ($j = 0; $j -lt 3; $j++) { $grid[$i, $j] = $row[$j] } }
    $top = $ws.Range($spec.sheet.at)
    $ws.Range($top, $top.Offset($d.Count - 1, 2)).Value2 = $grid
    if ($spec.sheet.table) { $lo = $ws.ListObjects.Add(1, $ws.Range($top, $top.Offset($d.Count - 1, 2)), [Type]::Missing, 1); $lo.Name = 'Stock' }
    $cr = $ws.Range($spec.sheet.chart.range)
    $co = $ws.ChartObjects().Add($cr.Left, $cr.Top, $cr.Width, $cr.Height)
    $co.Chart.ChartType = 51
    [void]$co.Chart.SetSourceData($ws.Range($spec.sheet.chart.series))
    $sr = $ws.Range($spec.sheet.shape)
    $sh = $ws.Shapes.AddShape(1, $sr.Left, $sr.Top, $sr.Width, $sr.Height)
    foreach ($n in $spec.sheet.notes) { [void]$ws.Range($n).AddComment("note $n") }
    if ($spec.sheet.formula) { $ws.Range($spec.sheet.formula.at).Formula = $spec.sheet.formula.text }
    $before = @{ chart = (Box $co); shape = (Box $sh) }
    $refused = $null
    $s = $case.step
    try {
      $r = $ws.Range($s.range)
      $shift = if ($s.shift -eq 'down') { -4121 } elseif ($s.shift -eq 'right') { -4161 } elseif ($s.shift -eq 'up') { -4162 } elseif ($s.shift -eq 'left') { -4159 } else { $null }
      # (A whole-line range inserts whole lines by itself; cells shift as told.)
      if ($s.do -eq 'insert') { if ($null -ne $shift) { [void]$r.Insert($shift) } else { [void]$r.Insert() } }
      if ($s.do -eq 'delete') { if ($null -ne $shift) { [void]$r.Delete($shift) } else { [void]$r.Delete() } }
    } catch { $refused = $_.Exception.Message }
    $table = if ($ws.ListObjects.Count) { $ws.ListObjects.Item(1).Range.Address($false, $false) } else { $null }
    $columns = if ($ws.ListObjects.Count) { @($ws.ListObjects.Item(1).ListColumns | ForEach-Object { $_.Name }) } else { @() }
    $header = if ($ws.ListObjects.Count -and $ws.ListObjects.Item(1).ShowHeaders) { @($ws.ListObjects.Item(1).HeaderRowRange.Cells | ForEach-Object { "$($_.Address($false, $false))=$($_.Value2)" }) } else { @() }
    $formula = $null
    if ($spec.sheet.formula) { $hit = $ws.UsedRange.Find('SUM(', [Type]::Missing, -4123, 2); if ($hit) { $formula = "$($hit.Address($false, $false)) $($hit.Formula) = $($hit.Text)" } }
    $charts = @(); foreach ($c in $ws.ChartObjects()) { $charts += , (Box $c) }
    $shapes = @(); foreach ($p in $ws.Shapes) { if ($p.Type -eq 1) { $shapes += , (Box $p) } }
    $notes = @(); foreach ($c in $ws.Comments) { $notes += $c.Parent.Address($false, $false) }
    $results += [pscustomobject]@{ name = $case.name; refused = $refused; table = $table; columns = $columns; header = $header; formula = $formula; charts = $charts; shapes = $shapes; notes = $notes; before = $before }
    $wb.Close($false)
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$results | ConvertTo-Json -Depth 8 | Set-Content -Encoding utf8 $Out
"$($results.Count) cases"
