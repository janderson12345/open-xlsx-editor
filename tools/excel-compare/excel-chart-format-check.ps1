# Excel's reading of formatted charts: a workbook opened read only in an invisible Excel, and each sheet's first chart
# read through the object model: its axes' titles (Axis.HasTitle, AxisTitle.Text), each series' chart type and axis
# group, fill and line colour (Format.Fill / Format.Line: ForeColor.RGB as RRGGBB, Visible), its second point's fill,
# the chart and plot areas' fill and line, and the title's, horizontal axis's and legend's fonts. Run on Excel's own
# workbook of excel-chart-format.ps1 and on the one chartformat.test.js writes (test/out/charts/chartformat-all.xlsx),
# so the two readings can be compared. Nothing saved, no keys sent; only the Excel this script starts is used and quit.
# usage: pwsh tools/excel-compare/excel-chart-format-check.ps1 <workbook> <out.json>
param([Parameter(Mandatory)][string]$Book, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
function hex($v) { $r = $v -band 255; $g = ($v -shr 8) -band 255; $b = ($v -shr 16) -band 255; '{0:X2}{1:X2}{2:X2}' -f $r, $g, $b }
function look($f) { [pscustomobject]@{ fill = $(if ($f.Fill.Visible) { hex $f.Fill.ForeColor.RGB } else { 'none' }); line = $(if ($f.Line.Visible) { hex $f.Line.ForeColor.RGB } else { 'none' }) } }
function font($f) { [pscustomobject]@{ name = $f.Name; size = $f.Size; bold = [bool]$f.Bold; color = (hex $f.Color) } }
function title($ch, $type, $group) { if ($ch.HasAxis($type, $group) -and $ch.Axes($type, $group).HasTitle) { $ch.Axes($type, $group).AxisTitle.Text } else { $null } }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$found = [ordered]@{}
try {
  $wb = $x.Workbooks.Open((Resolve-Path $Book).Path, 0, $true)
  foreach ($ws in $wb.Worksheets) {
    if ($ws.ChartObjects().Count -lt 1) { continue }
    $ch = $ws.ChartObjects(1).Chart
    $series = @(foreach ($s in $ch.SeriesCollection()) {
      $p = $null
      if ($s.Points().Count -ge 2) { $p = (look $s.Points(2).Format).fill }
      [pscustomobject]@{ type = $s.ChartType; group = $s.AxisGroup; look = (look $s.Format); point2 = $p }
    })
    $found[$ws.Name] = [pscustomobject]@{
      titles = @((title $ch 1 1), (title $ch 2 1), (title $ch 2 2))
      series = $series
      area = (look $ch.ChartArea.Format); plot = (look $ch.PlotArea.Format)
      title = $(if ($ch.HasTitle) { font $ch.ChartTitle.Font } else { $null })
      cat = $(if ($ch.HasAxis(1, 1)) { font $ch.Axes(1).TickLabels.Font } else { $null })
      legend = $(if ($ch.HasLegend) { font $ch.Legend.Font } else { $null })
    }
  }
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
[pscustomobject]@{ source = "tools/excel-compare/excel-chart-format-check.ps1 on $(Split-Path -Leaf $Book)"; charts = $found } | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 $Out
'done'
