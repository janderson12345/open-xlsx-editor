# Excel's reading of charts written here (chartnew.test.js): test/out/charts/chartnew-all.xlsx opened in an invisible
# Excel, read only; each sheet's first chart's type, series count and first series' formula. Nothing saved, no keys sent.
# usage: pwsh tools/excel-compare/excel-chart-readback.ps1 <workbook> <out.json>
param([Parameter(Mandatory)][string]$Book, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$found = @()
try {
  $wb = $x.Workbooks.Open((Resolve-Path $Book).Path, 0, $true)
  foreach ($ws in $wb.Worksheets) {
    if ($ws.ChartObjects().Count -lt 1) { continue }
    $ch = $ws.ChartObjects(1).Chart
    $found += [pscustomobject]@{ sheet = $ws.Name; type = $ch.ChartType; series = $ch.SeriesCollection().Count; formula1 = $ch.SeriesCollection(1).Formula; title = $(if ($ch.HasTitle) { $ch.ChartTitle.Text } else { $null }); legend = $ch.HasLegend }
  }
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
[pscustomobject]@{ source = 'tools/excel-compare/excel-chart-readback.ps1 on test/out/charts/chartnew-all.xlsx (chartnew.test.js)'; charts = $found } | ConvertTo-Json -Depth 3 | Set-Content -Encoding utf8 $Out
'done'
