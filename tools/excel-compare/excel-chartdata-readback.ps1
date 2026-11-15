# Excel's reading of the charts whose data was changed here (chartdata.test.js): test/out/charts/chartdata-all.xlsx
# opened in an invisible Excel, read only; each sheet's first chart's series formulas, in their order. Nothing saved,
# no keys sent.
# usage: pwsh tools/excel-compare/excel-chartdata-readback.ps1 <workbook> <out.json>
param([Parameter(Mandatory)][string]$Book, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$found = [ordered]@{}
try {
  $wb = $x.Workbooks.Open((Resolve-Path $Book).Path, 0, $true)
  foreach ($ws in $wb.Worksheets) {
    if ($ws.ChartObjects().Count -lt 1) { continue }
    $ch = $ws.ChartObjects(1).Chart
    $f = @()
    for ($s = 1; $s -le $ch.SeriesCollection().Count; $s++) { $f += $ch.SeriesCollection($s).Formula }
    $found[$ws.Name] = $f
  }
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$json = [pscustomobject]@{ source = 'tools/excel-compare/excel-chartdata-readback.ps1 on test/out/charts/chartdata-all.xlsx (chartdata.test.js)'; charts = $found } | ConvertTo-Json -Depth 4
[IO.File]::WriteAllText($Out, $json, (New-Object Text.UTF8Encoding($false)))
'done'
