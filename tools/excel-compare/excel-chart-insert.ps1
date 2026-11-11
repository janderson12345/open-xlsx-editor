# What Excel writes when it inserts a chart, for chartnew.test.js: in an invisible Excel, a new workbook with a small
# made-up table (Region, Sales, Cost) on Sheet1; one chart of each kind added with Shapes.AddChart2 (Excel's default
# style, as Insert > Charts gives it) on its own sheet, the workbook saved outside the repository and every part of it
# under xl/charts, xl/drawings and the sheets' relationships read back. No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-chart-insert.ps1 <out.json> <folder for the saved workbook>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
# Each kind: Excel's XlChartType.
$KINDS = [ordered]@{ columnClustered = 51; columnStacked = 52; columnStacked100 = 53; barClustered = 57; barStacked = 58; line = 4; lineMarkers = 65;
  pie = 5; doughnut = -4120; area = 1; areaStacked = 76; scatter = -4169; scatterLines = 74; scatterSmooth = 72 }
$DATA = @(@('Region', 'Sales', 'Cost'), @('North', 120, 80), @('South', 95, 70), @('East', 140, 90), @('West', 60, 45))
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$read = @{}
$order = @()
try {
  $wb = $x.Workbooks.Add()
  $first = $wb.Worksheets.Item(1)
  $k = 0
  foreach ($name in $KINDS.Keys) {
    $ws = if ($k -eq 0) { $first } else { $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $ws.Name = $name.Substring(0, [Math]::Min(31, $name.Length))
    for ($r = 0; $r -lt $DATA.Count; $r++) { for ($c = 0; $c -lt 3; $c++) { $ws.Range("$([char](65 + $c))$($r + 1)").Formula = [string]$DATA[$r][$c] } }
    $ws.Range('A1:C5').Select()
    $shape = $ws.Shapes.AddChart2(-1, $KINDS[$name], 240, 15, 360, 216)
    $order += [pscustomobject]@{ kind = $name; sheet = $ws.Name; type = $shape.Chart.ChartType; series = $shape.Chart.SeriesCollection().Count; formula1 = $shape.Chart.SeriesCollection(1).Formula }
    $k++
  }
  $first.Activate()
  $file = Join-Path $Books 'chart-insert.xlsx'
  if (Test-Path $file) { Remove-Item $file }
  $wb.SaveAs($file, 51)
  $wb.Close($false)
  $zip = [IO.Compression.ZipFile]::OpenRead($file)
  foreach ($e in $zip.Entries) {
    if ($e.FullName -match '^(xl/charts/|xl/drawings/|xl/worksheets/_rels/|\[Content_Types\]\.xml|xl/worksheets/sheet\d+\.xml)') {
      $s = New-Object IO.StreamReader($e.Open()); $read[$e.FullName] = $s.ReadToEnd(); $s.Close()
    }
  }
  $zip.Dispose()
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
[pscustomobject]@{ charts = $order; parts = $read } | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 $Out
'done'
