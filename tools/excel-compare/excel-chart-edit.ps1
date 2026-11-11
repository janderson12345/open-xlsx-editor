# What Excel writes when a chart is edited, for chartedit.test.js: in an invisible Excel, a new workbook with the small
# table of excel-chart-insert.ps1 on each sheet and a clustered column chart of it (Shapes.AddChart2); each case one
# edit through the chart's object model (its title's text, no title, the legend moved or gone, data labels on, the
# chart type changed); the workbook saved outside the repository and each sheet's chart part read back.
# No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-chart-edit.ps1 <out.json> <folder for the saved workbook>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$CASES = [ordered]@{
  'none' = { param($ch) }
  'titleText' = { param($ch) $ch.HasTitle = $true; $ch.ChartTitle.Text = 'Sales & cost' }
  'noTitle' = { param($ch) $ch.HasTitle = $false }
  'noLegend' = { param($ch) $ch.HasLegend = $false }
  'legendRight' = { param($ch) $ch.Legend.Position = -4152 }
  'legendTop' = { param($ch) $ch.Legend.Position = -4160 }
  'legendLeft' = { param($ch) $ch.Legend.Position = -4131 }
  'dataLabels' = { param($ch) $ch.SetElement(201) }
  'toLine' = { param($ch) $ch.ChartType = 4 }
  'toBar' = { param($ch) $ch.ChartType = 57 }
  'toPie' = { param($ch) $ch.ChartType = 5 }
  'toArea' = { param($ch) $ch.ChartType = 1 }
  'toColumnStacked' = { param($ch) $ch.ChartType = 52 }
  'toScatter' = { param($ch) $ch.ChartType = -4169 }
  'toDoughnut' = { param($ch) $ch.ChartType = -4120 }
  'toLineMarkers' = { param($ch) $ch.ChartType = 65 }
}
$DATA = @(@('Region', 'Sales', 'Cost'), @('North', '120', '80'), @('South', '95', '70'), @('East', '140', '90'), @('West', '60', '45'))
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$read = @{}
try {
  $wb = $x.Workbooks.Add()
  $k = 0
  foreach ($name in $CASES.Keys) {
    $ws = if ($k -eq 0) { $wb.Worksheets.Item(1) } else { $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $ws.Name = $name
    for ($r = 0; $r -lt $DATA.Count; $r++) { for ($c = 0; $c -lt 3; $c++) { $ws.Range("$([char](65 + $c))$($r + 1)").Formula = $DATA[$r][$c] } }
    $ws.Range('A1:C5').Select()
    $ch = $ws.Shapes.AddChart2(-1, 51, 240, 15, 360, 216).Chart
    & $CASES[$name] $ch
    $k++
  }
  $wb.Worksheets.Item(1).Activate()
  $file = Join-Path $Books 'chart-edit.xlsx'
  if (Test-Path $file) { Remove-Item $file }
  $wb.SaveAs($file, 51)
  $wb.Close($false)
  $zip = [IO.Compression.ZipFile]::OpenRead($file)
  $i = 1
  foreach ($name in $CASES.Keys) {
    $s = New-Object IO.StreamReader($zip.GetEntry("xl/charts/chart$i.xml").Open()); $read[$name] = $s.ReadToEnd(); $s.Close()
    $i++
  }
  $zip.Dispose()
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
[pscustomobject]@{ charts = $read } | ConvertTo-Json -Depth 3 | Set-Content -Encoding utf8 $Out
'done'
