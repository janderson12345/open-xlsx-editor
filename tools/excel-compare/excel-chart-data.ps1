# What Excel writes when a chart's data is changed (Select Data), for chartdata.test.js: in an invisible Excel, a new
# workbook with a small made-up table (Region, Sales, Cost, Profit over four regions) on each sheet and a clustered
# column chart of A1:C5 (Shapes.AddChart2); each case one change through the chart's object model (SetSourceData,
# PlotBy, SeriesCollection.NewSeries, a series' Name, Values, XValues, Formula, PlotOrder, Delete, DisplayBlanksAs,
# PlotVisibleOnly); the workbook saved outside the repository and each sheet's chart part read back, with each
# series' formula as Excel gives it.
# No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-chart-data.ps1 <out.json> <folder for the saved workbook>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$CASES = [ordered]@{
  'none' = { param($ch, $ws) }
  'sourceWider' = { param($ch, $ws) $ch.SetSourceData($ws.Range('A1:D5')) }
  'sourceSmaller' = { param($ch, $ws) $ch.SetSourceData($ws.Range('A1:B4')) }
  'sourceRows' = { param($ch, $ws) $ch.SetSourceData($ws.Range('A1:C5'), 1) }
  'sourceCols' = { param($ch, $ws) $ch.SetSourceData($ws.Range('A1:D3'), 2) }
  'plotByRows' = { param($ch, $ws) $ch.PlotBy = 1 }
  'newSeries' = { param($ch, $ws) $s = $ch.SeriesCollection().NewSeries(); $s.Name = "='newSeries'!`$D`$1"; $s.Values = "='newSeries'!`$D`$2:`$D`$5"; $s.XValues = "='newSeries'!`$A`$2:`$A`$5" }
  'newSeriesBare' = { param($ch, $ws) $s = $ch.SeriesCollection().NewSeries(); $s.Values = "='newSeriesBare'!`$D`$2:`$D`$5" }
  'nameText' = { param($ch, $ws) $ch.SeriesCollection(1).Name = 'Takings' }
  'nameRef' = { param($ch, $ws) $ch.SeriesCollection(1).Name = "='nameRef'!`$D`$1" }
  'valuesRef' = { param($ch, $ws) $ch.SeriesCollection(2).Values = "='valuesRef'!`$D`$2:`$D`$5" }
  'xvaluesRef' = { param($ch, $ws) $ch.SeriesCollection(1).XValues = "='xvaluesRef'!`$D`$2:`$D`$5" }
  'formula' = { param($ch, $ws) $ch.SeriesCollection(2).Formula = "=SERIES('formula'!`$D`$1,'formula'!`$A`$2:`$A`$5,'formula'!`$D`$2:`$D`$5,2)" }
  'plotOrder' = { param($ch, $ws) $ch.SeriesCollection(2).PlotOrder = 1 }
  'deleteFirst' = { param($ch, $ws) $ch.SeriesCollection(1).Delete() }
  'deleteThenAdd' = { param($ch, $ws) $ch.SeriesCollection(1).Delete(); $s = $ch.SeriesCollection().NewSeries(); $s.Name = "='deleteThenAdd'!`$D`$1"; $s.Values = "='deleteThenAdd'!`$D`$2:`$D`$5" }
  'threeMoveDown' = { param($ch, $ws) $ch.SetSourceData($ws.Range('A1:D5')); $ch.SeriesCollection(1).PlotOrder = 2 }
  'blanksZero' = { param($ch, $ws) $ch.DisplayBlanksAs = 2 }
  'blanksSpan' = { param($ch, $ws) $ch.DisplayBlanksAs = 3 }
  'hiddenShown' = { param($ch, $ws) $ch.PlotVisibleOnly = $false }
  'labelsThenSource' = { param($ch, $ws) $ch.SetElement(201); $ch.SetSourceData($ws.Range('A1:D5')) }
  'titleThenSource' = { param($ch, $ws) $ch.HasTitle = $true; $ch.ChartTitle.Text = 'Own'; $ch.SetSourceData($ws.Range('A1:B5')) }
  'colorThenSource' = { param($ch, $ws) $ch.SeriesCollection(2).Format.Fill.ForeColor.RGB = 255; $ch.SetSourceData($ws.Range('A1:D5')) }
  'colorThenRows' = { param($ch, $ws) $ch.SeriesCollection(2).Format.Fill.ForeColor.RGB = 255; $ch.PlotBy = 1 }
  'colorThenDelete' = { param($ch, $ws) $ch.SeriesCollection(2).Format.Fill.ForeColor.RGB = 255; $ch.SeriesCollection(1).Delete(); $ch.SetSourceData($ws.Range('A1:D5')) }
  'labelsThenNew' = { param($ch, $ws) $ch.SetElement(201); $s = $ch.SeriesCollection().NewSeries(); $s.Values = "='labelsThenNew'!`$D`$2:`$D`$5" }
  'labelsThenRows' = { param($ch, $ws) $ch.SetElement(201); $ch.PlotBy = 1 }
  'sourceShort' = { param($ch, $ws) $ch.SetSourceData($ws.Range('A1:D3')) }
  'rowsThenSource' = { param($ch, $ws) $ch.PlotBy = 1; $ch.SetSourceData($ws.Range('A1:D5')) }
  'valuesRelative' = { param($ch, $ws) $ch.SeriesCollection(2).Values = "=valuesRelative!D2:D5" }
  'xvaluesEmpty' = { param($ch, $ws) try { $ch.SeriesCollection(1).XValues = '' } catch { 'refused' } }
  'nameEmpty' = { param($ch, $ws) try { $ch.SeriesCollection(1).Name = '' } catch { 'refused' } }
  'lineNewSeries' = { param($ch, $ws) $ch.ChartType = 65; $s = $ch.SeriesCollection().NewSeries(); $s.Values = "='lineNewSeries'!`$D`$2:`$D`$5" }
  'lineSource' = { param($ch, $ws) $ch.ChartType = 65; $ch.SetSourceData($ws.Range('A1:D5')) }
  'pieSource' = { param($ch, $ws) $ch.ChartType = 5; $ch.SetSourceData($ws.Range('A1:D5')) }
  'scatterSource' = { param($ch, $ws) $ch.ChartType = -4169; $ch.SetSourceData($ws.Range('A1:D5')) }
}
$DATA = @(@('Region', 'Sales', 'Cost', 'Profit'), @('North', '120', '80', '40'), @('South', '95', '70', '25'), @('East', '140', '90', '50'), @('West', '60', '45', '15'))
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$read = [ordered]@{}
$formulas = [ordered]@{}
try {
  $wb = $x.Workbooks.Add()
  $k = 0
  foreach ($name in $CASES.Keys) {
    $ws = if ($k -eq 0) { $wb.Worksheets.Item(1) } else { $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $ws.Name = $name
    for ($r = 0; $r -lt $DATA.Count; $r++) { for ($c = 0; $c -lt 4; $c++) { $ws.Range("$([char](65 + $c))$($r + 1)").Formula = $DATA[$r][$c] } }
    $ws.Range('A1:C5').Select()
    $ch = $ws.Shapes.AddChart2(-1, 51, 300, 15, 360, 216).Chart
    & $CASES[$name] $ch $ws
    $f = @()
    for ($s = 1; $s -le $ch.SeriesCollection().Count; $s++) { $f += $ch.SeriesCollection($s).Formula }
    $formulas[$name] = $f
    $k++
  }
  $wb.Worksheets.Item(1).Activate()
  $file = Join-Path $Books 'chart-data.xlsx'
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
$json = [pscustomobject]@{ source = 'tools/excel-compare/excel-chart-data.ps1: Excel 365, a clustered column chart of A1:C5 (Region, Sales, Cost; Profit in D) on each sheet, one data change each through the chart object model; saved; each chart part and each series formula'; charts = $read; formulas = $formulas } | ConvertTo-Json -Depth 4
[IO.File]::WriteAllText($Out, $json, (New-Object Text.UTF8Encoding($false)))
'done'
