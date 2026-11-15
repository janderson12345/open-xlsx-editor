# What Excel writes when a chart is formatted, for chartformat.test.js: in an invisible Excel, a new workbook with the
# small table of excel-chart-edit.ps1 on each sheet and a clustered column chart of it (Shapes.AddChart2); each case one
# change through the chart's object model: an axis's title (Axis.HasTitle, AxisTitle.Text), a series' or point's fill
# and line (Format.Fill / Format.Line .ForeColor.RGB), the chart and plot area's fill and border, the fonts of the
# title, axes, legend and data labels, a series drawn as a line or on the secondary axis (Series.ChartType,
# Series.AxisGroup). The workbook saved outside the repository, each sheet's chart part read back, the secondary axis's
# scale read, and the workbook exported as a PDF (one page a sheet) for the drawing.
# Only the Excel this script starts is used and quit. No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-chart-format.ps1 <out.json> <folder for the saved workbook and PDF>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
# (Colours as COM takes them: red + green * 256 + blue * 65536.)
function rgb([string]$hex) { [Convert]::ToInt32($hex.Substring(0, 2), 16) + 256 * [Convert]::ToInt32($hex.Substring(2, 2), 16) + 65536 * [Convert]::ToInt32($hex.Substring(4, 2), 16) }
function font1($f, $name, $size, $bold, $hex) { $f.Name = $name; $f.Size = $size; $f.Bold = $bold; $f.Color = (rgb $hex) }
function font2($f, $name, $size, $bold, $hex) { if ($name) { $f.Name = $name }; if ($size) { $f.Size = $size }; if ($null -ne $bold) { $f.Bold = $bold }; if ($hex) { $f.Fill.ForeColor.RGB = (rgb $hex) } }
$CASES = [ordered]@{
  'none' = { param($ch) }
  'catTitle' = { param($ch) $a = $ch.Axes(1); $a.HasTitle = $true; $a.AxisTitle.Text = 'Region' }
  'valTitle' = { param($ch) $a = $ch.Axes(2); $a.HasTitle = $true; $a.AxisTitle.Text = 'Amount' }
  'catTitleAuto' = { param($ch) $ch.Axes(1).HasTitle = $true }
  'titlesOff' = { param($ch) foreach ($k in 1, 2) { $a = $ch.Axes($k); $a.HasTitle = $true; $a.AxisTitle.Text = 'X' }; foreach ($k in 1, 2) { $ch.Axes($k).HasTitle = $false } }
  'seriesFill' = { param($ch) $ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = (rgb 'C00000') }
  'seriesLine' = { param($ch) $l = $ch.SeriesCollection(1).Format.Line; $l.Visible = -1; $l.ForeColor.RGB = (rgb '7030A0') }
  'seriesNoLine' = { param($ch) $l = $ch.SeriesCollection(1).Format.Line; $l.Visible = -1; $l.ForeColor.RGB = (rgb '7030A0'); $l.Visible = 0 }
  'pointFill' = { param($ch) $ch.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = (rgb '00B050') }
  'chartFill' = { param($ch) $ch.ChartArea.Format.Fill.ForeColor.RGB = (rgb 'FFF2CC') }
  'chartLine' = { param($ch) $ch.ChartArea.Format.Line.ForeColor.RGB = (rgb '1F4E79') }
  'chartNoLine' = { param($ch) $ch.ChartArea.Format.Line.Visible = 0 }
  'plotFill' = { param($ch) $ch.PlotArea.Format.Fill.ForeColor.RGB = (rgb 'DDEBF7') }
  'plotLine' = { param($ch) $l = $ch.PlotArea.Format.Line; $l.Visible = -1; $l.ForeColor.RGB = (rgb '808080') }
  'titleFontName' = { param($ch) font2 $ch.ChartTitle.Format.TextFrame2.TextRange.Font 'Arial' $null $null $null }
  'titleFontSize' = { param($ch) font2 $ch.ChartTitle.Format.TextFrame2.TextRange.Font $null 18 $null $null }
  'titleBold' = { param($ch) font2 $ch.ChartTitle.Format.TextFrame2.TextRange.Font $null $null -1 $null }
  'titleColor' = { param($ch) font2 $ch.ChartTitle.Format.TextFrame2.TextRange.Font $null $null $null 'C00000' }
  'titleTextFont' = { param($ch) $ch.ChartTitle.Text = 'Sales'; font2 $ch.ChartTitle.Format.TextFrame2.TextRange.Font 'Arial' 18 -1 'C00000' }
  'catFont' = { param($ch) font1 $ch.Axes(1).TickLabels.Font 'Arial' 11 $true '1F4E79' }
  'valFont' = { param($ch) font1 $ch.Axes(2).TickLabels.Font 'Arial' 11 $true '1F4E79' }
  'legendFont' = { param($ch) font2 $ch.Legend.Format.TextFrame2.TextRange.Font 'Arial' 11 -1 '1F4E79' }
  'labelFont' = { param($ch) $ch.SetElement(201); font2 $ch.SeriesCollection(1).DataLabels().Format.TextFrame2.TextRange.Font 'Arial' 11 -1 '1F4E79' }
  'catTitleFont' = { param($ch) $a = $ch.Axes(1); $a.HasTitle = $true; $a.AxisTitle.Text = 'Region'; font2 $a.AxisTitle.Format.TextFrame2.TextRange.Font 'Arial' 12 -1 'C00000' }
  'titleFont1' = { param($ch) font1 $ch.ChartTitle.Font 'Arial' 18 $true 'C00000' }
  'titleFontName1' = { param($ch) $ch.ChartTitle.Font.Name = 'Arial' }
  'catFontSize1' = { param($ch) $ch.Axes(1).TickLabels.Font.Size = 12 }
  'legendFont1' = { param($ch) font1 $ch.Legend.Font 'Arial' 11 $true '1F4E79' }
  'labelFont1' = { param($ch) $ch.SetElement(201); font1 $ch.SeriesCollection(1).DataLabels().Font 'Arial' 11 $true '1F4E79' }
  'catTitleFont1' = { param($ch) $a = $ch.Axes(1); $a.HasTitle = $true; $a.AxisTitle.Text = 'Region'; font1 $a.AxisTitle.Font 'Arial' 12 $true 'C00000' }
  'lineColor' = { param($ch) $ch.ChartType = 4; $ch.SeriesCollection(1).Format.Line.ForeColor.RGB = (rgb 'C00000') }
  'pieSlice' = { param($ch) $ch.ChartType = 5; $ch.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = (rgb '00B050') }
  'comboLine' = { param($ch) $ch.SeriesCollection(2).ChartType = 4 }
  'comboLineMarkers' = { param($ch) $ch.SeriesCollection(2).ChartType = 65 }
  'secondary' = { param($ch) $ch.SeriesCollection(2).AxisGroup = 2 }
  'comboSecondary' = { param($ch) $s = $ch.SeriesCollection(2); $s.ChartType = 4; $s.AxisGroup = 2 }
  'secondaryBack' = { param($ch) $s = $ch.SeriesCollection(2); $s.AxisGroup = 2; $s.AxisGroup = 1 }
  'comboBack' = { param($ch) $s = $ch.SeriesCollection(2); $s.ChartType = 4; $s.ChartType = 51 }
  'comboSecondaryBack' = { param($ch) $s = $ch.SeriesCollection(2); $s.ChartType = 4; $s.AxisGroup = 2; $s.AxisGroup = 1; $s.ChartType = 51 }
  'comboArea' = { param($ch) $ch.SeriesCollection(2).ChartType = 1 }
  'comboOwnColor' = { param($ch) $s = $ch.SeriesCollection(2); $s.Format.Fill.ForeColor.RGB = (rgb 'C00000'); $s.ChartType = 4 }
  'secondaryTitle' = { param($ch) $ch.SeriesCollection(2).AxisGroup = 2; $a = $ch.Axes(2, 2); $a.HasTitle = $true; $a.AxisTitle.Text = 'Cost' }
}
$DATA = @(@('Region', 'Sales', 'Cost'), @('North', '120', '80'), @('South', '95', '70'), @('East', '140', '90'), @('West', '60', '45'))
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$read = [ordered]@{}
$scales = [ordered]@{}
$types = [ordered]@{}
try {
  $wb = $x.Workbooks.Add()
  $k = 0
  foreach ($name in $CASES.Keys) {
    $ws = if ($k -eq 0) { $wb.Worksheets.Item(1) } else { $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $ws.Name = $name
    for ($r = 0; $r -lt $DATA.Count; $r++) { for ($c = 0; $c -lt 3; $c++) { $ws.Range("$([char](65 + $c))$($r + 1)").Formula = $DATA[$r][$c] } }
    [void]$ws.Range('A1:C5').Select()
    $ch = $ws.Shapes.AddChart2(-1, 51, 240, 15, 360, 216).Chart
    & $CASES[$name] $ch
    $types[$name] = @(foreach ($s in $ch.SeriesCollection()) { [pscustomobject]@{ type = $s.ChartType; group = $s.AxisGroup } })
    if ($ch.HasAxis(2, 2)) { $a = $ch.Axes(2, 2); $p = $ch.Axes(2, 1); $scales[$name] = [pscustomobject]@{ min = $a.MinimumScale; max = $a.MaximumScale; major = $a.MajorUnit; primary = @($p.MinimumScale, $p.MaximumScale, $p.MajorUnit) } }
    $ws.PageSetup.Orientation = 2
    $k++
  }
  $wb.Worksheets.Item(1).Activate()
  $file = Join-Path $Books 'chart-format.xlsx'
  $pdf = Join-Path $Books 'chart-format.pdf'
  foreach ($f in $file, $pdf) { if (Test-Path $f) { Remove-Item $f } }
  $wb.SaveAs($file, 51)
  $wb.ExportAsFixedFormat(0, $pdf)
  $wb.Close($false)
  $zip = [IO.Compression.ZipFile]::OpenRead($file)
  $i = 1
  foreach ($name in $CASES.Keys) {
    $s = New-Object IO.StreamReader($zip.GetEntry("xl/charts/chart$i.xml").Open()); $read[$name] = $s.ReadToEnd(); $s.Close()
    $i++
  }
  $zip.Dispose()
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
[pscustomobject]@{ source = 'tools/excel-compare/excel-chart-format.ps1: Excel 365, a clustered column chart of A1:C5 (Region, Sales, Cost) on each sheet, one change each through the chart object model; saved; each chart part'; charts = $read; types = $types; secondary = $scales } | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 $Out
'done'
