# How Excel saves sparklines, for sparklines.test.js and docs/SPARKLINES.md: in an invisible Excel, a new workbook
# with made-up numbers; sparkline groups of each type (line, column, win/loss) added with Range.SparklineGroups.Add and
# each option set through COM (colours, markers, high, low, first and last points, line weight, empty cells, hidden
# data, axis minimum and maximum, the date axis, right to left). The workbook saved outside the repository and its
# sheet's x14:sparklineGroups extension read back; each group's options as COM reports them; and the sheet exported as
# PDF by Excel, the reference the drawing is compared with. No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-sparklines.ps1 <out.json> <folder for the saved workbook and PDF>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);'
# Made-up numbers, one row per sparkline (B..M), and empty cells where a row has $null.
$ROWS = @(
  @(3, 5, 2, 8, 6, 9, 4, 7, 5, 10, 6, 8),
  @(3, -2, 5, 8, -4, 9, 1, 7, -6, 10, 2, 4),
  @(4, -3, 6, 2, -5, 7, 3, -1, 8, 5, -2, 6),
  @(5, 9, 3, 7, 11, 2, 8, 6, 10, 4, 12, 1),
  @(1, -1, 2, -3, 0, 4, -2, 1, -1, 3, -4, 2),
  @(-2, -5, 1, 4, -3, 6, 2, -4, 5, -1, 3, 0),
  @(4, 6, $null, 8, 5, $null, $null, 7, 9, 3, $null, 6),
  @(4, 6, $null, 8, 5, $null, $null, 7, 9, 3, $null, 6),
  @(4, 6, $null, 8, 5, $null, $null, 7, 9, 3, $null, 6),
  @(2, 4, 3, 5, 4, 6, 5, 7, 6, 8, 7, 9),
  @(10, 14, 12, 16, 15, 18, 17, 20, 19, 22, 21, 24),
  @(3, 8, 1, 12, 6, 15, 4, 9, 2, 11, 7, 14),
  @(5, 7, 6, 9, 8, 11, 10, 9, 12, 14, 13, 15),
  @(3, 5, 2, 8, 6, 9, 4, 7, 5, 10, 6, 8),
  @(3, 5, 2, 8, 6, 9, 4, 7, 5, 10, 6, 8),
  @(3, 5, 2, 8, 6, 9, 4, 7, 5, 10, 6, 8)
)
# Dates for the date axis (row 20), uneven gaps: serials.
$DATES = @(45658, 45659, 45661, 45665, 45666, 45670, 45671, 45680, 45681, 45682, 45690, 45700)
# Each group: its location, its data, its type (1 line, 2 column, 3 win/loss) and a block setting its options.
$GROUPS = [ordered]@{
  'line'          = @('O2', 'B2:M2', 1, { param($g) })
  'lineMarked'    = @('O3', 'B3:M3', 1, { param($g)
      $g.LineWeight = 2.25; $g.SeriesColor.Color = 0x8B4513
      $p = $g.Points; $p.Markers.Visible = $true; $p.Markers.Color.Color = 0x00A5FF
      $p.Negative.Visible = $true; $p.Negative.Color.Color = 0x0000FF
      $p.Highpoint.Visible = $true; $p.Highpoint.Color.Color = 0x00FF00
      $p.Lowpoint.Visible = $true; $p.Lowpoint.Color.Color = 0xFF00FF
      $p.Firstpoint.Visible = $true; $p.Firstpoint.Color.Color = 0xFFFF00
      $p.Lastpoint.Visible = $true; $p.Lastpoint.Color.Color = 0x808080 })
  'column'        = @('O4', 'B4:M4', 2, { param($g) $g.Points.Negative.Visible = $true })
  'columnPoints'  = @('O5', 'B5:M5', 2, { param($g)
      $p = $g.Points; $p.Highpoint.Visible = $true; $p.Lowpoint.Visible = $true; $p.Firstpoint.Visible = $true; $p.Lastpoint.Visible = $true })
  'winloss'       = @('O6', 'B6:M6', 3, { param($g) $g.Points.Negative.Visible = $true })
  'lineAxis'      = @('O7', 'B7:M7', 1, { param($g) $g.Axes.Horizontal.Axis.Visible = $true; $g.Axes.Horizontal.Axis.Color.Color = 0x0000C0 })
  'emptyGaps'     = @('O8', 'B8:M8', 1, { param($g) $g.DisplayBlanksAs = 1 })
  'emptyZero'     = @('O9', 'B9:M9', 1, { param($g) $g.DisplayBlanksAs = 2 })
  'emptyConnect'  = @('O10', 'B10:M10', 1, { param($g) $g.DisplayBlanksAs = 3 })
  'groupScale'    = @('O11:O12', 'B11:M12', 2, { param($g) $g.Axes.Vertical.MinScaleType = 1; $g.Axes.Vertical.MaxScaleType = 1 })
  'customScale'   = @('O13', 'B13:M13', 1, { param($g)
      $g.Axes.Vertical.MinScaleType = 3; $g.Axes.Vertical.CustomMinScaleValue = 0
      $g.Axes.Vertical.MaxScaleType = 3; $g.Axes.Vertical.CustomMaxScaleValue = 20 })
  'dateAxis'      = @('O14', 'B14:M14', 1, { param($g) $g.DateRange = 'B20:M20'; $g.Points.Markers.Visible = $true })
  'rightToLeft'   = @('O15', 'B15:M15', 1, { param($g) $g.Axes.Horizontal.RightToLeftPlotOrder = $true; $g.Points.Markers.Visible = $true })
  'hiddenOff'     = @('O16', 'B16:N16', 1, { param($g) })
  'hiddenOn'      = @('O17', 'B17:N17', 1, { param($g) $g.DisplayHidden = $true })
}
# What COM says of a group.
$state = { param($g)
  $p = $g.Points
  $col = { param($c) try { [int]$c.Color } catch { $null } }
  [ordered]@{
    type = $g.Type; location = $g.Location.Address($false, $false); source = $g.SourceData; dateRange = $g.DateRange
    lineWeight = $g.LineWeight; series = & $col $g.SeriesColor; blanks = $g.DisplayBlanksAs; hidden = $g.DisplayHidden
    negative = @($p.Negative.Visible, (& $col $p.Negative.Color)); markers = @($p.Markers.Visible, (& $col $p.Markers.Color))
    high = @($p.Highpoint.Visible, (& $col $p.Highpoint.Color)); low = @($p.Lowpoint.Visible, (& $col $p.Lowpoint.Color))
    first = @($p.Firstpoint.Visible, (& $col $p.Firstpoint.Color)); last = @($p.Lastpoint.Visible, (& $col $p.Lastpoint.Color))
    axis = @($g.Axes.Horizontal.Axis.Visible, (& $col $g.Axes.Horizontal.Axis.Color)); rtl = $g.Axes.Horizontal.RightToLeftPlotOrder
    minType = $g.Axes.Vertical.MinScaleType; maxType = $g.Axes.Vertical.MaxScaleType
    min = $g.Axes.Vertical.CustomMinScaleValue; max = $g.Axes.Vertical.CustomMaxScaleValue
  }
}
$x = New-Object -ComObject Excel.Application
# (This Excel's own process, from its window handle, so only it is stopped if Quit leaves it running.)
$mine = [uint32]0
[void][W.U]::GetWindowThreadProcessId([IntPtr]$x.Hwnd, [ref]$mine)
Write-Host "Excel PID $mine"
$x.Visible = $false; $x.DisplayAlerts = $false
$result = [ordered]@{}
try {
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  for ($i = 0; $i -lt $ROWS.Count; $i++) {
    for ($j = 0; $j -lt 12; $j++) { $v = $ROWS[$i][$j]; if ($null -ne $v) { $ws.Cells.Item($i + 2, $j + 2).Value2 = $v } }
  }
  # (Column N: an outlier in the hidden column, for rows 16 and 17 only.)
  $ws.Range('N16').Value2 = 30; $ws.Range('N17').Value2 = 30
  for ($j = 0; $j -lt 12; $j++) { $ws.Cells.Item(20, $j + 2).Value2 = $DATES[$j] }
  $ws.Range('O1').ColumnWidth = 30
  $ws.Range('O2:O17').RowHeight = 36
  $made = [ordered]@{}
  foreach ($name in $GROUPS.Keys) {
    $at, $src, $type, $set = $GROUPS[$name]
    $g = $ws.Range($at).SparklineGroups.Add($type, $src)
    & $set $g
    $made[$name] = $g
  }
  $ws.Range('N1').EntireColumn.Hidden = $true
  foreach ($name in $made.Keys) { $result[$name] = & $state $made[$name] }
  $file = Join-Path $Books 'sparklines-excel.xlsx'
  if (Test-Path $file) { Remove-Item $file }
  $wb.SaveAs($file, 51)
  # Excel's own PDF of the sheet: xlTypePDF = 0.
  $pdf = Join-Path $Books 'sparklines-excel.pdf'
  $ws.ExportAsFixedFormat(0, $pdf)
  # Sizes: the same four numbers (-2, 5, 1, 4) drawn in cells of three widths (8, 20, 45 characters) and four heights
  # (12, 24, 48, 96 points), as columns (sheet Columns) and as lines with markers (sheet Lines); and lines with
  # markers of five weights (0.25 to 6 points) in cells 45 by 48 (sheet Weights). Each sheet printed to its own PDF.
  foreach ($kind in 'Columns', 'Lines', 'Weights') {
    $s = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
    $s.Name = $kind
    $vals = -2, 5, 1, 4
    for ($j = 0; $j -lt 4; $j++) { $s.Cells.Item(1, $j + 1).Value2 = $vals[$j] }
    if ($kind -eq 'Weights') {
      $s.Range('F1').ColumnWidth = 45
      $w = 0.25, 0.75, 1.5, 3, 6
      for ($i = 0; $i -lt 5; $i++) {
        $s.Rows.Item($i + 2).RowHeight = 48
        $g = $s.Range("F$($i + 2)").SparklineGroups.Add(1, 'A1:D1'); $g.LineWeight = [double]$w[$i]; $g.Points.Markers.Visible = $true
      }
    } else {
      $widths = 8, 20, 45; $heights = 12, 24, 48, 96
      for ($c = 0; $c -lt 3; $c++) { $s.Columns.Item($c + 6).ColumnWidth = $widths[$c] }
      for ($r = 0; $r -lt 4; $r++) {
        $s.Rows.Item($r + 2).RowHeight = $heights[$r]
        for ($c = 0; $c -lt 3; $c++) {
          $g = $s.Cells.Item($r + 2, $c + 6).SparklineGroups.Add($(if ($kind -eq 'Columns') { 2 } else { 1 }), 'A1:D1')
          if ($kind -eq 'Lines') { $g.Points.Markers.Visible = $true }
        }
      }
    }
    $s.ExportAsFixedFormat(0, (Join-Path $Books "sparklines-$($kind.ToLower()).pdf"))
  }
  $wb.Close($false)
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  Start-Sleep -Milliseconds 500
  $p = Get-Process -Id $mine -ErrorAction SilentlyContinue
  if ($p -and $p.MainWindowHandle -eq 0) { Stop-Process -Id $mine -Force }
}
# The saved sheet's sparkline extension, as text.
$zip = [IO.Compression.ZipFile]::OpenRead($file)
try {
  $s = New-Object IO.StreamReader($zip.GetEntry('xl/worksheets/sheet1.xml').Open()); $sheet = $s.ReadToEnd(); $s.Close()
} finally { $zip.Dispose() }
$m = [regex]::Match($sheet, '<ext [^>]*uri="\{05C60535-1F16-4fd2-B633-F4F36F0B64E0\}".*?</ext>', 'Singleline')
$answer = [ordered]@{ groups = $result; ext = $m.Value }
$json = ConvertTo-Json -InputObject $answer -Depth 6
[IO.File]::WriteAllText($Out, $json + "`n", (New-Object Text.UTF8Encoding $false))
