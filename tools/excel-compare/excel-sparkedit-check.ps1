# Excel reads back the sparklines the editor wrote: each .xlsx in a folder opened read-only in an invisible Excel, and
# every sparkline group of its first sheet read through COM (type, sparklines, options, colours with their theme
# colour and tint). For sparkedit.test.js (test/sparkedit-readback-excel.json). Nothing saved, no keys sent.
# usage: pwsh tools/excel-compare/excel-sparkedit-check.ps1 <folder> <out.json>
param([Parameter(Mandatory)][string]$In, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$In = (Resolve-Path $In).Path
Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);'
$x = New-Object -ComObject Excel.Application
$mine = [uint32]0
[void][W.U]::GetWindowThreadProcessId([IntPtr]$x.Hwnd, [ref]$mine)
Write-Host "Excel PID $mine"
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$files = @()
try {
  foreach ($f in Get-ChildItem $In -Filter *.xlsx) {
    $wb = $x.Workbooks.Open($f.FullName, 0, $true)
    $ws = $wb.Worksheets.Item(1)
    $all = $ws.Cells.SparklineGroups
    $groups = @()
    for ($k = 1; $k -le $all.Count; $k++) {
      $g = $all.Item($k)
      $p = $g.Points
      $cells = @(); for ($j = 1; $j -le $g.Count; $j++) { $s = $g.Item($j); $cells += "$($s.Location.Address($false, $false)) <- $($s.SourceData)" }
      $groups += [ordered]@{
        type = $g.Type; cells = $cells; lineWeight = $g.LineWeight; blanks = $g.DisplayBlanksAs; hidden = $g.DisplayHidden
        markers = $p.Markers.Visible; high = $p.Highpoint.Visible; low = $p.Lowpoint.Visible; first = $p.Firstpoint.Visible; last = $p.Lastpoint.Visible; negative = $p.Negative.Visible
        axis = $g.Axes.Horizontal.Axis.Visible; rtl = $g.Axes.Horizontal.RightToLeftPlotOrder
        minType = $g.Axes.Vertical.MinScaleType; maxType = $g.Axes.Vertical.MaxScaleType; min = $g.Axes.Vertical.CustomMinScaleValue; max = $g.Axes.Vertical.CustomMaxScaleValue
        dateRange = $g.DateRange; seriesColor = $g.SeriesColor.Color; seriesTheme = $g.SeriesColor.ThemeColor; seriesTint = $g.SeriesColor.TintAndShade
        markersTheme = $p.Markers.Color.ThemeColor; negativeColor = $p.Negative.Color.Color
      }
    }
    $files += [ordered]@{ file = $f.Name; groups = $groups }
    $wb.Close($false)
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  Start-Sleep -Milliseconds 500
  $proc = Get-Process -Id $mine -ErrorAction SilentlyContinue
  if ($proc -and $proc.MainWindowHandle -eq 0) { Stop-Process -Id $mine -Force }
}
$json = ConvertTo-Json -InputObject $files -Depth 8
[IO.File]::WriteAllText($Out, $json.Replace("`r`n", "`n") + "`n", (New-Object Text.UTF8Encoding $false))
