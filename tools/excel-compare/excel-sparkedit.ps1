# How Excel makes and edits sparklines, for sparkedit.test.js and docs/SPARKLINES.md: in an invisible Excel, new
# workbooks with made-up numbers. (1) Range.SparklineGroups.Add for pairs of location and data shapes: the sparklines
# made (each one's cell and source) or Excel's refusal in its words. (2) On a sheet with a group of three (O2:O4) and
# one of its own (O6): Group, Ungroup, Clear and ClearGroups as the Sparkline tab runs them, each from the same
# start, and the groups read back; a theme colour set through ThemeColor and TintAndShade; each workbook saved and its
# x14:sparklineGroups kept. No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-sparkedit.ps1 <out.json> <folder for the saved workbooks>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);'
# Location, data (type 1, a line, unless a third is given).
$ADDS = @(
  @('O2', 'B2:M2'), @('O2:O3', 'B2:M3'), @('B6:M6', 'B2:M3'), @('O2:O4', 'B2:D4'), @('B6:D6', 'B2:D4'),
  @('O2:O3', 'B2:M2'), @('O2:P3', 'B2:M3'), @('O2', 'B2:M3'), @('O2', 'B2'), @('O2', "'My data'!B2:M2"),
  @('B2', 'B2:M2'), @('O2:O13', 'B2:B13'), @('O2', 'B2:B13'), @('O2:P2', 'B2:C13'), @('', 'B2:M2'), @('O2', '')
)
# The groups of a sheet: [{ type, cells: ['O2 <- B2:M2', ...] }], newest first as Excel lists them.
$groupsOf = { param($ws)
  $all = $ws.Cells.SparklineGroups
  $list = @()
  for ($k = 1; $k -le $all.Count; $k++) {
    $g = $all.Item($k)
    $cells = @(); for ($j = 1; $j -le $g.Count; $j++) { $s = $g.Item($j); $cells += "$($s.Location.Address($false, $false)) <- $($s.SourceData)" }
    $list += [ordered]@{ type = $g.Type; markers = $g.Points.Markers.Visible; series = $g.SeriesColor.Color; cells = $cells }
  }
  , $list
}
$extOf = { param($file)
  $zip = [IO.Compression.ZipFile]::OpenRead($file)
  try { $s = New-Object IO.StreamReader($zip.GetEntry('xl/worksheets/sheet1.xml').Open()); $t = $s.ReadToEnd(); $s.Close() } finally { $zip.Dispose() }
  [regex]::Match($t, '<x14:sparklineGroups.*?</x14:sparklineGroups>', 'Singleline').Value
}
$fill = { param($ws)
  for ($r = 2; $r -le 13; $r++) { for ($c = 2; $c -le 13; $c++) { $ws.Cells.Item($r, $c).Value2 = (($r * 7 + $c * 3) % 11) - 3 } }
}
$x = New-Object -ComObject Excel.Application
$mine = [uint32]0
[void][W.U]::GetWindowThreadProcessId([IntPtr]$x.Hwnd, [ref]$mine)
Write-Host "Excel PID $mine"
$x.Visible = $false; $x.DisplayAlerts = $false
$result = [ordered]@{ adds = @(); edits = [ordered]@{} }
try {
  foreach ($case in $ADDS) {
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    & $fill $ws
    $other = $wb.Worksheets.Add([Type]::Missing, $ws); $other.Name = 'My data'; & $fill $other
    [void]$ws.Activate()
    $answer = [ordered]@{ location = $case[0]; data = $case[1] }
    try {
      if (-not $case[0]) { throw 'no location' }
      # (Not through an if expression: PowerShell would enumerate a range of several cells into an array.)
      $at = $ws.Range($case[0])
      [void]$at.SparklineGroups.Add(1, $case[1])
      $answer.groups = & $groupsOf $ws
    } catch { $answer.error = ($_.Exception.InnerException ?? $_.Exception).Message }
    $result.adds += $answer
    $wb.Close($false)
  }
  # Edits, each from the same start: O2:O4 a line group over B2:M4; O6 a column with markers off, its own group.
  $EDITS = [ordered]@{
    start = { param($ws) }
    group = { param($ws) [void]$ws.Range('O3,O6').SparklineGroups.Group($ws.Range('O6')) }
    groupAtLine = { param($ws) [void]$ws.Range('O3,O6').SparklineGroups.Group($ws.Range('O3')) }
    ungroup = { param($ws) [void]$ws.Range('O3').SparklineGroups.Ungroup() }
    ungroupTwo = { param($ws) [void]$ws.Range('O2:O3').SparklineGroups.Ungroup() }
    addOver = { param($ws) [void]$ws.Range('O3').SparklineGroups.Add(2, 'B3:M3') }
    weight = { param($ws) $g = $ws.Range('O2').SparklineGroups.Item(1); $g.LineWeight = 1.5; $g.Axes.Vertical.MinScaleType = 3; $g.Axes.Vertical.CustomMinScaleValue = -2; $g.DisplayBlanksAs = 2 }
    clear = { param($ws) [void]$ws.Range('O3').SparklineGroups.Clear() }
    clearGroups = { param($ws) [void]$ws.Range('O3').SparklineGroups.ClearGroups() }
    clearAll = { param($ws) [void]$ws.Range('O2:O6').SparklineGroups.Clear() }
    themeColor = { param($ws) $g = $ws.Range('O2').SparklineGroups.Item(1); $g.SeriesColor.ThemeColor = 5; $g.SeriesColor.TintAndShade = -0.25; $g.Points.Markers.Color.ThemeColor = 10 }
    modify = { param($ws) $g = $ws.Range('O2').SparklineGroups.Item(1); $g.Modify($ws.Range('P2:P4'), 'C2:M4') }
    single = { param($ws) $ws.Range('O3').SparklineGroups.Item(1).Item(1).ModifySourceData('C3:H3') }
  }
  foreach ($name in $EDITS.Keys) {
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    & $fill $ws
    [void]$ws.Range('O2:O4').SparklineGroups.Add(1, 'B2:M4')
    $g = $ws.Range('O6').SparklineGroups.Add(2, 'B6:M6'); $g.SeriesColor.Color = 0x0000FF
    $answer = [ordered]@{}
    try { & $EDITS[$name] $ws; $answer.groups = & $groupsOf $ws } catch { $answer.error = ($_.Exception.InnerException ?? $_.Exception).Message }
    $file = Join-Path $Books "sparkedit-$name.xlsx"
    if (Test-Path $file) { Remove-Item $file }
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $answer.ext = & $extOf $file
    $result.edits[$name] = $answer
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  Start-Sleep -Milliseconds 500
  $p = Get-Process -Id $mine -ErrorAction SilentlyContinue
  if ($p -and $p.MainWindowHandle -eq 0) { Stop-Process -Id $mine -Force }
}
$json = ConvertTo-Json -InputObject $result -Depth 8
[IO.File]::WriteAllText($Out, $json.Replace("`r`n", "`n") + "`n", (New-Object Text.UTF8Encoding $false))
