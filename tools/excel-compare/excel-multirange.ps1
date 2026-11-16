# What Excel's commands do with a selection of several areas (docs/GOTOSPECIAL.md). Each case: a new workbook, A1:F12
# filled with numbers (row * 10 + column), a few cells set up, the areas selected, then the ribbon's command run
# (CommandBars.ExecuteMso, as its button does; whether the button is enabled from GetEnabledMso, and with `probe`
# only that) or the object
# model's call where the ribbon's would only open a box. A box Excel puts up is read and closed by dialog-watch.ps1
# (window messages to this Excel alone, nothing through the desktop). Read after: the selection, A1:H14's values and
# formulas, rows 1-14 hidden and heights, columns A-H hidden and widths, merges, the filter, tables, charts' series.
# Starts an Excel of its own, its window off the screen, each case in a new workbook closed without saving.
# usage: pwsh -NoProfile -File tools/excel-compare/excel-multirange.ps1 tools/excel-compare/multirange-cases.json test/multirange-excel.json
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out, [string]$Only = '')
$ErrorActionPreference = 'Stop'
$list = Get-Content $Cases -Raw | ConvertFrom-Json
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$x = New-Object -ComObject Excel.Application
# (Alerts on: Excel's refusals are what's measured; the watcher answers each box with its first button.)
$x.DisplayAlerts = $false
$x.Visible = $true
$x.WindowState = -4143
$x.Left = -20000
$x.Top = -20000
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class XP { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid); }'
[uint32]$xpid = 0
[void][XP]::GetWindowThreadProcessId([IntPtr]$x.Hwnd, [ref]$xpid)
$tmp = Join-Path ([IO.Path]::GetTempPath()) "oxe-multirange-$xpid"
New-Item -ItemType Directory -Force $tmp | Out-Null
$log = Join-Path $tmp 'boxes.log'; $stop = Join-Path $tmp 'stop'
Remove-Item $log, $stop -ErrorAction SilentlyContinue
$watch = Start-Process pwsh -ArgumentList @('-NoProfile', '-File', (Join-Path $here 'dialog-watch.ps1'), $xpid, $log, $stop) -PassThru -WindowStyle Hidden
$answers = New-Object System.Collections.Generic.List[object]
function Props($o) { if ($null -eq $o) { @() } else { $o.PSObject.Properties } }
function Save { [IO.File]::WriteAllText($Out, ([ordered]@{ excel = 'Excel.Application (one of its own, its window off the screen): CommandBars.ExecuteMso, or the object model; boxes read by dialog-watch.ps1'; cases = $answers } | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding $false)) }
function Boxes { if (Test-Path $log) { @(Get-Content $log | ForEach-Object { $_ | ConvertFrom-Json }) } else { @() } }
try {
  foreach ($case in $list) {
    if ($Only -and $case.name -notmatch $Only) { continue }
    Write-Host "case: $($case.name)"
    $wb = $x.Workbooks.Add()
    try {
      $ws = $wb.Worksheets.Item(1)
      for ($r = 1; $r -le 12; $r++) { for ($c = 1; $c -le 6; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 10 + $c } }
      $s = $case.setup
      foreach ($p in Props $s.cells) { $ws.Range($p.Name).Formula = [string]$p.Value }
      if ($s.merge) { foreach ($m in $s.merge) { [void]$ws.Range([string]$m).Merge() } }
      if ($s.hideRows) { $ws.Rows([string]$s.hideRows).Hidden = $true }
      if ($s.hideCols) { $ws.Columns([string]$s.hideCols).Hidden = $true }
      $x.DisplayAlerts = $true
      [void]$ws.Activate()
      [void]$ws.Range([string]$case.select).Select()
      if ($case.active) { [void]$ws.Range([string]$case.active).Activate() }
      $before = (Boxes).Count
      $a = [ordered]@{ name = [string]$case.name; select = [string]$case.select }
      try {
        if ($case.mso) {
          $a.enabled = [bool]$x.CommandBars.GetEnabledMso([string]$case.mso)
          if ($a.enabled -and -not $case.probe) { [void]$x.CommandBars.ExecuteMso([string]$case.mso) }
        } else {
          $sel = $x.Selection
          $a.result = "$(Invoke-Expression ([string]$case.om))"
        }
      } catch { $a.error = "$($_.Exception.Message)" }
      Start-Sleep -Milliseconds 700
      $a.boxes = @((Boxes) | Select-Object -Skip $before)
      $a.selection = [string]$x.Selection.Address($false, $false)
      $cells = [ordered]@{}
      foreach ($c in $ws.Range('A1:H14').Cells) {
        $one = [ordered]@{}
        if ($c.HasFormula) { $one.f = [string]$c.Formula }
        $v = $c.Value2
        if ($null -ne $v) { $one.v = if ($v -is [double] -or $v -is [bool]) { $v } else { "$v" } }
        if ($c.NumberFormat -ne 'General') { $one.fmt = [string]$c.NumberFormat }
        if ($c.MergeCells -and $c.MergeArea.Cells.Item(1, 1).Address() -eq $c.Address()) { $one.merge = [string]$c.MergeArea.Address($false, $false) }
        if ($one.Count) { $cells[$c.Address($false, $false)] = $one }
      }
      $a.cells = $cells
      $a.hiddenRows = @(1..14 | Where-Object { $ws.Rows($_).Hidden })
      $a.hiddenCols = @(1..8 | Where-Object { $ws.Columns($_).Hidden })
      $a.heights = @(1..14 | ForEach-Object { [double]$ws.Rows($_).RowHeight })
      $a.widths = @(1..8 | ForEach-Object { [double]$ws.Columns($_).ColumnWidth })
      if ($ws.AutoFilterMode) { $a.filter = [string]$ws.AutoFilter.Range.Address($false, $false) }
      $a.tables = @(foreach ($lo in $ws.ListObjects) { [string]$lo.Range.Address($false, $false) })
      $a.names = @(foreach ($n in $wb.Names) { "$($n.Name) $($n.RefersTo)" })
      $a.charts = @(foreach ($co in $ws.ChartObjects()) { , @(foreach ($sr in $co.Chart.SeriesCollection()) { [string]$sr.Formula }) })
      $a.conditions = @(foreach ($fc in $ws.Cells.FormatConditions) { [string]$fc.AppliesTo.Address($false, $false) })
      $x.DisplayAlerts = $false
      $answers.Add($a)
      Save
    } finally { $wb.Close($false) }
  }
} finally {
  New-Item -ItemType File $stop -Force | Out-Null
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  Start-Sleep -Milliseconds 500
  if (-not $watch.HasExited) { $watch.Kill() }
}
