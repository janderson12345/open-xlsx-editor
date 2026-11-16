# Excel's AutoSum (and Average, Count Numbers, Max, Min) on the cases of autosum2-cases.mjs: as excel-autosum.ps1, with
# a filtered list (Range.AutoFilter), a table (ListObjects.Add, its own filter), rows hidden by hand and selections of
# several areas set up first. The ribbon's command needs a window, so this starts an Excel of its own, its window off the
# screen (minimized, it refused the command) and left alone (no key sent, no other workbook touched), each case in a new workbook closed without saving, and quits
# it at the end. Only selections where Excel enters the formulas at once: one cell, or nothing to sum, would leave it
# editing the cell and answering nothing.
# usage: pwsh -NoProfile -File tools/excel-compare/excel-autosum2.ps1 tools/excel-compare/autosum2-cases.json test/autosum2-excel.json
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$list = Get-Content $Cases -Raw | ConvertFrom-Json
$ids = @{ sum = 'AutoSum'; average = 'AutoSumAverage'; count = 'AutoSumCount'; max = 'AutoSumMax'; min = 'AutoSumMin' }
$x = New-Object -ComObject Excel.Application
$x.DisplayAlerts = $false
$x.Visible = $true
# (Normal size, as a minimized window refused the command, but placed off the screen.)
$x.WindowState = -4143
$x.Left = -20000
$x.Top = -20000
$answers = New-Object System.Collections.Generic.List[object]
function Props($o) { if ($null -eq $o) { @() } else { $o.PSObject.Properties } }
function Save { [IO.File]::WriteAllText($Out, ([ordered]@{ excel = 'Excel.Application (one of its own, its window off the screen): CommandBars.ExecuteMso AutoSum'; cases = $answers } | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding $false)) }
try {
  foreach ($case in $list) {
    Write-Host "case: $($case.name)"
    $wb = $x.Workbooks.Add()
    try {
      $ws = $wb.Worksheets.Item(1)
      $s = $case.setup
      foreach ($p in Props $s.cells) { $ws.Range($p.Name).Formula = [string]$p.Value }
      if ($s.autofilter) {
        $r = $ws.Range([string]$s.autofilter.range)
        if ($s.autofilter.field) { [void]$r.AutoFilter([int]$s.autofilter.field, [string]$s.autofilter.criteria) } else { [void]$r.AutoFilter() }
      }
      if ($s.table) {
        $lo = $ws.ListObjects.Add(1, $ws.Range([string]$s.table.range), [Type]::Missing, 1)
        $lo.Name = [string]$s.table.name
        if ($s.tableFilter) { [void]$lo.Range.AutoFilter([int]$s.tableFilter.field, [string]$s.tableFilter.criteria) }
      }
      if ($s.hideRows) { $ws.Rows([string]$s.hideRows).Hidden = $true }
      $sel = $ws.Range($case.select)
      if ($sel.Cells.Count -lt 2) { throw "one cell selected in '$($case.name)': AutoSum would wait for Enter" }
      [void]$ws.Activate()
      [void]$sel.Select()
      try { [void]$x.CommandBars.ExecuteMso($ids[[string]$case.fn]) } catch { $answers.Add([ordered]@{ name = [string]$case.name; error = "$_" }); Save; continue }
      if (-not $x.Ready) { throw "Excel is editing a cell after '$($case.name)'" }
      $a = [ordered]@{ name = [string]$case.name; selection = [string]$x.Selection.Address($false, $false) }
      $cells = [ordered]@{}
      foreach ($c in $ws.Range($case.read).Cells) {
        $one = [ordered]@{}
        if ($c.HasFormula) { $one.f = [string]$c.Formula }
        $v = $c.Value2
        if ($null -ne $v) { $one.v = if ($v -is [double] -or $v -is [bool]) { $v } elseif ($v -is [int]) { [double]$v } else { "$v" } }
        if ($c.NumberFormat -ne 'General') { $one.fmt = [string]$c.NumberFormat }
        if ($one.Count) { $cells[$c.Address($false, $false)] = $one }
      }
      $a.cells = $cells
      $answers.Add($a)
      Save
    } finally { $wb.Close($false) }
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Save
"$($answers.Count) cases"
