# Has Excel open workbooks saved by the editor and say what is in them, as excel-steps.ps1 says what came of its own
# steps (a list: [{ name, file, sheet, sheet2, read, cells2 }], as test/shift.test.js writes it with SHIFT_SAVE): the
# cells of `read` on `sheet` (formula, value, number format, bold, fill colour, bottom border), the formulas of cells2
# on `sheet2`, and the merged ranges. An invisible instance, macros disabled for it only, each file closed without
# saving; no Excel setting is changed and no keys are sent.
# usage: pwsh tools/excel-compare/excel-readback.ps1 <folder>/saved.json <out.json>
param([Parameter(Mandatory)][string]$List, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$items = Get-Content $List -Raw | ConvertFrom-Json
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3 # msoAutomationSecurityForceDisable: this instance runs no macros
$answers = New-Object System.Collections.Generic.List[object]
function CellOf($c) {
  $a = [ordered]@{}
  if ($c.HasFormula) { $a.f = [string]$c.Formula }
  $v = $c.Value2
  if ($null -ne $v) { $a.v = if ($v -is [double] -or $v -is [bool]) { $v } elseif ($v -is [int]) { [double]$v } else { "$v" } }
  if ($c.NumberFormat -ne 'General') { $a.fmt = [string]$c.NumberFormat }
  if ($c.Font.Bold) { $a.bold = $true }
  if ($c.Interior.ColorIndex -ne -4142) { $a.color = [int]$c.Interior.Color }
  if ($c.Borders.Item(9).LineStyle -ne -4142) { $a.border = $true }
  $a
}
try {
  foreach ($it in $items) {
    $wb = $x.Workbooks.Open([string]$it.file, 0, $true)
    $ws = $wb.Worksheets.Item([string]$it.sheet)
    $a = [ordered]@{ name = [string]$it.name }
    $cells = [ordered]@{}
    foreach ($c in $ws.Range($it.read).Cells) {
      $one = CellOf $c
      if ($one.Count) { $cells[$c.Address($false, $false)] = $one }
    }
    $a.cells = $cells
    $two = [ordered]@{}
    foreach ($ref in @($it.cells2)) { if ($ref) { $two[$ref] = [string]$wb.Worksheets.Item([string]$it.sheet2).Range($ref).Formula } }
    if ($two.Count) { $a.sheet2 = $two }
    $merged = New-Object System.Collections.Generic.List[string]
    foreach ($c in $ws.Range($it.read).Cells) { if ($c.MergeCells) { $ad = $c.MergeArea.Address($false, $false); if (-not $merged.Contains($ad)) { $merged.Add($ad) } } }
    if ($merged.Count) { $a.merges = $merged }
    $answers.Add($a)
    $wb.Close($false)
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
[ordered]@{ excel = 'Excel.Application: Workbooks.Open, read only'; cases = $answers } | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 $Out
"$($answers.Count) workbooks"
