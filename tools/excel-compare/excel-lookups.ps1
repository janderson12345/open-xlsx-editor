# Has Excel answer approximate lookups over unsorted lists (lookup-cases.mjs): MATCH(x,list,1), MATCH(x,list,-1) and
# VLOOKUP(x,list,1,TRUE) for each, written as test/lookup-excel.json. A new workbook in an invisible instance, macros
# disabled for this instance only, closed without saving; no Excel setting is changed and no keys are sent.
# usage: pwsh tools/excel-compare/excel-lookups.ps1 <cases.json> <out.json>
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$all = Get-Content $Cases -Raw | ConvertFrom-Json # (not $cases: PowerShell names ignore case, and $Cases is the path)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$x.AutomationSecurity = 3 # msoAutomationSecurityForceDisable: this instance runs no macros
$answers = New-Object System.Collections.Generic.List[object] # (not $out: that is $Out, the path)
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $x.Calculation = -4135 # manual while the cells are written
  $k = 0
  foreach ($c in $all) {
    $col = $k + 1; $i = 0
    # (Written through Formula as text: PowerShell's COM binding fails when one call site takes a number then a string.)
    foreach ($v in $c.list) { $i++; $cell = $ws.Cells.Item($i, $col); if ($v -is [string]) { $cell.Formula = [string]$v } elseif ($null -ne $v) { $cell.Formula = [string][double]$v } }
    $n = $c.list.Count
    $addr = $ws.Range($ws.Cells.Item(1, $col), $ws.Cells.Item($n, $col)).Address($false, $false)
    $ws.Cells.Item(20, $col).Formula = "=MATCH($($c.x),$addr,1)"
    $ws.Cells.Item(21, $col).Formula = "=MATCH($($c.x),$addr,-1)"
    $ws.Cells.Item(22, $col).Formula = "=VLOOKUP($($c.x),$addr,1,TRUE)"
    $k++
  }
  $x.Calculation = -4105; $x.CalculateFull()
  $k = 0
  foreach ($c in $all) {
    $col = $k + 1
    # (Value2 gives an error as an Int32: here only #N/A.)
    $get = { param($r) $v = $ws.Cells.Item($r, $col).Value2; if ($v -is [int]) { '#N/A' } else { $v } }
    $answers.Add([ordered]@{ list = @($c.list); x = $c.x; match1 = & $get 20; matchDown = & $get 21; vlookup = & $get 22 })
    $k++
  }
  $wb.Close($false)
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
[ordered]@{ excel = 'Excel.Application CalculateFull, MATCH(x,list,1), MATCH(x,list,-1), VLOOKUP(x,list,1,TRUE)'; cases = $answers } | ConvertTo-Json -Depth 4 -Compress | Set-Content -Encoding utf8 $Out
"$($answers.Count) cases"
