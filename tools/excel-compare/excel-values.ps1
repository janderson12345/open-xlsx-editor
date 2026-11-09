# Has Excel work out the formulas of a workbook and writes what it got to JSON, for the engine's tests: each formula
# cell's sheet, address, formula and value. Opened read-only and invisibly, links not updated, macros disabled for this
# instance only, every formula worked out (CalculateFull), closed without saving. No Excel setting is changed and no
# keys are sent.
# usage: pwsh tools/excel-compare/excel-values.ps1 <book.xlsx> <out.json>
param([Parameter(Mandatory)][string]$Book, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$Book = (Resolve-Path $Book).Path
# Value2 gives an error as an Int32 (CVErr): the codes Excel uses.
$errors = @{ -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!'; -2146826265 = '#REF!';
  -2146826259 = '#NAME?'; -2146826252 = '#NUM!'; -2146826246 = '#N/A' }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$x.AutomationSecurity = 3 # msoAutomationSecurityForceDisable: this instance runs no macros
$cells = New-Object System.Collections.Generic.List[object]
try {
  $wb = $x.Workbooks.Open($Book, 0, $true)
  $x.CalculateFull()
  foreach ($ws in $wb.Worksheets) {
    foreach ($c in $ws.UsedRange.Cells) {
      if (-not $c.HasFormula) { continue }
      $v = $c.Value2
      $entry = [ordered]@{ sheet = $ws.Name; ref = $c.Address($false, $false); formula = $c.Formula }
      if ($v -is [double]) { $entry.t = 'n'; $entry.v = $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) }
      elseif ($v -is [bool]) { $entry.t = 'b'; $entry.v = $v }
      elseif ($v -is [int]) { $entry.t = 'e'; $entry.v = $(if ($errors.ContainsKey($v)) { $errors[$v] } else { $c.Text }) }
      elseif ($null -eq $v) { $entry.t = 'blank'; $entry.v = $null }
      else { $entry.t = 's'; $entry.v = [string]$v }
      $cells.Add($entry)
    }
  }
  $wb.Close($false)
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
[ordered]@{ book = (Split-Path $Book -Leaf); excel = 'Excel.Application CalculateFull'; cells = $cells } | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 $Out
"$($cells.Count) formula cells"
