# Has Excel read typed entries (entry-cases.json: { text, before? }, `before` a number format the cell has first): each
# is put in a cell of a new workbook as typing would (Range.Formula = text, which Excel reads in the computer's
# settings: United States here), and its value, type (an error value as Excel shows it) and number format are written to JSON with today's date (a date
# typed without a year takes the current one). An invisible instance, macros disabled for it only, closed without
# saving; no Excel setting is changed and no keys are sent.
# usage: pwsh tools/excel-compare/excel-entries.ps1 tools/excel-compare/entry-cases.json test/entry-excel.json
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
# (Not $cases or $out: PowerShell's names ignore case, and those are the paths.)
$list = Get-Content $Cases -Raw | ConvertFrom-Json
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3 # msoAutomationSecurityForceDisable: this instance runs no macros
$answers = New-Object System.Collections.Generic.List[object]
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $ws.Columns.Item(1).ColumnWidth = 40 # (wide enough that an error shows its name, not ####)
  $i = 0
  foreach ($e in $list) {
    $i++
    $c = $ws.Cells.Item($i, 1)
    if ($e.before) { $c.NumberFormat = [string]$e.before }
    $c.Formula = [string]$e.text
    $v = $c.Value2
    # (An error value comes back as a number of COM's, not a double: written as Excel shows it.)
    $type = if ($null -eq $v) { 'blank' } elseif ($v -is [double]) { 'number' } elseif ($v -is [bool]) { 'boolean' } elseif ($v -is [int]) { 'error' } else { 'text' }
    $value = if ($v -is [double]) { $v } elseif ($v -is [bool]) { $v } elseif ($v -is [int]) { [string]$c.Text } else { "$v" }
    $a = [ordered]@{ text = [string]$e.text }
    if ($e.before) { $a.before = [string]$e.before }
    $a.type = $type; $a.value = $value; $a.format = [string]$c.NumberFormat
    $answers.Add($a)
  }
  $wb.Close($false)
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
[ordered]@{ excel = 'Excel.Application, Range.Formula = text (United States settings)'; today = (Get-Date).ToString('yyyy-MM-dd'); entries = $answers } | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 $Out
"$($answers.Count) entries"
