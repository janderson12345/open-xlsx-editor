# Excel opens rules.xlsx (dv-ours.mjs) and reads each listed cell's validation: Type, Operator, Formula1, AlertStyle,
# and its titles and messages (or none), and InCellDropdown for the list. Invisible, read-only, closed without saving.
# usage: pwsh tools/excel-compare/excel-dv-check.ps1 <folder> <out.json>
param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$results = @{}
try {
  $wb = $x.Workbooks.Open((Join-Path (Resolve-Path $Folder).Path 'rules.xlsx'), 0, $true)
  $ws = $wb.Worksheets.Item(1)
  foreach ($ref in 'C1', 'C3', 'C4', 'C5', 'C6', 'C10', 'D2', 'E2', 'F2', 'C11') {
    $v = $ws.Range($ref).Validation
    try {
      $t = $v.Type
      $op = $null; try { $op = $v.Operator } catch { }
      $msg = if ($v.InputTitle -or $v.ErrorTitle) { "$($v.InputTitle)|$($v.InputMessage)|$($v.ErrorTitle)|$($v.ErrorMessage)" } else { '' }
      $results[$ref] = @($t, $(if ($t -eq 3) { $null } else { $op }), $v.Formula1, $v.AlertStyle, $msg, $v.InCellDropdown)
    } catch { $results[$ref] = $null }
  }
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$results | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 $Out
'read'
