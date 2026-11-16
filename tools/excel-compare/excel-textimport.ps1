# Has Excel's Text Import Wizard open made-up text files (cases.json, from textimport-cases.mjs) with the
# settings each case gives (Workbooks.OpenText: code page, start row, delimited or fixed width, delimiters, text
# qualifier, consecutive delimiters as one, each column's type, the decimal and thousands separators), and writes each
# cell's value, type, number format and formula to JSON. An invisible instance, macros disabled for it only, every
# workbook closed without saving; no Excel setting is changed and no keys are sent.
# usage: pwsh tools/excel-compare/excel-textimport.ps1 ..\probe\textimport\cases.json test/textimport-excel.json
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$list = Get-Content $Cases -Raw | ConvertFrom-Json
$types = @{ general = 1; text = 2; MDY = 3; DMY = 4; YMD = 5; skip = 9 }
# (UTF-16: code page 1200 is refused over COM; xlWindows, and Excel reads the byte order mark.)
$origins = @{ 'utf-8' = 65001; 'utf-16le' = 2; 'windows-1252' = 1252 }
$quotes = @{ '"' = 1; "'" = 2; '' = -4142 }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
$answers = New-Object System.Collections.Generic.List[object]
try {
  foreach ($c in $list) {
    $seps = [string]$c.seps
    $fixed = $c.kind -eq 'fixed'
    $info = New-Object System.Collections.Generic.List[object]
    for ($k = 0; $k -lt $c.columns.Count; $k++) {
      $at = if ($fixed) { if ($k -eq 0) { 0 } else { [int]$c.breaks[$k - 1] } } else { $k + 1 }
      $info.Add([object[]]@($at, $types[[string]$c.columns[$k]]))
    }
    $known = ";,`t "
    $other = ($seps.ToCharArray() | Where-Object { $known.IndexOf($_) -lt 0 } | Select-Object -First 1)
    $dec = if ($c.decimal) { [string]$c.decimal } else { '.' }
    $grp = if ($c.group) { [string]$c.group } else { ',' }
    $start = if ($c.startRow) { [int]$c.startRow } else { 1 }
    $args = @([string]$c.file, [int]$origins[[string]$c.encoding], [int]$start, [int]$(if ($fixed) { 2 } else { 1 }), [int]$quotes[[string]$c.quote],
      [bool]$c.consecutive, $seps.Contains("`t"), $seps.Contains(';'), $seps.Contains(','), $seps.Contains(' '), [bool]$other,
      [string]$(if ($other) { $other } else { '~' }), $info.ToArray(), 1, $dec, $grp)
    try { $x.Workbooks.OpenText.Invoke($args) } catch { throw "$($c.name): $($_.Exception.Message) ($($args[1..11] -join ' '))" }
    $wb = $x.ActiveWorkbook
    $ws = $wb.Worksheets.Item(1)
    $used = $ws.UsedRange
    $cells = New-Object System.Collections.Generic.List[object]
    for ($r = 1; $r -le $used.Rows.Count; $r++) {
      for ($k = 1; $k -le $used.Columns.Count; $k++) {
        $cell = $used.Cells.Item($r, $k)
        $v = $cell.Value2
        if ($null -eq $v) { continue }
        $type = if ($v -is [double]) { 'number' } elseif ($v -is [bool]) { 'boolean' } elseif ($v -is [int]) { 'error' } else { 'text' }
        $value = if ($v -is [double] -or $v -is [bool]) { $v } elseif ($v -is [int]) { [string]$cell.Text } else { "$v" }
        $a = [ordered]@{ ref = $cell.Address($false, $false); type = $type; value = $value; format = [string]$cell.NumberFormat }
        if ($cell.HasFormula) { $a.formula = [string]$cell.Formula }
        $cells.Add($a)
      }
    }
    $answers.Add([ordered]@{ name = [string]$c.name; cells = $cells })
    $wb.Close($false)
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
$json = [ordered]@{ excel = 'Excel.Application, Workbooks.OpenText (United States settings)'; today = (Get-Date).ToString('yyyy-MM-dd'); cases = $answers } | ConvertTo-Json -Depth 6
# (UTF-8 without a byte order mark: the repo's byte scan refuses one.)
[IO.File]::WriteAllText((Join-Path (Get-Location) $Out), $json, (New-Object Text.UTF8Encoding $false))
"$($answers.Count) cases"
