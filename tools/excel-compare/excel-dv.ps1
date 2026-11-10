# Has Excel judge values against data validation rules (cases: [{ name, setup: { cells, formats }, rule: { range, type:
# "whole" | "decimal" | "list" | "date" | "time" | "length" | "custom", op, f1, f2, ignoreBlank, alert }, read }]):
# for each cell of `read`, whether Excel takes its value as valid (Range.Validation.Value), and, from the workbook Excel
# then saves, its dataValidations element, which the tests put into a sheet of their own. An invisible instance,
# macros disabled for it only; each workbook saved to a temporary folder, read, and deleted.
# usage: pwsh tools/excel-compare/excel-dv.ps1 tools/excel-compare/dv-cases.json test/dv-excel.json
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$list = Get-Content $Cases -Raw | ConvertFrom-Json
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('oxe-dv-' + [guid]::NewGuid().ToString('n').Substring(0, 8))
New-Item -ItemType Directory $tmp | Out-Null
$types = @{ whole = 1; decimal = 2; list = 3; date = 4; time = 5; length = 6; custom = 7 }
$ops = @{ between = 1; notBetween = 2; equal = 3; notEqual = 4; greater = 5; less = 6; greaterEqual = 7; lessEqual = 8 }
$alerts = @{ stop = 1; warning = 2; information = 3 }
$m = [Type]::Missing
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3 # msoAutomationSecurityForceDisable: this instance runs no macros
$answers = New-Object System.Collections.Generic.List[object]
function Props($o) { if ($null -eq $o) { @() } else { $o.PSObject.Properties } }
function PartText($zip, $name) { $e = $zip.GetEntry($name); if (-not $e) { return '' }; $r = New-Object IO.StreamReader($e.Open()); try { $r.ReadToEnd() } finally { $r.Dispose() } }
try {
  $i = 0
  foreach ($case in $list) {
    $i++
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    foreach ($p in Props $case.setup.cells) { $ws.Range($p.Name).Formula = [string]$p.Value }
    foreach ($f in @($case.setup.formats)) { if ($f -and $f.numFmt) { $ws.Range($f.range).NumberFormat = [string]$f.numFmt } }
    $rule = $case.rule
    $v = $ws.Range($rule.range).Validation
    $v.Delete()
    $op = if ($rule.op) { $ops[$rule.op] } else { $m }
    $f2 = if ($rule.f2) { [string]$rule.f2 } else { $m }
    $alert = if ($rule.alert) { $alerts[$rule.alert] } else { 1 }
    [void]$v.Add($types[$rule.type], $alert, $op, [string]$rule.f1, $f2)
    $v.IgnoreBlank = $(if ($null -eq $rule.ignoreBlank) { $true } else { [bool]$rule.ignoreBlank })
    $a = [ordered]@{ name = [string]$case.name }
    $valid = [ordered]@{}
    foreach ($cell in $ws.Range($case.read).Cells) { $valid[$cell.Address($false, $false)] = [bool]$cell.Validation.Value }
    $a.valid = $valid
    $file = Join-Path $tmp "$i.xlsx"
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    try {
      $dv = [regex]::Match((PartText $zip 'xl/worksheets/sheet1.xml'), '<dataValidations.*?</dataValidations>')
      $a.dataValidations = if ($dv.Success) { $dv.Value } else { $null }
    } finally { $zip.Dispose() }
    $answers.Add($a)
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
}
[ordered]@{ excel = 'Excel.Application: Validation.Add, Range.Validation.Value, then the saved sheet XML'; cases = $answers } | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 $Out
"$($answers.Count) cases"
