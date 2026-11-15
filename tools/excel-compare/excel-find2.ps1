# What Excel's Find and Replace does with Look in Notes and Comments, and with formats (Find Format, Replace Format):
# each case on a new workbook of made-up cells in an invisible Excel; the cells a search finds (Range.Find, then
# FindNext round to the first), or the cells after a Replace (value, bold, fill colour). Application.FindFormat and
# ReplaceFormat are cleared after each case. No keys sent, no setting saved.
# usage: pwsh tools/excel-compare/excel-find2.ps1 <out.json>
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
trap { Write-Host "LINE $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; break }
$m = [Type]::Missing
function Found($ws, $what, $lookIn, $byFormat) {
  $first = $ws.Cells.Find($what, $ws.Cells.Item(1, 1), $lookIn, 2, 1, 1, $false, $m, $byFormat)
  if (-not $first) { return @() }
  $out = @($first.Address($false, $false)); $at = $first
  for ($k = 0; $k -lt 20; $k++) { $at = $ws.Cells.FindNext($at); if (-not $at -or $at.Address() -eq $first.Address()) { break }; $out += $at.Address($false, $false) }
  return $out
}
function Notes($ws) {
  [void]$ws.Range('A1').AddComment('Check the apple stock'); [void]$ws.Range('B2').AddComment('nothing here')
  [void]$ws.Range('E5').AddComment('APPLES')
  [void]$ws.Range('C3').AddCommentThreaded('apple pie thread'); $ws.Range('D4').Value2 = 'apple'
}
function Formats($ws) {
  $ws.Range('A1').Value2 = 'apple'; $ws.Range('A1').Font.Bold = $true
  $ws.Range('A2').Value2 = 'apple'
  $ws.Range('A3').Value2 = 'pear'; $ws.Range('A3').Font.Bold = $true
  $ws.Range('B1').Value2 = 'apple tart'; $ws.Range('B1').Font.Bold = $true
}
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$results = @()
try {
  $cases = [ordered]@{
    'notes: apple' = { param($ws) Notes $ws; @{ found = @(Found $ws 'apple' -4144 $false) } }
    'comments: apple' = { param($ws) Notes $ws; @{ found = @(Found $ws 'apple' -4184 $false) } }
    'formulas: apple, with notes about' = { param($ws) Notes $ws; @{ found = @(Found $ws 'apple' -4123 $false) } }
    'format bold: anything' = { param($ws) Formats $ws; $x.FindFormat.Clear(); $x.FindFormat.Font.Bold = $true; @{ found = @(Found $ws '' -4123 $true) } }
    'format bold: apple' = { param($ws) Formats $ws; $x.FindFormat.Clear(); $x.FindFormat.Font.Bold = $true; @{ found = @(Found $ws 'apple' -4123 $true) } }
    'replace apple with fruit, bold only, filled yellow' = { param($ws) Formats $ws; $x.FindFormat.Clear(); $x.FindFormat.Font.Bold = $true; $x.ReplaceFormat.Clear(); $x.ReplaceFormat.Interior.Color = 65535
      [void]$ws.Cells.Replace('apple', 'fruit', 2, 1, $false, $true, $true); @{} }
    'replace formats only: bold filled yellow' = { param($ws) Formats $ws; $x.FindFormat.Clear(); $x.FindFormat.Font.Bold = $true; $x.ReplaceFormat.Clear(); $x.ReplaceFormat.Interior.Color = 65535
      [void]$ws.Cells.Replace('', '', 2, 1, $false, $true, $true); @{} }
  }
  foreach ($name in $cases.Keys) {
    $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
    Write-Host "case: $name"
    $r = & $cases[$name] $ws
    $cells = [ordered]@{}
    foreach ($c in $ws.UsedRange.Cells) { if ($null -ne $c.Value2) { $cells[$c.Address($false, $false)] = @([string]$c.Value2, [bool]$c.Font.Bold, [int]$c.Interior.ColorIndex) } }
    $results += [pscustomobject]@{ name = $name; found = $r.found; cells = $cells }
    $x.FindFormat.Clear(); $x.ReplaceFormat.Clear()
    $wb.Close($false)
  }
} finally { $x.FindFormat.Clear(); $x.ReplaceFormat.Clear(); $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x); [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
$results | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 $Out
"$($results.Count) cases"
