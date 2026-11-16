# Has Excel work out a list of formulas and writes what it got to JSON, for finding Excel's rules: each formula is
# written into a cell of a new workbook (or of -Book, opened read-only), worked out, its Value2 read, and the workbook
# closed without saving. Invisible, macros disabled for this instance only, no Excel setting changed, no keys sent.
# Text results carry their UTF-16 units too (`u`), so half of a surrogate pair survives the JSON.
# usage: pwsh tools/excel-compare/excel-formulas.ps1 <formulas.json> <out.json> [-Book <book.xlsx>]
#   formulas.json: an array of formula strings, each starting with =.
param([Parameter(Mandatory)][string]$In, [Parameter(Mandatory)][string]$Out, [string]$Book)
$ErrorActionPreference = 'Stop'
$formulas = Get-Content -Raw -Encoding utf8 $In | ConvertFrom-Json
$errors = @{ -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!'; -2146826265 = '#REF!';
  -2146826259 = '#NAME?'; -2146826252 = '#NUM!'; -2146826246 = '#N/A'; -2146826245 = '#GETTING_DATA';
  -2146826238 = '#SPILL!'; -2146826236 = '#CALC!'; -2146826237 = '#FIELD!'; -2146826239 = '#BLOCKED!';
  -2146826240 = '#UNKNOWN!'; -2146826241 = '#BUSY!'; -2146826242 = '#CONNECT!'; -2146826243 = '#PYTHON!' }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$x.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
$results = New-Object System.Collections.Generic.List[object]
try {
  if ($Book) { $wb = $x.Workbooks.Open((Resolve-Path $Book).Path, 0, $true) } else { $wb = $x.Workbooks.Add() }
  # (A sheet of its own, so nothing of the book's is in the way; it goes when the book closes unsaved.)
  $ws = $wb.Worksheets.Add()
  $i = 0
  foreach ($f in $formulas) {
    $i++
    $c = $ws.Cells.Item($i, 1)
    $entry = [ordered]@{ f = $f }
    try { $c.Formula2 = $f } catch { $entry.t = 'refused'; $results.Add($entry); continue }
    $v = $c.Value2
    if ($v -is [object[,]]) { $v = $v[1, 1] }
    if ($v -is [double]) { $entry.t = 'n'; $entry.v = $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) }
    elseif ($v -is [bool]) { $entry.t = 'b'; $entry.v = $v }
    elseif ($v -is [int]) { $entry.t = 'e'; $entry.v = $(if ($errors.ContainsKey($v)) { $errors[$v] } else { "err$v" }) }
    elseif ($null -eq $v) { $entry.t = 'blank' }
    else { $s = [string]$v; $entry.t = 's'; $entry.u = @($s.ToCharArray() | ForEach-Object { [int]$_ }) }
    $results.Add($entry)
  }
  $wb.Close($false)
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
$json = ConvertTo-Json -InputObject $results.ToArray() -Depth 4 -Compress
[IO.File]::WriteAllText($Out, $json, (New-Object Text.UTF8Encoding $false))
"$($results.Count) formulas"
