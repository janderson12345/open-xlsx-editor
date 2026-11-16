# Has Excel save a workbook's sheet as text files: Save As Text (Tab delimited) (xlText, -4158), Unicode Text (42),
# CSV UTF-8 (62) and CSV (6), each a copy of the sheet named by -Sheet. An invisible instance, macros disabled for it
# only, the workbook opened read-only and closed without saving; no Excel setting is changed and no keys are sent.
# usage: pwsh tools/excel-compare/excel-saveas.ps1 ..\probe\save\save.xlsx Save ..\probe\save
param([Parameter(Mandatory)][string]$Book, [Parameter(Mandatory)][string]$Sheet, [Parameter(Mandatory)][string]$OutDir)
$ErrorActionPreference = 'Stop'
$book = (Resolve-Path $Book).Path
$dir = (Resolve-Path $OutDir).Path
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
try {
  foreach ($f in @(@{ n = 'text.txt'; t = -4158 }, @{ n = 'unicode.txt'; t = 42 }, @{ n = 'utf8.csv'; t = 62 }, @{ n = 'ansi.csv'; t = 6 })) {
    $wb = $x.Workbooks.Open($book, 0, $true)
    $wb.Worksheets.Item($Sheet).Activate()
    $wb.SaveAs((Join-Path $dir $f.n), $f.t)
    $wb.Close($false)
    "$($f.n)"
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
