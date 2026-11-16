# Has Excel break one of two links to other workbooks (Workbook.BreakLink, as Edit Links' Break Link) and says what
# came of the formulas and defined names that used it, and of the other link's number. Three new workbooks of
# made-up numbers in an invisible instance (macros disabled for it only), saved to %TEMP% and moved to $Out:
# brka.xlsx and brkb.xlsx (the sources), brkdst.xlsx (before) and brkafter.xlsx (after the break); brk.json
# holds the formulas, values and names before and after.
# usage: pwsh tools/excel-compare/excel-breaklink.ps1 <folder outside the repo>
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $Out | Out-Null
$t = $env:TEMP
$paths = @{ a = (Join-Path $t 'brka.xlsx'); b = (Join-Path $t 'brkb.xlsx'); d = (Join-Path $t 'brkdst.xlsx'); e = (Join-Path $t 'brkafter.xlsx') }
foreach ($p in $paths.Values) { Remove-Item $p -ErrorAction SilentlyContinue }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$res = [ordered]@{}
try {
  $a = $x.Workbooks.Add(); $a.Worksheets.Item(1).Range('A1').Value2 = 5; $a.Worksheets.Item(1).Range('A2').Value2 = 'kept'; $a.SaveAs($paths.a, 51)
  $b = $x.Workbooks.Add(); $b.Worksheets.Item(1).Range('A1').Value2 = 9; $b.SaveAs($paths.b, 51)
  $d = $x.Workbooks.Add()
  $s = $d.Worksheets.Item(1)
  $s.Range('A1').Formula = '=[brka.xlsx]Sheet1!$A$1*2'
  $s.Range('A2').Formula = '=[brkb.xlsx]Sheet1!$A$1'
  $s.Range('A3').Formula = '=SUM([brka.xlsx]Sheet1!$A$1,[brkb.xlsx]Sheet1!$A$1)'
  $s.Range('A4').Formula = '=Na+1'
  $s.Range('A5').Formula = '=[brka.xlsx]Sheet1!$A$2'
  $s.Range('A6').Formula = '=Nb'
  [void]$d.Names.Add('Na', '=[brka.xlsx]Sheet1!$A$1')
  [void]$d.Names.Add('Nb', '=[brkb.xlsx]Sheet1!$A$1')
  [void]$d.Names.Add('Nc', '=[brka.xlsx]Sheet1!$A$1+[brkb.xlsx]Sheet1!$A$1')
  $read = { param($label)
    $cells = [ordered]@{}
    foreach ($r in 1..6) { $c = $s.Range("A$r"); $cells["A$r"] = [ordered]@{ formula = [string]$c.Formula; value = $c.Value2 } }
    $names = [ordered]@{}
    foreach ($n in $d.Names) { $names[$n.Name] = [string]$n.RefersTo }
    $res[$label] = [ordered]@{ cells = $cells; names = $names; links = @($d.LinkSources(1)) }
  }
  $d.SaveAs($paths.d, 51)
  $a.Close($false); $b.Close($false)
  & $read 'before'
  $d.BreakLink($paths.a, 1)
  & $read 'after'
  $d.SaveAs($paths.e, 51)
  $d.Close($false)
} finally {
  $x.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
foreach ($k in 'a', 'b', 'd', 'e') { Move-Item -Force $paths[$k] (Join-Path $Out (Split-Path $paths[$k] -Leaf)) }
$res | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 (Join-Path $Out 'brk.json')
