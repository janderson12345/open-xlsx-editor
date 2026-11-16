# Has Excel make links to another workbook and saves what it writes: Paste Link (Range.Copy in a source workbook,
# Worksheet.Paste with Link in another), defined names that refer to the source, a formula using the source's own
# defined name, and Workbook.LinkSources. Two new workbooks of made-up numbers in an invisible instance (macros
# disabled for it only); nothing else is changed and no keys are sent. Range.Copy uses the Windows clipboard: its
# text is put back afterwards.
# Writes $Out\src.xlsx, $Out\dst.xlsx (saved with the source open, then again once it was closed) and
# $Out\extlinks.json: each linked cell's Formula, Value and format with the source open and closed, the names'
# RefersTo, and LinkSources.
# usage: pwsh tools/excel-compare/excel-extlinks.ps1 <folder outside the repo>
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
$keep = [System.Windows.Forms.Clipboard]::GetText()
New-Item -ItemType Directory -Force $Out | Out-Null
# (Excel's SaveAs refuses long paths: short ones in %TEMP%, moved afterwards.)
$src = Join-Path $env:TEMP 'oxesrc.xlsx'
$dst = Join-Path $env:TEMP 'oxedst.xlsx'
Remove-Item $src, $dst -ErrorAction SilentlyContinue
$missing = [Type]::Missing
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$res = [ordered]@{}
try {
  $s = $x.Workbooks.Add()
  $ss = $s.Worksheets.Item(1)
  $ss.Range('A1').Value2 = 5
  $ss.Range('B1').Value2 = 2.5
  $ss.Range('A2').Value2 = 'text'
  $ss.Range('B2').Formula = '=A1*2'
  $ss.Range('C1').Value2 = $true
  $ss.Range('C2').Formula = '=1/0'
  $two = $s.Worksheets.Add([Type]::Missing, $ss)
  $two.Name = 'Data Two'
  $two.Range('A1').Value2 = 7
  [void]$s.Names.Add('SrcRate', '=Sheet1!$B$1')
  $s.SaveAs($src, 51)
  $d = $x.Workbooks.Add()
  $ds = $d.Worksheets.Item(1)
  # A block, one cell, another sheet's cell, a block with a blank.
  $links = @(@('Sheet1', 'A1:C2', 'D4'), @('Sheet1', 'B1', 'A1'), @('Data Two', 'A1', 'A10'), @('Sheet1', 'A1:A3', 'H1'))
  foreach ($l in $links) {
    [void]$s.Worksheets.Item($l[0]).Range($l[1]).Copy()
    [void]$d.Activate(); [void]$ds.Activate(); [void]$ds.Range($l[2]).Select()
    [void]$ds.Paste($missing, $true)
    $x.CutCopyMode = 0
  }
  [void]$d.Names.Add('ExtRate', '=[oxesrc.xlsx]Sheet1!$A$1')
  [void]$d.Names.Add('ExtBlock', '=[oxesrc.xlsx]Sheet1!$A$1:$B$2')
  [void]$ds.Names.Add('LocalExt', "='[oxesrc.xlsx]Data Two'!`$A`$1")
  [void]$d.Names.Add('ExtConst', '=[oxesrc.xlsx]Sheet1!$A$1*2')
  $ds.Range('A12').Formula = '=SUM(ExtBlock)'
  $ds.Range('A13').Formula = '=oxesrc.xlsx!SrcRate'
  $ds.Range('A14').Formula = '=[oxesrc.xlsx]Sheet1!$A$1+1'
  $read = { param($label)
    $cells = [ordered]@{}
    foreach ($a in @('D4', 'E4', 'F4', 'D5', 'E5', 'F5', 'A1', 'A10', 'H1', 'H2', 'H3', 'A12', 'A13', 'A14')) {
      $c = $ds.Range($a)
      $cells[$a] = [ordered]@{ formula = [string]$c.Formula; value = $c.Value2; text = [string]$c.Text; numFmt = [string]$c.NumberFormat }
    }
    $names = [ordered]@{}
    foreach ($n in $d.Names) { $names[$n.Name] = [string]$n.RefersTo }
    $res[$label] = [ordered]@{ cells = $cells; names = $names; links = @($d.LinkSources(1)) }
  }
  & $read 'open'
  $d.SaveAs($dst, 51)
  $s.Close($false)
  & $read 'closed'
  $d.Save()
  $d.Close($false)
} finally {
  $x.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  if ($keep) { [System.Windows.Forms.Clipboard]::SetText($keep) }
}
Move-Item -Force $src (Join-Path $Out 'src.xlsx')
Move-Item -Force $dst (Join-Path $Out 'dst.xlsx')
$res | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 (Join-Path $Out 'extlinks.json')
