# What Excel writes after the Rules Manager's changes, for cfmanage.test.js: in an invisible Excel, a new workbook with
# 1 to 10 in A1:A10; four rules added in turn (greater than 5, a data bar, a 3 arrows icon set, a formula rule =MOD(A1,2)=0),
# then, a case at a time on a fresh copy of that: the second rule deleted, the third moved up, the first moved last,
# Stop If True set on the formula rule, the formula rule's range widened to A1:B10; each saved outside the repository
# and the sheet's conditionalFormatting and extLst read back. No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-cf-manage.ps1 <out.json> <folder for the saved workbooks>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$CASES = [ordered]@{
  'start' = { param($fc) }
  'deleteSecond' = { param($fc) $fc.Item(2).Delete() }
  'thirdUp' = { param($fc) $fc.Item(3).Priority = 2 }
  'firstLast' = { param($fc) $fc.Item(1).SetLastPriority() }
  'stopFormula' = { param($fc) $fc.Item(4).StopIfTrue = $true }
  'widen' = { param($fc) $fc.Item(4).ModifyAppliesToRange($fc.Parent.Worksheet.Range('A1:B10')) }
}
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$read = [ordered]@{}
try {
  foreach ($name in $CASES.Keys) {
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    for ($r = 1; $r -le 10; $r++) { $ws.Range("A$r").Formula = "$r" }
    $rg = $ws.Range('A1:A10')
    # (COM gives each rule added the last priority: 1 greater than 5, 2 the bar, 3 the arrows, 4 the formula.)
    $c1 = $rg.FormatConditions.Add(1, 5, '5'); $c1.Interior.Color = 0xCEC7FF; $c1.Font.Color = 0x06009C
    [void]$rg.FormatConditions.AddDatabar()
    $is = $rg.FormatConditions.AddIconSetCondition(); $is.IconSet = $wb.IconSets.Item(1)
    $c4 = $rg.FormatConditions.Add(2, 0, '=MOD(A1,2)=0'); $c4.Font.Bold = $true
    & $CASES[$name] $rg.FormatConditions
    $file = Join-Path $Books "cf-manage-$name.xlsx"
    if (Test-Path $file) { Remove-Item $file }
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    $s = New-Object IO.StreamReader($zip.GetEntry('xl/worksheets/sheet1.xml').Open()); $sheet = $s.ReadToEnd(); $s.Close()
    $zip.Dispose()
    $read[$name] = [pscustomobject]@{ cf = @([regex]::Matches($sheet, '<conditionalFormatting\b[\s\S]*?</conditionalFormatting>') | ForEach-Object { $_.Value }); ext = [regex]::Match($sheet, '<extLst>(?:(?!<extLst>)[\s\S])*</extLst>(?=</worksheet>)').Value }
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
[pscustomobject]@{ cases = $read } | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 $Out
'done'
