# What Excel writes when a conditional formatting rule is edited (the Rules Manager's Edit Rule; over COM,
# FormatCondition.Modify, which changes a rule's type and criteria in place): for each case a new workbook in an
# invisible Excel, numbers 1 to 10 in A1:A10, two or three rules added, one of them changed, the workbook saved to a
# folder outside the repository and its conditionalFormatting and dxfs read back (no user name in either). No keys
# sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-cf-edit.ps1 <out.json> <folder for the saved workbooks>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
trap { Write-Host "LINE $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; break }
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$m = [Type]::Missing
# Each case: rules added to A1:A10 (in order), then the change. Rules: @(type, operator, f1, f2, fill colour).
$CASES = [ordered]@{
  'greater than 5 to between 3 and 7' = @(@(@(1, 5, '5', $m, 13551615), @(2, $m, '=A1=2', $m, 10284031)), { param($fcs) $fcs.Item(1).Modify(1, 1, '3', '7') })
  'greater than 5 to a formula' = @(@(@(1, 5, '5', $m, 13551615), @(2, $m, '=A1=2', $m, 10284031)), { param($fcs) $fcs.Item(1).Modify(2, $m, '=MOD(A1,2)=0') })
  'formula to less than 4' = @(@(@(1, 5, '5', $m, 13551615), @(2, $m, '=A1=2', $m, 10284031)), { param($fcs) $fcs.Item(2).Modify(1, 6, '4') })
  'the middle of three, stop if true off' = @(@(@(1, 5, '8', $m, 13551615), @(1, 6, '3', $m, 10284031), @(2, $m, '=A1=5', $m, 13561798)), { param($fcs) $fcs.Item(2).StopIfTrue = $false; $fcs.Item(2).Modify(1, 3, '6') })
}
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$results = @()
try {
  $k = 0
  foreach ($name in $CASES.Keys) {
    $k++
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    for ($i = 1; $i -le 10; $i++) { $ws.Cells.Item($i, 1).Value2 = $i }
    $r = $ws.Range('A1:A10')
    foreach ($rule in $CASES[$name][0]) {
      $fc = if ($rule[0] -eq 1) { $r.FormatConditions.Add(1, $rule[1], $rule[2], $rule[3]) } else { $r.FormatConditions.Add(2, $m, $rule[2]) }
      $fc.Interior.Color = $rule[4]
    }
    & $CASES[$name][1] $r.FormatConditions
    $file = Join-Path $Books "cf-edit-$k.xlsx"
    if (Test-Path $file) { Remove-Item $file }
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    $read = { param($n) $s = New-Object IO.StreamReader($zip.GetEntry($n).Open()); $t = $s.ReadToEnd(); $s.Close(); $t }
    $sheet = & $read 'xl/worksheets/sheet1.xml'
    $styles = & $read 'xl/styles.xml'
    $zip.Dispose()
    $cf = ([regex]::Matches($sheet, '<conditionalFormatting\b[\s\S]*?</conditionalFormatting>') | ForEach-Object { $_.Value }) -join ''
    $dxfs = ([regex]::Match($styles, '<dxfs\b[\s\S]*?</dxfs>|<dxfs\b[^>]*/>')).Value
    $results += [pscustomobject]@{ name = $name; conditionalFormatting = $cf; dxfs = $dxfs }
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x); [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
$results | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 $Out
"$($results.Count) cases"
