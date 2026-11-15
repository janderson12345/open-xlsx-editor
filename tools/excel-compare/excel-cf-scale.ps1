# What Excel writes for a colour scale with thresholds and colours of its own (New Formatting Rule's Format Style
# 2-Color or 3-Color Scale; over COM, AddColorScale and its ColorScaleCriteria): each case a new workbook in an invisible
# Excel, 1 to 10 in A1:A10, saved to a folder outside the repository, its conditionalFormatting read back, and the
# fill Excel shows in A1, A5 and A10 (DisplayFormat). No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-cf-scale.ps1 <out.json> <folder for the saved workbooks>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
trap { Write-Host "LINE $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; break }
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
# Each: points, then [type (XlConditionValueTypes), value or $null, colour (Excel's B*65536+G*256+R)] per point.
$CASES = [ordered]@{
  '3-color: number 2, percent 50, highest' = @(3, @(@(0, 2, 255), @(3, 50, 65535), @(2, $null, 32768)))
  '2-color: percentile 10, percent 90' = @(2, @(@(5, 10, 16777215), @(3, 90, 12611584)))
  '2-color: lowest, number 6' = @(2, @(@(1, $null, 13434879), @(0, 6, 26112)))
}
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$results = @()
try {
  $k = 0
  foreach ($name in $CASES.Keys) {
    $k++
    $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
    for ($i = 1; $i -le 10; $i++) { $ws.Cells.Item($i, 1).Value2 = $i }
    $spec = $CASES[$name]
    $cs = $ws.Range('A1:A10').FormatConditions.AddColorScale($spec[0])
    for ($p = 0; $p -lt $spec[0]; $p++) {
      $c = $cs.ColorScaleCriteria.Item($p + 1); $pt = $spec[1][$p]
      $c.Type = [int]$pt[0]; if ($null -ne $pt[1]) { $c.Value = $pt[1] }; $c.FormatColor.Color = [int]$pt[2]
    }
    $shown = [ordered]@{}; foreach ($a in 'A1', 'A5', 'A10') { $shown[$a] = [int]$ws.Range($a).DisplayFormat.Interior.Color }
    $file = Join-Path $Books "cf-scale-$k.xlsx"
    if (Test-Path $file) { Remove-Item $file }
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    $s = New-Object IO.StreamReader($zip.GetEntry('xl/worksheets/sheet1.xml').Open()); $sheet = $s.ReadToEnd(); $s.Close()
    $zip.Dispose()
    $cf = ([regex]::Matches($sheet, '<conditionalFormatting\b[\s\S]*?</conditionalFormatting>') | ForEach-Object { $_.Value }) -join ''
    $results += [pscustomobject]@{ name = $name; conditionalFormatting = $cf; shown = $shown }
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x); [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
$results | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 $Out
"$($results.Count) cases"
