# What Excel writes for a conditional format's own format (Format... in New Formatting Rule; over COM the
# FormatCondition's NumberFormat, Font, Interior and Borders): each case a new workbook in an invisible Excel, 1 to 10
# in A1:A10, a rule "greater than 5" with that format, the workbook saved to a folder outside the repository and the
# styles' dxfs read back, and the shown text of A7 (7) and A2 (2). No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-cf-format.ps1 <out.json> <folder for the saved workbooks>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
trap { Write-Host "LINE $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; break }
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$CASES = [ordered]@{
  'number format 0.00%' = { param($fc) $fc.NumberFormat = '0.00%' }
  'bold, italic, red text, green fill' = { param($fc) $fc.Font.Bold = $true; $fc.Font.Italic = $true; $fc.Font.Color = 255; $fc.Interior.Color = 5296274 }
  'thin border all round' = { param($fc) foreach ($s in -4131, -4152, -4160, -4107) { $fc.Borders.Item($s).LineStyle = 1 } }
  'red bottom border' = { param($fc) $b = $fc.Borders.Item(-4107); $b.LineStyle = 1; $b.Color = 255 }
  'dashed top border' = { param($fc) $fc.Borders.Item(-4160).LineStyle = -4115 }
  'all of it' = { param($fc) $fc.NumberFormat = '#,##0.0'; $fc.Font.Bold = $true; $fc.Interior.Color = 65535; foreach ($s in -4131, -4152, -4160, -4107) { $fc.Borders.Item($s).LineStyle = 1 } }
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
    $fc = $ws.Range('A1:A10').FormatConditions.Add(1, 5, '5')
    & $CASES[$name] $fc
    $shown = [ordered]@{ A7 = [string]$ws.Range('A7').Text; A2 = [string]$ws.Range('A2').Text }
    $file = Join-Path $Books "cf-format-$k.xlsx"
    if (Test-Path $file) { Remove-Item $file }
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    $s = New-Object IO.StreamReader($zip.GetEntry('xl/styles.xml').Open()); $styles = $s.ReadToEnd(); $s.Close()
    $zip.Dispose()
    $dxfs = ([regex]::Match($styles, '<dxfs\b[\s\S]*?</dxfs>|<dxfs\b[^>]*/>')).Value
    $results += [pscustomobject]@{ name = $name; dxfs = $dxfs; shown = $shown }
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x); [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
$results | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 $Out
"$($results.Count) cases"
