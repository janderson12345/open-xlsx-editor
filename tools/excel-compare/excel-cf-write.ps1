# What Excel writes for the rules its Conditional Formatting menu adds (Home > Conditional Formatting: Highlight Cells
# Rules, Top/Bottom Rules, Color Scales), each with the menu's preset format, for cf-write-cases.mjs: for each case a
# new workbook in an invisible Excel, numbers 1 to 10 and text in A1:B10, the rules added as the menu adds them, the
# workbook saved to a folder outside the repository and the sheet's conditionalFormatting elements and the styles'
# dxfs read from it (no user name in either). No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-cf-write.ps1 <cases.json> <out.json> <folder for the saved workbooks>
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
trap { Write-Host "LINE $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; break }
Add-Type -AssemblyName System.IO.Compression.FileSystem
$spec = Get-Content -Raw $Cases | ConvertFrom-Json
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$m = [Type]::Missing
# The menu's presets: [fill, font] as Excel colours (B * 65536 + G * 256 + R), $null for none.
$PRESETS = @{
  lightRed = @(13551615, 393372); yellow = @(10284031, 22428); green = @(13561798, 24832); redFill = @(13551615, $null); redText = @($null, 393372)
}
$OPS = @{ greaterThan = 5; lessThan = 6; between = 1; equal = 3 }
function Set-Preset($fc, $preset) {
  $p = $PRESETS[$preset]
  if ($null -ne $p[0]) { $fc.Interior.Color = $p[0] }
  if ($null -ne $p[1]) { $fc.Font.Color = $p[1] }
}
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$results = @()
try {
  $k = 0
  foreach ($case in $spec.cases) {
    $k++
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $words = @('banana', 'apple', 'cherry', 'band', 'apple', 'kiwi', 'mango', 'banana', 'pear', 'plum')
    $grid = New-Object 'object[,]' 10, 2
    for ($i = 0; $i -lt 10; $i++) { $grid[$i, 0] = $i + 1; $grid[$i, 1] = $words[$i] }
    $ws.Range('A1:B10').Value2 = $grid
    foreach ($rule in $case.rules) {
      $r = $ws.Range($rule.range)
      switch ($rule.kind) {
        'cellIs' { $fc = $r.FormatConditions.Add(1, $OPS[$rule.operator], $rule.f1, $(if ($rule.f2) { $rule.f2 } else { $m })); Set-Preset $fc ([string]$rule.preset) }
        'containsText' { $fc = $r.FormatConditions.Add(9, $m, $m, $m, $rule.text, 0); Set-Preset $fc ([string]$rule.preset) }
        'duplicate' { $fc = $r.FormatConditions.AddUniqueValues(); $fc.DupeUnique = 1; Set-Preset $fc ([string]$rule.preset) }
        'top10' { $fc = $r.FormatConditions.AddTop10(); $fc.TopBottom = $(if ($rule.bottom) { 0 } else { 1 }); $fc.Rank = $rule.rank; $fc.Percent = [bool]$rule.percent; Set-Preset $fc ([string]$rule.preset) }
        'average' { $fc = $r.FormatConditions.AddAboveAverage(); $fc.AboveBelow = $(if ($rule.below) { 1 } else { 0 }); Set-Preset $fc ([string]$rule.preset) }
        'colorScale' { [void]$r.FormatConditions.AddColorScale($rule.points) }
      }
    }
    $file = Join-Path $Books "cf-write-$k.xlsx"
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
    $results += [pscustomobject]@{ name = $case.name; conditionalFormatting = $cf; dxfs = $dxfs }
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$results | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 $Out
"$($results.Count) cases"
