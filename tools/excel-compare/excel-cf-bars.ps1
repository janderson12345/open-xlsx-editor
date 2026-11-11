# What Excel writes for data bars and icon sets, for cfbars.test.js: in an invisible Excel, a new workbook with the
# numbers 1 to 10 in each case's column (one case also with negatives); each case one rule added as Home >
# Conditional Formatting adds it (FormatConditions.AddDatabar with the menu's colour and fill, AddIconSetCondition with
# the menu's icon set); saved outside the repository and the sheet's conditionalFormatting, its extLst and the
# workbook's dxfs read back. No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-cf-bars.ps1 <out.json> <folder for the saved workbook>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
# Each case: [kind, colour (BGR for COM) or icon set, solid fill?]
$CASES = [ordered]@{
  'barBlueGradient' = @('bar', 0xC68E63, $false); 'barGreenGradient' = @('bar', 0x84C363, $false); 'barRedGradient' = @('bar', 0x5A55FF, $false)
  'barOrangeGradient' = @('bar', 0x28B6FF, $false); 'barLightBlueGradient' = @('bar', 0xEF8A00, $false); 'barPurpleGradient' = @('bar', 0x7B00D6, $false)
  'barBlueSolid' = @('bar', 0xC68E63, $true); 'barGreenSolid' = @('bar', 0x84C363, $true); 'barRedSolid' = @('bar', 0x5A55FF, $true)
  'barOrangeSolid' = @('bar', 0x28B6FF, $true); 'barLightBlueSolid' = @('bar', 0xEF8A00, $true); 'barPurpleSolid' = @('bar', 0x7B00D6, $true)
  'barNegative' = @('bar', 0xC68E63, $false)
  'icons3Arrows' = @('icons', 1); 'icons3ArrowsGray' = @('icons', 2); 'icons3Flags' = @('icons', 3); 'icons3TrafficLights1' = @('icons', 4)
  'icons3TrafficLights2' = @('icons', 5); 'icons3Signs' = @('icons', 6); 'icons3Symbols' = @('icons', 7); 'icons3Symbols2' = @('icons', 8)
  'icons4Arrows' = @('icons', 9); 'icons4ArrowsGray' = @('icons', 10); 'icons4RedToBlack' = @('icons', 11); 'icons4Rating' = @('icons', 12)
  'icons4TrafficLights' = @('icons', 13); 'icons5Arrows' = @('icons', 14); 'icons5ArrowsGray' = @('icons', 15); 'icons5Rating' = @('icons', 16)
  'icons5Quarters' = @('icons', 17); 'icons3Stars' = @('icons', 18); 'icons3Triangles' = @('icons', 19); 'icons5Boxes' = @('icons', 20)
}
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$cols = @{}
try {
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $c = 1
  foreach ($name in $CASES.Keys) {
    $col = $ws.Cells.Item(1, $c).Address($false, $false) -replace '\d', ''
    for ($r = 1; $r -le 10; $r++) { $ws.Range("$col$r").Formula = "$(if ($name -eq 'barNegative') { $r - 4 } else { $r })" }
    $range = $ws.Range("${col}1:${col}10")
    $k = $CASES[$name]
    if ($k[0] -eq 'bar') {
      $db = $range.FormatConditions.AddDatabar()
      $db.BarColor.Color = $k[1]
      $db.BarFillType = $(if ($k[2]) { 0 } else { 1 })
      # (As the ribbon's Data Bars give it: automatic minimum and maximum, the axis automatic, a gradient bar bordered in
      # its colour, negative bars red.)
      [void]$db.MinPoint.Modify(6); [void]$db.MaxPoint.Modify(7); $db.AxisPosition = 0
      if (-not $k[2]) { $db.BarBorder.Type = 1; $db.BarBorder.Color.Color = $k[1] }
      $db.NegativeBarFormat.ColorType = 0; $db.NegativeBarFormat.Color.Color = 0x0000FF
      if (-not $k[2]) { $db.NegativeBarFormat.BorderColorType = 0; $db.NegativeBarFormat.BorderColor.Color = 0x0000FF }
    } else {
      $is = $range.FormatConditions.AddIconSetCondition()
      $is.IconSet = $wb.IconSets.Item($k[1])
    }
    $cols[$name] = $col
    $c++
  }
  $file = Join-Path $Books 'cf-bars.xlsx'
  if (Test-Path $file) { Remove-Item $file }
  $wb.SaveAs($file, 51)
  $wb.Close($false)
  $zip = [IO.Compression.ZipFile]::OpenRead($file)
  $read = { param($n) $s = New-Object IO.StreamReader($zip.GetEntry($n).Open()); $t = $s.ReadToEnd(); $s.Close(); $t }
  $sheet = & $read 'xl/worksheets/sheet1.xml'
  $styles = & $read 'xl/styles.xml'
  $zip.Dispose()
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$cf = @([regex]::Matches($sheet, '<conditionalFormatting\b[\s\S]*?</conditionalFormatting>') | ForEach-Object { $_.Value })
$ext = [regex]::Match($sheet, '<extLst>(?:(?!<extLst>)[\s\S])*</extLst>(?=</worksheet>)').Value
$root = [regex]::Match($sheet, '<worksheet\b[^>]*>').Value
[pscustomobject]@{ columns = $cols; root = $root; conditionalFormatting = $cf; extLst = $ext; dxfs = [regex]::Match($styles, '<dxfs\b[\s\S]*?(/>|</dxfs>)').Value } | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 $Out
'done'
