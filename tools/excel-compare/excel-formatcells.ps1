# How Excel saves what the Format Cells box sets, for formatcells.test.js: in an invisible Excel, a new workbook; one
# cell per case formatted through COM (Range.Orientation, ShrinkToFit, Interior gradients and patterns, Borders,
# Font, Locked), the workbook saved outside the repository and, for each case's cell, its xf and the font, fill and
# border it points to read back from styles.xml. No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-formatcells.ps1 <out.json> <folder for the saved workbook>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
# Each case: a name and a script block given the cell (a Range).
$CASES = [ordered]@{
  'rotate45'       = { param($c) $c.Orientation = 45 }
  'rotateMinus90'  = { param($c) $c.Orientation = -90 }
  'rotate90'       = { param($c) $c.Orientation = 90 }
  'vertical'       = { param($c) $c.Orientation = -4166 }
  'shrink'         = { param($c) $c.ShrinkToFit = $true }
  'hFill'          = { param($c) $c.HorizontalAlignment = 5 }
  'hJustify'       = { param($c) $c.HorizontalAlignment = -4130 }
  'hDistributed'   = { param($c) $c.HorizontalAlignment = -4117 }
  'hCenterAcross'  = { param($c) $c.HorizontalAlignment = 7 }
  'vJustify'       = { param($c) $c.VerticalAlignment = -4130 }
  'vDistributed'   = { param($c) $c.VerticalAlignment = -4117 }
  'vTop'           = { param($c) $c.VerticalAlignment = -4160 }
  'readingRtl'     = { param($c) $c.ReadingOrder = -5004 }
  'gradientLinear' = { param($c) $c.Interior.Pattern = 4000; $c.Interior.Gradient.Degree = 90; $c.Interior.Gradient.ColorStops.Clear(); $s = $c.Interior.Gradient.ColorStops.Add(0); $s.Color = 0xFFFFFF; $s = $c.Interior.Gradient.ColorStops.Add(1); $s.Color = 0xC07000 }
  'gradientLinear45' = { param($c) $c.Interior.Pattern = 4000; $c.Interior.Gradient.Degree = 45; $c.Interior.Gradient.ColorStops.Clear(); $s = $c.Interior.Gradient.ColorStops.Add(0); $s.Color = 0x0000FF; $s = $c.Interior.Gradient.ColorStops.Add(0.5); $s.Color = 0xFFFFFF; $s = $c.Interior.Gradient.ColorStops.Add(1); $s.Color = 0x0000FF }
  'gradientPath'   = { param($c) $c.Interior.Pattern = 4001; $g = $c.Interior.Gradient; $g.RectangleLeft = 0.5; $g.RectangleRight = 0.5; $g.RectangleTop = 0.5; $g.RectangleBottom = 0.5; $g.ColorStops.Clear(); $s = $g.ColorStops.Add(0); $s.Color = 0xFFFFFF; $s = $g.ColorStops.Add(1); $s.Color = 0x00C0FF }
  'patternGray50'  = { param($c) $c.Interior.Pattern = -4125 }
  'patternDown'    = { param($c) $c.Interior.Pattern = -4121; $c.Interior.PatternColor = 0x0000FF; $c.Interior.Color = 0x00FFFF }
  'patternGray125' = { param($c) $c.Interior.Pattern = 17; $c.Interior.PatternColor = 0xFF0000 }
  'borderRedDashed' = { param($c) $b = $c.Borders.Item(9); $b.LineStyle = -4115; $b.Weight = 2; $b.Color = 0x0000FF }
  'borderDiagDown' = { param($c) $b = $c.Borders.Item(5); $b.LineStyle = 1; $b.Weight = 2 }
  'borderDiagUp'   = { param($c) $b = $c.Borders.Item(6); $b.LineStyle = 1; $b.Weight = 2 }
  'borderDiagBoth' = { param($c) $b = $c.Borders.Item(5); $b.LineStyle = 1; $b.Weight = 2; $b = $c.Borders.Item(6); $b.LineStyle = 1; $b.Weight = 2 }
  'borderLeftHair' = { param($c) $b = $c.Borders.Item(7); $b.LineStyle = 1; $b.Weight = 1 }
  'borderTopThick' = { param($c) $b = $c.Borders.Item(8); $b.LineStyle = 1; $b.Weight = 4; $b.ThemeColor = 5 }
  'borderSlant'    = { param($c) $b = $c.Borders.Item(10); $b.LineStyle = 13; $b.Weight = -4138 }
  'fontSuper'      = { param($c) $c.Font.Superscript = $true }
  'fontSub'        = { param($c) $c.Font.Subscript = $true }
  'fontDoubleU'    = { param($c) $c.Font.Underline = -4119 }
  'fontAccountingU' = { param($c) $c.Font.Underline = 4 }
  'fontDoubleAccU' = { param($c) $c.Font.Underline = 5 }
  'unlocked'       = { param($c) $c.Locked = $false }
  'hidden'         = { param($c) $c.FormulaHidden = $true }
  'unlockedHidden' = { param($c) $c.Locked = $false; $c.FormulaHidden = $true }
  'rotateShrinkIndent' = { param($c) $c.HorizontalAlignment = -4131; $c.IndentLevel = 2; $c.ShrinkToFit = $true; $c.Orientation = 30 }
  'rotateThenIndent' = { param($c) $c.Orientation = 30; $c.HorizontalAlignment = -4131; $c.IndentLevel = 2 }
  'indentThenRotate' = { param($c) $c.HorizontalAlignment = -4131; $c.IndentLevel = 2; $c.Orientation = 30 }
  'wrapThenShrink' = { param($c) $c.WrapText = $true; $c.ShrinkToFit = $true }
  'shrinkThenWrap' = { param($c) $c.ShrinkToFit = $true; $c.WrapText = $true }
  'fillThenRotate' = { param($c) $c.HorizontalAlignment = 5; $c.Orientation = 45 }
  'centerRotate'   = { param($c) $c.HorizontalAlignment = -4108; $c.Orientation = 45 }
  'gradientTheme'  = { param($c) $c.Interior.Pattern = 4000; $c.Interior.Gradient.Degree = 0; $g = $c.Interior.Gradient; $g.ColorStops.Clear(); $s = $g.ColorStops.Add(0); $s.ThemeColor = 1; $s = $g.ColorStops.Add(1); $s.ThemeColor = 5 }
  'solidPatternColor' = { param($c) $c.Interior.Color = 0x00FFFF; $c.Interior.Pattern = 1 }
}
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$at = @{}
try {
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $row = 1
  foreach ($name in $CASES.Keys) {
    $c = $ws.Range("A$row")
    $c.Value2 = $name
    & $CASES[$name] $c
    $at[$name] = $row
    $row += 2
  }
  # A range's borders (the Border tab on D2:F4): first E1 given a medium bottom and G3 a thin left; then D2:F4 an
  # outline of thin lines and inside lines dashed. Every cell of D1:G5 read back below.
  $b = $ws.Range('E1').Borders.Item(9); $b.LineStyle = 1; $b.Weight = -4138
  $b = $ws.Range('G3').Borders.Item(7); $b.LineStyle = 1; $b.Weight = 2
  $r = $ws.Range('D2:F4')
  foreach ($k in 7, 8, 9, 10) { $b = $r.Borders.Item($k); $b.LineStyle = 1; $b.Weight = 2 }
  foreach ($k in 11, 12) { $b = $r.Borders.Item($k); $b.LineStyle = -4115; $b.Weight = 2 }
  $file = Join-Path $Books 'formatcells.xlsx'
  if (Test-Path $file) { Remove-Item $file }
  $wb.SaveAs($file, 51)
  $wb.Close($false)
  $zip = [IO.Compression.ZipFile]::OpenRead($file)
  $read = { param($n) $s = New-Object IO.StreamReader($zip.GetEntry($n).Open()); $t = $s.ReadToEnd(); $s.Close(); $t }
  $styles = & $read 'xl/styles.xml'
  $sheet = & $read 'xl/worksheets/sheet1.xml'
  $zip.Dispose()
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$list = { param($tag, $item) $m = [regex]::Match($styles, "<$tag[^>]*>([\s\S]*?)</$tag>"); [regex]::Matches($m.Groups[1].Value, "<$item(?:\s[^>]*?)?(?:/>|>[\s\S]*?</$item>)") | ForEach-Object { $_.Value } }
$fonts = @(& $list 'fonts' 'font')
$fills = @(& $list 'fills' 'fill')
$borders = @(& $list 'borders' 'border')
$xfs = @(& $list 'cellXfs' 'xf')
$found = @()
foreach ($name in $CASES.Keys) {
  $m = [regex]::Match($sheet, "<c r=""A$($at[$name])""[^>]*?\ss=""(\d+)""")
  $s = if ($m.Success) { [int]$m.Groups[1].Value } else { 0 }
  $xf = $xfs[$s]
  $id = { param($a) $mm = [regex]::Match($xf, "$a=""(\d+)"""); if ($mm.Success) { [int]$mm.Groups[1].Value } else { 0 } }
  $found += [pscustomobject]@{ name = $name; xf = $xf; font = $fonts[(& $id 'fontId')]; fill = $fills[(& $id 'fillId')]; border = $borders[(& $id 'borderId')] }
}
$range = [ordered]@{}
foreach ($r in 1..5) { foreach ($col in 'D', 'E', 'F', 'G') {
  $m = [regex]::Match($sheet, "<c r=""$col$r""[^>]*?\ss=""(\d+)""")
  $s = if ($m.Success) { [int]$m.Groups[1].Value } else { 0 }
  $mm = [regex]::Match($xfs[$s], 'borderId="(\d+)"')
  $range["$col$r"] = $borders[$(if ($mm.Success) { [int]$mm.Groups[1].Value } else { 0 })]
} }
[pscustomobject]@{ cases = $found; range = $range } | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 $Out
'done'
