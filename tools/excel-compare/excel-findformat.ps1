# What Excel's Find Format matches: colours given as a theme colour or as RGB, tints, fills, number formats, fonts,
# borders, alignment, and a whole cell's format set property by property (what Choose Format From Cell copies into
# Application.FindFormat). Each case on a new workbook of made-up cells in an invisible Excel; the cells found by
# Range.Find with SearchFormat, after each other round to the first (FindNext drops the format). With -Save, each
# case's workbook is saved to that folder (outside the repository) to read the styles Excel wrote.
# No keys sent, no setting saved; FindFormat and ReplaceFormat cleared after each case.
# usage: pwsh tools/excel-compare/excel-findformat.ps1 <out.json> [-Save <folder>]
param([Parameter(Mandatory)][string]$Out, [string]$Save)
$ErrorActionPreference = 'Stop'
trap { Write-Host "LINE $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; break }
$m = [Type]::Missing
function Found($ws, $what) {
  $first = $ws.Cells.Find($what, $ws.Cells.Item(1, 1), -4123, 2, 1, 1, $false, $m, $true)
  if (-not $first) { return @() }
  $out = @($first.Address($false, $false)); $at = $first
  for ($k = 0; $k -lt 40; $k++) {
    $at = $ws.Cells.Find($what, $at, -4123, 2, 1, 1, $false, $m, $true)
    if (-not $at -or $at.Address() -eq $first.Address()) { break }
    $out += $at.Address($false, $false)
  }
  return $out
}
# Column A: font colours. A1 automatic, A2 RGB black, A3 theme Text 1, A4 red, A5 accent 1, A6 accent 1's RGB,
# A7 accent 1 lighter 40%.
function Inks($ws) {
  foreach ($i in 1..7) { $ws.Cells.Item($i, 1).Value2 = 'x' }
  $ws.Range('A2').Font.Color = 0
  $ws.Range('A3').Font.ThemeColor = 2
  $ws.Range('A4').Font.Color = 255
  $ws.Range('A5').Font.ThemeColor = 5
  $ws.Range('A6').Font.Color = $ws.Parent.Theme.ThemeColorScheme.Colors(5).RGB
  $ws.Range('A7').Font.ThemeColor = 5; $ws.Range('A7').Font.TintAndShade = 0.4
}
# Column B: fills. B1 none, B2 RGB yellow, B3 accent 2, B4 accent 2's RGB, B5 accent 2 lighter 60%, B6 a pattern.
function Fills($ws) {
  foreach ($i in 1..6) { $ws.Cells.Item($i, 2).Value2 = 'x' }
  $ws.Range('B2').Interior.Color = 65535
  $ws.Range('B3').Interior.ThemeColor = 6
  $ws.Range('B4').Interior.Color = $ws.Parent.Theme.ThemeColorScheme.Colors(6).RGB
  $ws.Range('B5').Interior.ThemeColor = 6; $ws.Range('B5').Interior.TintAndShade = 0.6
  $ws.Range('B6').Interior.Pattern = 9; $ws.Range('B6').Interior.PatternColor = 255
}
# Column C: number formats, fonts. C1 General, C2 0.00, C3 #,##0.00, C4 Arial, C5 size 14, C6 Arial 14.
function Fonts($ws) {
  foreach ($i in 1..6) { $ws.Cells.Item($i, 3).Value2 = 5 }
  $ws.Range('C2').NumberFormat = '0.00'; $ws.Range('C3').NumberFormat = '#,##0.00'
  $ws.Range('C4').Font.Name = 'Arial'; $ws.Range('C5').Font.Size = 14
  $ws.Range('C6').Font.Name = 'Arial'; $ws.Range('C6').Font.Size = 14
}
# Column D: borders, alignment. D1 none, D2 thin bottom, D3 thick bottom, D4 thin all round, D5 red thin bottom,
# D6 centred, D7 wrapped, D8 centred and wrapped, D9 indent 2 (left), D10 unlocked.
function Lines($ws) {
  foreach ($i in 1..10) { $ws.Cells.Item($i, 4).Value2 = 'x' }
  $ws.Range('D2').Borders.Item(9).LineStyle = 1
  $ws.Range('D3').Borders.Item(9).LineStyle = 1; $ws.Range('D3').Borders.Item(9).Weight = 4
  $ws.Range('D4').BorderAround(1, 2) | Out-Null
  $ws.Range('D5').Borders.Item(9).LineStyle = 1; $ws.Range('D5').Borders.Item(9).Color = 255
  $ws.Range('D6').HorizontalAlignment = -4108
  $ws.Range('D7').WrapText = $true
  $ws.Range('D8').HorizontalAlignment = -4108; $ws.Range('D8').WrapText = $true
  $ws.Range('D9').HorizontalAlignment = -4131; $ws.Range('D9').IndentLevel = 2
  $ws.Range('D10').Locked = $false
}
# Column E: a whole format. E1 bold Arial 12 red on yellow, 0.00, centred, a thin bottom border; E2 the same; E3 also
# italic; E4 without the border; E5 plain; E6 the same but its colour accent 1 where E1 has red... (E6 = E1 with blue).
function Whole($ws) {
  foreach ($i in 1..6) { $ws.Cells.Item($i, 5).Value2 = 7 }
  foreach ($a in 'E1', 'E2', 'E3', 'E4', 'E6') {
    $r = $ws.Range($a); $r.Font.Bold = $true; $r.Font.Name = 'Arial'; $r.Font.Size = 12; $r.Font.Color = 255
    $r.Interior.Color = 65535; $r.NumberFormat = '0.00'; $r.HorizontalAlignment = -4108; $r.Borders.Item(9).LineStyle = 1
  }
  $ws.Range('E3').Font.Italic = $true
  $ws.Range('E4').Borders.Item(9).LineStyle = -4142
  $ws.Range('E6').Font.Color = 16711680
}
# FindFormat set as Choose Format From Cell would from E1: every property of its format.
function AsCell($c) {
  $f = $x.FindFormat; $f.Clear()
  $f.Font.Name = $c.Font.Name; $f.Font.Size = $c.Font.Size; $f.Font.Bold = $c.Font.Bold; $f.Font.Italic = $c.Font.Italic
  $f.Font.Underline = [int]$c.Font.Underline
  $f.Font.Strikethrough = [bool]$c.Font.Strikethrough
  $f.Font.Color = [int]$c.Font.Color
  $f.Interior.Pattern = [int]$c.Interior.Pattern; if ([int]$c.Interior.Pattern -ne -4142) { $f.Interior.Color = [int]$c.Interior.Color }
  $f.NumberFormat = [string]$c.NumberFormat; $f.HorizontalAlignment = [int]$c.HorizontalAlignment; $f.VerticalAlignment = [int]$c.VerticalAlignment
  $f.WrapText = [bool]$c.WrapText; $f.IndentLevel = [int]$c.IndentLevel; $f.Orientation = [int]$c.Orientation
  $f.Locked = [bool]$c.Locked; $f.FormulaHidden = [bool]$c.FormulaHidden
  foreach ($k in 7, 8, 9, 10) {
    $b = $c.Borders.Item($k); $fb = $f.Borders.Item($k); $fb.LineStyle = [int]$b.LineStyle
    if ([int]$b.LineStyle -ne -4142) { $fb.Weight = [int]$b.Weight; $fb.Color = [int]$b.Color }
  }
}
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$results = @()
try {
  $cases = [ordered]@{
    'ink rgb black' = { param($ws) Inks $ws; $x.FindFormat.Font.Color = 0 }
    'ink theme text 1' = { param($ws) Inks $ws; $x.FindFormat.Font.ThemeColor = 2 }
    'ink automatic' = { param($ws) Inks $ws; $x.FindFormat.Font.ColorIndex = -4105 }
    'ink red' = { param($ws) Inks $ws; $x.FindFormat.Font.Color = 255 }
    'ink theme accent 1' = { param($ws) Inks $ws; $x.FindFormat.Font.ThemeColor = 5 }
    'ink accent 1 rgb' = { param($ws) Inks $ws; $x.FindFormat.Font.Color = $ws.Parent.Theme.ThemeColorScheme.Colors(5).RGB }
    'fill yellow' = { param($ws) Fills $ws; $x.FindFormat.Interior.Color = 65535 }
    'fill theme accent 2' = { param($ws) Fills $ws; $x.FindFormat.Interior.ThemeColor = 6 }
    'fill accent 2 rgb' = { param($ws) Fills $ws; $x.FindFormat.Interior.Color = $ws.Parent.Theme.ThemeColorScheme.Colors(6).RGB }
    'fill none' = { param($ws) Fills $ws; $x.FindFormat.Interior.Pattern = -4142 }
    'numfmt 0.00' = { param($ws) Fonts $ws; $x.FindFormat.NumberFormat = '0.00' }
    'numfmt general' = { param($ws) Fonts $ws; $x.FindFormat.NumberFormat = 'General' }
    'font arial' = { param($ws) Fonts $ws; $x.FindFormat.Font.Name = 'Arial' }
    'font size 14' = { param($ws) Fonts $ws; $x.FindFormat.Font.Size = 14 }
    'border bottom thin' = { param($ws) Lines $ws; $x.FindFormat.Borders.Item(9).LineStyle = 1; $x.FindFormat.Borders.Item(9).Weight = 2 }
    'border bottom any' = { param($ws) Lines $ws; $x.FindFormat.Borders.Item(9).LineStyle = 1 }
    'border bottom red' = { param($ws) Lines $ws; $x.FindFormat.Borders.Item(9).LineStyle = 1; $x.FindFormat.Borders.Item(9).Color = 255 }
    'border none bottom' = { param($ws) Lines $ws; $x.FindFormat.Borders.Item(9).LineStyle = -4142 }
    'align centre' = { param($ws) Lines $ws; $x.FindFormat.HorizontalAlignment = -4108 }
    'align wrap' = { param($ws) Lines $ws; $x.FindFormat.WrapText = $true }
    'align indent 2' = { param($ws) Lines $ws; $x.FindFormat.IndentLevel = 2 }
    'unlocked' = { param($ws) Lines $ws; $x.FindFormat.Locked = $false }
    'whole format of E1' = { param($ws) Whole $ws; AsCell $ws.Range('E1') }
    'whole format of E5 (plain)' = { param($ws) Whole $ws; AsCell $ws.Range('E5') }
  }
  $n = 0
  foreach ($name in $cases.Keys) {
    $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
    Write-Host "case: $name"
    $x.FindFormat.Clear()
    & $cases[$name] $ws
    $found = @(Found $ws '*')
    $results += [pscustomobject]@{ name = $name; found = $found }
    $x.FindFormat.Clear(); $x.ReplaceFormat.Clear()
    if ($Save) { $n++; $wb.SaveAs((Join-Path $Save "ff$n.xlsx"), 51) }
    $wb.Close($false)
  }
} finally { $x.FindFormat.Clear(); $x.ReplaceFormat.Clear(); $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x); [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
$results | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 $Out
"$($results.Count) cases"
