# How Excel saves the borders and colours Home > Font's Draw Borders and More Colors give, for drawborders.test.js:
# in an invisible Excel, a new workbook; each scene set through COM (Range.BorderAround, Borders, Interior.Color,
# Font.Color), the workbook saved outside the repository and every cell of each scene's window read back from
# styles.xml (its border, fill and font); the ribbon's presets run through ExecuteMso, and which Draw Borders mode is
# pressed after each of a list of steps (GetPressedMso). Then a <colors><mruColors> list (Excel's Recent Colors) put into the saved
# styles.xml, the file opened and saved again, to see whether Excel keeps it. No keys sent, no setting changed.
# (Draw Border itself is a mouse gesture COM can't make: the scenes are the same edges set through COM.)
# usage: pwsh tools/excel-compare/excel-drawborders.ps1 <out.json> <folder for the saved workbooks>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
# Borders.Item: 5 diagonal down, 6 diagonal up, 7 left, 8 top, 9 bottom, 10 right, 11 inside vertical, 12 inside
# horizontal. LineStyle 1 continuous, -4115 dash, 4 dash dot, 5 dash dot dot, -4118 dot, -4119 double, 13 slant dash
# dot, -4142 none. Weight 1 hairline, 2 thin, -4138 medium, 4 thick.
# (The numbers typed [int]: passed as a script block's arguments they came as strings, which LineStyle refused.)
$edge = { param($r, [int]$k, [int]$ls, [int]$w, $color) $b = $r.Borders.Item($k); $b.LineStyle = $ls; if ($w) { $b.Weight = $w }; if ($null -ne $color) { $b.Color = $color } }
# Each scene: a name, the window of cells read back, and a script block given the sheet.
$SCENES = [ordered]@{
  # Draw Border over B2:D4 (an outline), its neighbours holding lines first: B1 a double bottom, E3 a red medium left,
  # A3 an automatic thin right (the same line as the outline's).
  'outline' = @('A1:E5', {
    param($ws)
    & $edge $ws.Range('B1') 9 -4119 4 $null
    & $edge $ws.Range('E3') 7 1 -4138 0x0000FF
    & $edge $ws.Range('A3') 10 1 2 $null
    [void]$ws.Range('B2:D4').BorderAround(1, 2)
  })
  # Draw Border Grid over B8:D10: every edge and inside line.
  'grid' = @('A7:E11', {
    param($ws)
    & $edge $ws.Range('E9') 7 1 -4138 0x0000FF
    $r = $ws.Range('B8:D10')
    foreach ($k in 7, 8, 9, 10, 11, 12) { & $edge $r $k 1 2 $null }
  })
  # Erase Border over B14:D16 holding a grid of thin lines and a diagonal in C15; neighbours A15 (right) and E15 (left)
  # and B13 (bottom) with lines on the facing sides.
  'erase' = @('A13:E17', {
    param($ws)
    $r = $ws.Range('B14:D16')
    foreach ($k in 7, 8, 9, 10, 11, 12) { & $edge $r $k 1 2 $null }
    & $edge $ws.Range('C15') 5 1 2 $null
    & $edge $ws.Range('A15') 10 1 -4138 $null
    & $edge $ws.Range('E15') 7 1 -4138 $null
    & $edge $ws.Range('B13') 9 -4119 4 $null
    $r.Borders.LineStyle = -4142
  })
  # The same with each edge and inside line set to none one by one (Borders.Item 7 to 12), diagonals left.
  'eraseEdges' = @('G13:K17', {
    param($ws)
    $r = $ws.Range('H14:J16')
    foreach ($k in 7, 8, 9, 10, 11, 12) { & $edge $r $k 1 2 $null }
    & $edge $ws.Range('I15') 5 1 2 $null
    & $edge $ws.Range('G15') 10 1 -4138 $null
    & $edge $ws.Range('K15') 7 1 -4138 $null
    foreach ($k in 7, 8, 9, 10, 11, 12) { $r.Borders.Item($k).LineStyle = -4142 }
  })
  # The ribbon's own presets (ExecuteMso, as docs/FORMATS.md measured them) over the same neighbours: Outside Borders
  # as Draw Border's outline, All Borders as Draw Border Grid's, No Border as Erase Border's.
  'msoOutline' = @('G1:K5', {
    param($ws)
    & $edge $ws.Range('H1') 9 -4119 4 $null
    & $edge $ws.Range('K3') 7 1 -4138 0x0000FF
    & $edge $ws.Range('G3') 10 1 2 $null
    $ws.Range('H2:J4').Select(); $x.CommandBars.ExecuteMso('BorderOutside')
  })
  'msoGrid' = @('G7:K11', {
    param($ws)
    & $edge $ws.Range('K9') 7 1 -4138 0x0000FF
    & $edge $ws.Range('H7') 9 1 2 $null
    $ws.Range('H8:J10').Select(); $x.CommandBars.ExecuteMso('BordersAll')
  })
  # Top Border on H20:I20 under H19's thin bottom and I19's medium one (docs/FORMATS.md: the thin one stayed), and
  # Outside Borders on H24 beside G24's thin right and I24's thin left.
  'msoTop' = @('G19:J25', {
    param($ws)
    & $edge $ws.Range('H19') 9 1 2 $null
    & $edge $ws.Range('I19') 9 1 -4138 $null
    $ws.Range('H20:I20').Select(); $x.CommandBars.ExecuteMso('BorderTop')
    & $edge $ws.Range('G24') 10 1 2 $null
    & $edge $ws.Range('I24') 7 1 2 $null
    $ws.Range('H24').Select(); $x.CommandBars.ExecuteMso('BorderOutside')
  })
  'msoErase' = @('M13:Q17', {
    param($ws)
    $r = $ws.Range('N14:P16')
    foreach ($k in 7, 8, 9, 10, 11, 12) { & $edge $r $k 1 2 $null }
    & $edge $ws.Range('O15') 5 1 2 $null
    & $edge $ws.Range('M15') 10 1 -4138 $null
    & $edge $ws.Range('Q15') 7 1 -4138 $null
    & $edge $ws.Range('N13') 9 -4119 4 $null
    $r.Select(); $x.CommandBars.ExecuteMso('BorderNone')
  })
  # The Line Style list's thirteen lines, one per cell's bottom (A20 down), and a custom RGB 12 34 56 colour.
  'styles' = @('A20:A32', {
    param($ws)
    $lines = @(@(1, 1), @(-4118, 2), @(5, 2), @(4, 2), @(-4115, 2), @(1, 2), @(5, -4138), @(13, -4138), @(4, -4138), @(-4115, -4138), @(1, -4138), @(1, 4), @(-4119, 4))
    for ($i = 0; $i -lt $lines.Count; $i++) { & $edge $ws.Range("A$(20 + $i)") 9 $lines[$i][0] $lines[$i][1] $null }
  })
  # More Colors: a custom colour (RGB 18, 52, 86) as a line, a fill and a font colour.
  'custom' = @('C20:C22', {
    param($ws)
    & $edge $ws.Range('C20') 9 1 2 0x563412
    $ws.Range('C21').Interior.Color = 0x563412
    $ws.Range('C22').Value2 = 'x'
    $ws.Range('C22').Font.Color = 0x563412
  })
}
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  foreach ($name in $SCENES.Keys) { & $SCENES[$name][1] $ws }
  # The three draw modes (Draw Border is BorderDrawLine): which is pressed after each step.
  $cb = $x.CommandBars
  $pressed = { (@('BorderDrawLine', 'BorderDrawGrid', 'BorderErase') | Where-Object { $cb.GetPressedMso($_) }) -join ',' }
  $modes = [ordered]@{ start = & $pressed }
  foreach ($step in 'BorderDrawLine', 'BorderDrawGrid', 'BorderErase', 'BorderErase', 'BorderDrawLine', 'select', 'BorderDrawGrid', 'value') {
    if ($step -eq 'select') { $ws.Range('D4').Select() } elseif ($step -eq 'value') { $ws.Range('Z1').Value2 = 5 } else { $cb.ExecuteMso($step) }
    $modes["$($modes.Count): $step"] = & $pressed
  }
  if (& $pressed) { $cb.ExecuteMso((& $pressed)) }
  $file = Join-Path $Books 'drawborders.xlsx'
  if (Test-Path $file) { Remove-Item $file }
  $wb.SaveAs($file, 51)
  $wb.Close($false)
  $zip = [IO.Compression.ZipFile]::OpenRead($file)
  $read = { param($n) $s = New-Object IO.StreamReader($zip.GetEntry($n).Open()); $t = $s.ReadToEnd(); $s.Close(); $t }
  $styles = & $read 'xl/styles.xml'
  $sheet = & $read 'xl/worksheets/sheet1.xml'
  $zip.Dispose()
  # Recent Colors: the saved file given an mruColors list, opened and saved again by Excel.
  $mruFile = Join-Path $Books 'mru.xlsx'
  $mruAgain = Join-Path $Books 'mru-again.xlsx'
  foreach ($f in $mruFile, $mruAgain) { if (Test-Path $f) { Remove-Item $f } }
  Copy-Item $file $mruFile
  $zip = [IO.Compression.ZipFile]::Open($mruFile, 'Update')
  $entry = $zip.GetEntry('xl/styles.xml')
  $s = New-Object IO.StreamReader($entry.Open()); $text = $s.ReadToEnd(); $s.Close()
  $mru = '<colors><mruColors><color rgb="FF123456"/><color rgb="FFABCDEF"/></mruColors></colors>'
  # (colors sits after tableStyles, before extLst, by the schema's order; a new workbook has none.)
  if ($text -match '<colors>') { throw 'the new workbook already has colors' }
  $text = $text -replace '(<tableStyles[^>]*?(/>|>[\s\S]*?</tableStyles>))', "`$1$mru"
  if ($text -notmatch '<mruColors>') { throw 'no tableStyles to put the colours after' }
  $entry.Delete()
  $w = New-Object IO.StreamWriter($zip.CreateEntry('xl/styles.xml').Open(), (New-Object Text.UTF8Encoding($false)))
  $w.Write($text); $w.Close()
  $zip.Dispose()
  $wb = $x.Workbooks.Open($mruFile)
  $wb.SaveAs($mruAgain, 51)
  $wb.Close($false)
  $zip = [IO.Compression.ZipFile]::OpenRead($mruAgain)
  $again = & $read 'xl/styles.xml'
  $zip.Dispose()
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$list = { param($tag, $item) $m = [regex]::Match($styles, "<$tag[^>]*>([\s\S]*?)</$tag>"); [regex]::Matches($m.Groups[1].Value, "<$item(?:\s[^>]*?)?(?:/>|>[\s\S]*?</$item>)") | ForEach-Object { $_.Value } }
$fonts = @(& $list 'fonts' 'font')
$fills = @(& $list 'fills' 'fill')
$borders = @(& $list 'borders' 'border')
$xfs = @(& $list 'cellXfs' 'xf')
$letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
$found = [ordered]@{}
foreach ($name in $SCENES.Keys) {
  $m = [regex]::Match($SCENES[$name][0], '^([A-Z])(\d+):([A-Z])(\d+)$')
  $cells = [ordered]@{}
  foreach ($r in [int]$m.Groups[2].Value..[int]$m.Groups[4].Value) {
    foreach ($ci in $letters.IndexOf($m.Groups[1].Value)..$letters.IndexOf($m.Groups[3].Value)) {
      $ref = "$($letters[$ci])$r"
      $mm = [regex]::Match($sheet, "<c r=""$ref""[^>]*?\ss=""(\d+)""")
      $s = if ($mm.Success) { [int]$mm.Groups[1].Value } else { 0 }
      $xf = $xfs[$s]
      $id = { param($a) $q = [regex]::Match($xf, "$a=""(\d+)"""); if ($q.Success) { [int]$q.Groups[1].Value } else { 0 } }
      $cells[$ref] = [pscustomobject]@{ s = $s; border = $borders[(& $id 'borderId')]; fill = $fills[(& $id 'fillId')]; font = $fonts[(& $id 'fontId')] }
    }
  }
  $found[$name] = $cells
}
$mruKept = [regex]::Match($again, '<colors>[\s\S]*?</colors>').Value
[pscustomobject]@{ scenes = $found; modes = $modes; mruAfterSave = $mruKept } | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 $Out
'done'
