# Has Excel say what it saves when a cell style is deleted, duplicated and merged in from another workbook: Delete is
# Style.Delete, for a style of the workbook's own and a built-in one, each with cells using it (one with a format of its
# own over the style), and for Normal (refused); Duplicate is the Style box on a copy, which COM doesn't reach, so it is
# made as the box makes it: Styles.Add("<name> 2", BasedOn) on a cell of the style alone, its Include... set as the
# style's; Merge is Styles.Merge(other workbook), with styles of the same name in both and without. Each case a new
# workbook saved to a temporary folder; written out are styles.xml's numFmts, fonts, fills, borders, cellStyleXfs,
# cellXfs and cellStyles and the sheet's sheetData (the rest of the package, docProps with the user's name, is not).
# An invisible instance, macros disabled for it only; no setting changed, no keys sent.
# usage: pwsh tools/excel-compare/excel-stylesmore.ps1 test/stylesmore-excel.json
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$dir = Join-Path ([IO.Path]::GetTempPath()) ("oxe-stylesmore-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $dir | Out-Null
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$answers = [ordered]@{}
$save = {
  param($wb, $key, $notes)
  $file = Join-Path $dir "$($answers.Count).xlsx"
  $wb.SaveAs($file, 51)
  $wb.Close($false)
  $zip = [IO.Compression.ZipFile]::OpenRead($file)
  try {
    $styles = (New-Object IO.StreamReader($zip.GetEntry('xl/styles.xml').Open())).ReadToEnd()
    $sheet = (New-Object IO.StreamReader($zip.GetEntry('xl/worksheets/sheet1.xml').Open())).ReadToEnd()
  } finally { $zip.Dispose() }
  $s = [ordered]@{}
  foreach ($part in 'numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles') {
    $s[$part] = if ($styles -match "(?s)<$part\b.*?</$part>") { $Matches[0] } else { $null }
  }
  $s.sheetData = if ($sheet -match '(?s)<sheetData\b.*?(</sheetData>|/>)') { $Matches[0] } else { $null }
  if ($notes) { $s.notes = $notes }
  $answers[$key] = $s
}
$look = {
  param($ws, $refs)
  $o = [ordered]@{}
  foreach ($a in $refs) {
    $c = $ws.Range($a)
    $o[$a] = [ordered]@{ style = $c.Style.Name; bold = $c.Font.Bold; italic = $c.Font.Italic; size = $c.Font.Size; color = $c.Font.Color; fill = $c.Interior.Color; pattern = $c.Interior.Pattern; bottom = $c.Borders.Item(9).LineStyle; format = $c.NumberFormat; align = $c.HorizontalAlignment }
  }
  $o
}
# B2: bold red 14pt, yellow fill, a thin bottom border, 0.00, centred; style Mine made from it, given to D4 and D5 (D5
# then italic of its own).
$mine = {
  param($wb)
  $ws = $wb.Worksheets.Item(1)
  $c = $ws.Range('B2')
  $c.Value2 = 1
  $c.Font.Bold = $true; $c.Font.Color = 255; $c.Font.Size = 14
  $c.Interior.Color = 65535
  $c.Borders.Item(9).LineStyle = 1
  $c.NumberFormat = '0.00'
  $c.HorizontalAlignment = -4108
  [void]$wb.Styles.Add('Mine', $c)
  foreach ($a in 'D4', 'D5') { $ws.Range($a).Value2 = 2; $ws.Range($a).Style = 'Mine' }
  $ws.Range('D5').Font.Italic = $true
  $ws
}
$names = { param($wb) @($wb.Styles | Where-Object { -not $_.BuiltIn } | ForEach-Object { $_.Name }) }
try {
  # Delete a style of the workbook's own with cells using it.
  $wb = $x.Workbooks.Add(); $ws = & $mine $wb
  $wb.Styles.Item('Mine').Delete()
  & $save $wb 'delete Mine' (& $look $ws 'B2', 'D4', 'D5')

  # Delete a style of the workbook's own no cell uses (made from B2, not given to any cell).
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $ws.Range('B2').Value2 = 1; $ws.Range('B2').Font.Bold = $true
  [void]$wb.Styles.Add('Mine', $ws.Range('B2'))
  $wb.Styles.Item('Mine').Delete()
  & $save $wb 'delete unused Mine'

  # Delete a built-in style with cells using it: Good on B2 and B3 (B3 then red of its own), Title on B4.
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  foreach ($a in 'B2', 'B3') { $ws.Range($a).Value2 = 1; $ws.Range($a).Style = 'Good' }
  $ws.Range('B3').Font.Color = 255
  $ws.Range('B4').Value2 = 1; $ws.Range('B4').Style = 'Title'
  $wb.Styles.Item('Good').Delete()
  $after = & $look $ws 'B2', 'B3', 'B4'
  $after.goodStill = @($wb.Styles | Where-Object { $_.Name -eq 'Good' }).Count
  & $save $wb 'delete Good' $after

  # Normal can't be deleted.
  $wb = $x.Workbooks.Add()
  $why = try { [void]$wb.Styles.Item('Normal').Delete(); 'deleted' } catch { $_.Exception.Message }
  & $save $wb 'delete Normal' ([ordered]@{ error = $why })

  # Duplicate Mine: "Mine 2" made from a cell of Mine alone (Z1, then cleared), its Include... as Mine's.
  $wb = $x.Workbooks.Add(); $ws = & $mine $wb
  $wb.Styles.Item('Mine').IncludeBorder = $false
  $z = $ws.Range('Z1'); $z.Style = 'Mine'
  [void]$wb.Styles.Add('Mine 2', $z)
  $src = $wb.Styles.Item('Mine'); $dup = $wb.Styles.Item('Mine 2')
  foreach ($k in 'IncludeNumber', 'IncludeFont', 'IncludeAlignment', 'IncludeBorder', 'IncludePatterns', 'IncludeProtection') { $dup.$k = $src.$k }
  $z.Style = 'Normal'
  & $save $wb 'duplicate Mine' ([ordered]@{ names = (& $names $wb) })

  # Duplicate Good: "Good 2" the same way.
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $z = $ws.Range('Z1'); $z.Style = 'Good'
  [void]$wb.Styles.Add('Good 2', $z)
  $src = $wb.Styles.Item('Good'); $dup = $wb.Styles.Item('Good 2')
  foreach ($k in 'IncludeNumber', 'IncludeFont', 'IncludeAlignment', 'IncludeBorder', 'IncludePatterns', 'IncludeProtection') { $dup.$k = $src.$k }
  $z.Style = 'Normal'
  & $save $wb 'duplicate Good' ([ordered]@{ names = (& $names $wb) })

  # Merge from a workbook with styles of other names: Other (bold, 0.0" kg", green fill), Third (italic, left out
  # all but its font); the target has none of its own.
  $from = $x.Workbooks.Add(); $fs = $from.Worksheets.Item(1)
  $c = $fs.Range('A1'); $c.Font.Bold = $true; $c.NumberFormat = '0.0" kg"'; $c.Interior.Color = 5296274
  [void]$from.Styles.Add('Other', $c)
  $c = $fs.Range('A2'); $c.Font.Italic = $true
  [void]$from.Styles.Add('Third', $c)
  $t = $from.Styles.Item('Third'); $t.IncludeNumber = $false; $t.IncludeBorder = $false; $t.IncludePatterns = $false; $t.IncludeAlignment = $false; $t.IncludeProtection = $false
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $ws.Range('A1').Value2 = 1; $ws.Range('A1').NumberFormat = '0.000'
  $wb.Styles.Merge($from)
  & $save $wb 'merge new names' ([ordered]@{ names = (& $names $wb) })
  $from.Close($false)

  # Merge from a workbook with styles of the same names (DisplayAlerts off: Excel's question answered by its default):
  # the source's Mine green italic 12pt, a dashed bottom border, 0%, left; its Good made bold (a built-in style changed);
  # Other as above. The target: Mine on D4 and D5 (D5 italic of its own), Good on B6.
  $from = $x.Workbooks.Add(); $fs = $from.Worksheets.Item(1)
  $c = $fs.Range('A1'); $c.Font.Italic = $true; $c.Font.Size = 12; $c.Interior.Color = 5296274; $c.Borders.Item(9).LineStyle = -4115; $c.NumberFormat = '0%'; $c.HorizontalAlignment = -4131
  [void]$from.Styles.Add('Mine', $c)
  $from.Styles.Item('Good').Font.Bold = $true
  $c = $fs.Range('A2'); $c.Font.Bold = $true
  [void]$from.Styles.Add('Other', $c)
  $wb = $x.Workbooks.Add(); $ws = & $mine $wb
  $ws.Range('B6').Value2 = 1; $ws.Range('B6').Style = 'Good'
  $wb.Styles.Merge($from)
  $after = & $look $ws 'B2', 'D4', 'D5', 'B6'
  $after.names = (& $names $wb)
  & $save $wb 'merge same names' $after
  $from.Close($false)

  # Merge from a workbook whose built-in styles are Title on A1 (as it is), Normal at 14pt and Bad made italic; the
  # target's Title made bold, on B2, and C3 plain.
  $from = $x.Workbooks.Add(); $fs = $from.Worksheets.Item(1)
  $fs.Range('A1').Value2 = 1; $fs.Range('A1').Style = 'Title'
  $from.Styles.Item('Normal').Font.Size = 14
  $from.Styles.Item('Bad').Font.Italic = $true
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $ws.Range('B2').Value2 = 1; $ws.Range('B2').Style = 'Title'
  $wb.Styles.Item('Title').Font.Bold = $true
  $ws.Range('C3').Value2 = 1
  $wb.Styles.Merge($from)
  $after = & $look $ws 'B2', 'C3'
  & $save $wb 'merge built-ins' $after
  $from.Close($false)

  # Delete a built-in style the workbook hasn't got.
  $wb = $x.Workbooks.Add()
  $why = try { [void]$wb.Styles.Item('Good').Delete(); 'deleted' } catch { $_.Exception.Message }
  & $save $wb 'delete Good unused' ([ordered]@{ error = $why })

  # Delete Good (on B2), then give Good to B5.
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $ws.Range('B2').Value2 = 1; $ws.Range('B2').Style = 'Good'
  $wb.Styles.Item('Good').Delete()
  $ws.Range('B5').Value2 = 1
  $why = try { $ws.Range('B5').Style = 'Good'; 'applied' } catch { $_.Exception.Message }
  $after = & $look $ws 'B2', 'B5'
  $after.error = $why
  & $save $wb 'delete Good then apply' $after
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  Remove-Item -Recurse -Force $dir -ErrorAction SilentlyContinue
}
$json = $answers | ConvertTo-Json -Depth 6
[IO.File]::WriteAllText(([IO.Path]::GetFullPath($Out, (Get-Location).Path)), $json + "`n", (New-Object Text.UTF8Encoding($false)))
