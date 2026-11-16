# Has Excel say what it saves when a cell style of the workbook's own is made and when a style is changed: New Cell
# Style is Styles.Add(name, BasedOn) on a formatted cell (the Style box's "By Example"), with some of its parts left
# out (Include...); Modify is the style's Font, Interior, Borders, NumberFormat and HorizontalAlignment set, for a
# style of the workbook's own and for a built-in one, with cells using it (one with a format of its own over the
# style). Each case a new workbook saved to a temporary folder; written out are styles.xml's numFmts, fonts, fills,
# borders, cellStyleXfs, cellXfs and cellStyles and the sheet's sheetData (the rest of the package, docProps with the
# user's name, is not). An invisible instance, macros disabled for it only; no setting changed, no keys sent.
# usage: pwsh tools/excel-compare/excel-newstyle.ps1 test/newstyle-excel.json
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$dir = Join-Path ([IO.Path]::GetTempPath()) ("oxe-newstyle-" + [guid]::NewGuid().ToString('N'))
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
# B2: bold red 14pt, yellow fill, a thin bottom border, 0.00, centred.
$example = {
  param($ws)
  $c = $ws.Range('B2')
  $c.Value2 = 1
  $c.Font.Bold = $true; $c.Font.Color = 255; $c.Font.Size = 14
  $c.Interior.Color = 65535
  $c.Borders.Item(9).LineStyle = 1
  $c.NumberFormat = '0.00'
  $c.HorizontalAlignment = -4108
  $c
}
try {
  # New Cell Style from B2, every part included; B2's own style afterwards.
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $c = & $example $ws
  [void]$wb.Styles.Add('Mine', $c)
  & $save $wb 'new from B2' ([ordered]@{ b2Style = $c.Style.Name })

  # The same with its border and alignment left out.
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $c = & $example $ws
  [void]$wb.Styles.Add('Mine', $c); $st = $wb.Styles.Item('Mine')
  $st.IncludeBorder = $false; $st.IncludeAlignment = $false
  & $save $wb 'new from B2 without border and alignment'

  # New Cell Style from B2, given to D4 and D5 (D5 then italic of its own), then changed: not bold, a green fill, no
  # border, 0%, left-aligned.
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $c = & $example $ws
  [void]$wb.Styles.Add('Mine', $c); $st = $wb.Styles.Item('Mine')
  foreach ($a in 'D4', 'D5') { $ws.Range($a).Value2 = 2; $ws.Range($a).Style = 'Mine' }
  $ws.Range('D5').Font.Italic = $true
  & $save $wb 'new applied to D4 and D5'
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $c = & $example $ws
  [void]$wb.Styles.Add('Mine', $c); $st = $wb.Styles.Item('Mine')
  foreach ($a in 'D4', 'D5') { $ws.Range($a).Value2 = 2; $ws.Range($a).Style = 'Mine' }
  $ws.Range('D5').Font.Italic = $true
  $st.Font.Bold = $false
  $st.Interior.Color = 5296274
  $st.Borders.Item(9).LineStyle = -4142
  $st.NumberFormat = '0%'
  $st.HorizontalAlignment = -4131
  & $save $wb 'modified after D4 and D5' ([ordered]@{ d4Bold = $ws.Range('D4').Font.Bold; d5Italic = $ws.Range('D5').Font.Italic; d4Fill = $ws.Range('D4').Interior.Color })

  # A built-in style changed: Good given to B2 and B3 (B3 then red of its own), Good made bold with a thin outline.
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  foreach ($a in 'B2', 'B3') { $ws.Range($a).Value2 = 1; $ws.Range($a).Style = 'Good' }
  $ws.Range('B3').Font.Color = 255
  $g = $wb.Styles.Item('Good')
  $g.Font.Bold = $true
  foreach ($k in 7, 8, 9, 10) { $g.Borders.Item($k).LineStyle = 1 }
  & $save $wb 'Good modified'

  # Normal changed (its font size): every cell's default.
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $ws.Range('B2').Value2 = 1
  $ws.Range('C3').Value2 = 1; $ws.Range('C3').Font.Bold = $true
  $wb.Styles.Item('Normal').Font.Size = 14
  & $save $wb 'Normal modified'
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  Remove-Item -Recurse -Force $dir -ErrorAction SilentlyContinue
}
$json = $answers | ConvertTo-Json -Depth 6
[IO.File]::WriteAllText(([IO.Path]::GetFullPath($Out, (Get-Location).Path)), $json + "`n", (New-Object Text.UTF8Encoding($false)))
