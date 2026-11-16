# Has Excel say what its built-in cell styles are and what it saves when one is applied: the workbook's Styles
# collection (each built-in style's name, what it includes and its font, fill, border and number format as COM reports
# them), then for each built-in style a new workbook with that style applied to B2 (Range.Style, as Home > Cell Styles
# applies it), saved to a temporary folder; written out are styles.xml's numFmts, fonts, fills, borders, cellStyleXfs,
# cellXfs and cellStyles and B2's <c> (the rest of the package, docProps with the user's name, is not). Also one
# workbook with nothing applied, and one with Good applied over a cell already bold and red. An invisible instance,
# macros disabled for it only; no setting changed, no keys sent.
# usage: pwsh tools/excel-compare/excel-cellstyles.ps1 test/cellstyles-excel.json
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$dir = Join-Path ([IO.Path]::GetTempPath()) ("oxe-cellstyles-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $dir | Out-Null
$hex = { param($v) if ($null -eq $v -or $v -is [DBNull]) { $null } else { $i = [int64]$v; '{0:X2}{1:X2}{2:X2}' -f ($i -band 255), (($i -shr 8) -band 255), (($i -shr 16) -band 255) } }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$answers = [ordered]@{ styles = @(); saved = [ordered]@{} }
$save = {
  param($wb, $key)
  $file = Join-Path $dir "$($answers.saved.Count).xlsx"
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
  $s.cell = if ($sheet -match '<c r="B2"[^>]*?(/>|>.*?</c>)') { $Matches[0] } else { $null }
  $answers.saved[$key] = $s
}
try {
  $wb = $x.Workbooks.Add()
  foreach ($st in $wb.Styles) {
    if (-not $st.BuiltIn) { continue }
    $answers.styles += [ordered]@{ name = $st.Name; local = $st.NameLocal
      includes = [ordered]@{ number = $st.IncludeNumber; font = $st.IncludeFont; alignment = $st.IncludeAlignment; border = $st.IncludeBorder; fill = $st.IncludePatterns; protection = $st.IncludeProtection }
      numberFormat = $st.NumberFormat
      font = [ordered]@{ name = $st.Font.Name; size = $st.Font.Size; bold = $st.Font.Bold; italic = $st.Font.Italic; color = (& $hex $st.Font.Color); theme = $st.Font.ThemeColor; tint = $st.Font.TintAndShade }
      fill = [ordered]@{ pattern = $st.Interior.Pattern; color = (& $hex $st.Interior.Color); theme = $st.Interior.ThemeColor; tint = $st.Interior.TintAndShade }
      bottom = [ordered]@{ line = $st.Borders.Item(9).LineStyle; weight = $st.Borders.Item(9).Weight; color = (& $hex $st.Borders.Item(9).Color) } }
  }
  $wb.Close($false)
  $wb = $x.Workbooks.Add(); $wb.Worksheets.Item(1).Range('B2').Value2 = 1; & $save $wb '(none)'
  foreach ($st in $answers.styles) {
    $wb = $x.Workbooks.Add()
    $c = $wb.Worksheets.Item(1).Range('B2')
    $c.Value2 = 1
    $c.Style = $st.name
    & $save $wb $st.name
  }
  $wb = $x.Workbooks.Add()
  $c = $wb.Worksheets.Item(1).Range('B2'); $c.Value2 = 1; $c.Font.Bold = $true; $c.Font.Color = 255; $c.Style = 'Good'
  & $save $wb 'Good over bold red'
  $wb = $x.Workbooks.Add()
  $c = $wb.Worksheets.Item(1).Range('B2'); $c.Value2 = 1; $c.Style = 'Good'; $c.Style = 'Heading 1'
  & $save $wb 'Good then Heading 1'
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  Remove-Item -Recurse -Force $dir -ErrorAction SilentlyContinue
}
$json = $answers | ConvertTo-Json -Depth 6
[IO.File]::WriteAllText(([IO.Path]::GetFullPath($Out, (Get-Location).Path)), $json + "`n", (New-Object Text.UTF8Encoding($false)))
