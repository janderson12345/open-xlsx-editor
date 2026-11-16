# Has Excel say what its built-in table styles are made of: each built-in table style (TableStyleLight1..21,
# Medium1..28, Dark1..11) is duplicated in a new workbook (TableStyle.Duplicate, as Table Design's gallery > right click
# > Duplicate does), which Excel saves as a custom style of differential formats (dxfs) in styles.xml's tableStyles.
# The saved <tableStyle> of each, with the <dxf>s it points to, is written out; the rest of the package (docProps holds
# the user's name) is not. Also each style's stripe sizes as COM reports them. An invisible instance of its own, macros
# disabled for it only, quit at the end; no Excel setting is changed and no keys are sent.
# usage: pwsh tools/excel-compare/excel-tablestyles.ps1 test/tablestyles-excel.json
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$dir = Join-Path ([IO.Path]::GetTempPath()) ("oxe-tablestyles-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $dir | Out-Null
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$names = @()
foreach ($n in 1..21) { $names += "TableStyleLight$n" }
foreach ($n in 1..28) { $names += "TableStyleMedium$n" }
foreach ($n in 1..11) { $names += "TableStyleDark$n" }
$answers = [ordered]@{ excel = $x.Version; styles = [ordered]@{} }
try {
  $wb = $x.Workbooks.Add()
  $copies = @{}
  foreach ($n in $names) {
    $ts = $wb.TableStyles.Item($n)
    $copy = $ts.Duplicate("Copy of $n")
    $copies[$n] = "Copy of $n"
    # (Stripe sizes: TableStyleElements 5..8 are the row and column stripes, xlRowStripe1 = 5 ... xlColumnStripe2 = 8.)
    $stripes = [ordered]@{}
    foreach ($k in 5..8) { $stripes["$k"] = $ts.TableStyleElements.Item($k).StripeSize }
    $answers.styles[$n] = [ordered]@{ stripes = $stripes }
  }
  $file = Join-Path $dir 'styles.xlsx'
  $wb.SaveAs($file, 51)
  $wb.Close($false)
  $zip = [IO.Compression.ZipFile]::OpenRead($file)
  try { $styles = (New-Object IO.StreamReader($zip.GetEntry('xl/styles.xml').Open())).ReadToEnd() } finally { $zip.Dispose() }
  $dxfs = [regex]::Matches(($styles -replace '(?s)^.*?<dxfs[^>]*>(.*?)</dxfs>.*$', '$1'), '(?s)<dxf>.*?</dxf>|<dxf/>') | ForEach-Object { $_.Value }
  foreach ($n in $names) {
    $pat = '(?s)<tableStyle name="' + [regex]::Escape($copies[$n]) + '"[^>]*>.*?</tableStyle>'
    $m = [regex]::Match($styles, $pat)
    $entry = $answers.styles[$n]
    $entry.tableStyle = $m.Value
    $entry.dxfs = [ordered]@{}
    foreach ($e in [regex]::Matches($m.Value, 'type="(\w+)"(?: size="(\d+)")? dxfId="(\d+)"')) {
      $entry.dxfs[$e.Groups[1].Value] = $dxfs[[int]$e.Groups[3].Value]
    }
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  Remove-Item -Recurse -Force $dir -ErrorAction SilentlyContinue
}
$json = $answers | ConvertTo-Json -Depth 6
[IO.File]::WriteAllText(([IO.Path]::GetFullPath($Out, (Get-Location).Path)), $json + "`n", (New-Object Text.UTF8Encoding($false)))
