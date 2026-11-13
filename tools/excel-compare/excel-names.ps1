# Which names Excel takes for a defined name (Formulas > Define Name), and how it saves them, for names.test.js: in an
# invisible Excel, a new workbook of two sheets; each candidate added with Names.Add (refers to =Sheet1!$A$1) and
# whether Excel took it recorded (COM's error wording isn't the dialog's, so only the verdict); then names added out of
# order, one local to Sheet2 and one hidden, the workbook saved outside the repository and its definedNames read back.
# No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-names.ps1 <out.json> <folder for the saved workbook>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$CANDIDATES = @('Rate', 'rate2', 'my rate', '1abc', '_x', '\x', 'Tax.Rate', 'A1', 'a1', 'XFD1', 'XFE1', 'AB12C', 'R1C1', 'R', 'C', 'r', 'c', 'RC',
  'R2', 'C3', 'RC1', 'R1C', 'Rate?', 'Q1-Q2', 'Total!', 'Sales%', ('n' * 255), ('n' * 256), 'TRUE', 'Print_Area', 'abc d', 'é', 'Ünits', 'x$', '_', '\', 'a\b', 'x\', 'x\y\')
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$verdicts = @()
try {
  $wb = $x.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $wb.Worksheets.Item(1).Name = 'Sheet1'; $wb.Worksheets.Item(2).Name = 'Sheet2'
  foreach ($n in $CANDIDATES) {
    $ok = $true
    try { [void]$wb.Names.Add($n, '=Sheet1!$A$1'); $wb.Names.Item($n).Delete() } catch { $ok = $false }
    $verdicts += [pscustomobject]@{ name = $n; ok = $ok }
  }
  # A second name with the same name, workbook-wide and on a sheet (Excel's rules for duplicates).
  [void]$wb.Names.Add('Dup', '=Sheet1!$A$1')
  $same = $true; try { [void]$wb.Names.Add('dup', '=Sheet1!$B$1') } catch { $same = $false }
  $local = $true; try { [void]$wb.Worksheets.Item(2).Names.Add('Dup', '=Sheet2!$A$1') } catch { $local = $false }
  $dupRefers = $wb.Names.Item('Dup').RefersTo
  # Saved: names added out of order, a local one, a hidden one, a constant and a formula.
  foreach ($n in $wb.Names) { }
  [void]$wb.Names.Add('Zeta', '=Sheet1!$C$3')
  [void]$wb.Names.Add('alpha', '=Sheet1!$A$1:$B$2')
  [void]$wb.Worksheets.Item(2).Names.Add('Local', '=Sheet2!$D$4')
  [void]$wb.Names.Add('Beta', '=0.07')
  [void]$wb.Names.Add('Gamma', '=SUM(Sheet1!$A$1:$A$3)*2')
  $h = $wb.Names.Add('Hidden1', '=Sheet1!$E$5'); $h.Visible = $false
  [void]$wb.Names.Add('_under', '="text"')
  $file = Join-Path $Books 'names.xlsx'
  if (Test-Path $file) { Remove-Item $file }
  $wb.SaveAs($file, 51)
  $wb.Close($false)
  $zip = [IO.Compression.ZipFile]::OpenRead($file)
  $s = New-Object IO.StreamReader($zip.GetEntry('xl/workbook.xml').Open()); $book = $s.ReadToEnd(); $s.Close(); $zip.Dispose()
  $names = ([regex]::Match($book, '<definedNames>[\s\S]*?</definedNames>')).Value
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
[pscustomobject]@{ verdicts = $verdicts; sameNameAgain = $same; sameNameOnASheet = $local; dupRefersTo = $dupRefers; definedNames = $names } | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 $Out
'done'
