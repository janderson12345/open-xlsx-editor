# Has Excel add hyperlinks to files (Hyperlinks.Add, as the Insert Hyperlink box's Existing File or Web Page and Create
# New Document do): a name relative to the workbook, a full path in the workbook's own folder, in a folder under it and
# in another folder, one going up a folder, a file name with spaces and a place in that file, names with & % [ ] and an
# accented letter, a web address with a space, one already encoded, and a name with %41 { ^ }. A new workbook of
# made-up text in an invisible instance (macros disabled for it only) is saved in $Out (a folder outside the repo) and
# each link's Address and SubAddress read back before and after the save; the saved sheet's <hyperlinks> and its
# relationships are written to filelinks.json beside it (it has the folder's path: keep it out of the repo;
# filelinks-extract.mjs reduces it to test/filelinks-excel.json). No files are made for the links to go to.
# usage: pwsh tools/excel-compare/excel-filelinks.ps1 <folder outside the repo>
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $Out | Out-Null
$Out = (Resolve-Path $Out).Path
$book = Join-Path $Out 'filelinks.xlsx'
Remove-Item $book -ErrorAction SilentlyContinue
$parent = Split-Path $Out -Parent
$cases = [ordered]@{
  A1 = @('report.xlsx', '')
  A2 = @((Join-Path $Out 'New Book.xlsx'), '')
  A3 = @((Join-Path $Out 'sub\notes.docx'), '')
  A4 = @((Join-Path $parent 'elsewhere\other.xlsx'), '')
  A5 = @('..\up.xlsx', '')
  A6 = @('My File.xlsx', 'Sheet2!B5')
  A7 = @('R&D 100%.xlsx', '')
  A8 = @("Caf$([char]0xE9) [1].xlsx", '')
  A9 = @('https://example.com/a b', '')
  A10 = @('https://example.com/a%20b', '')
  A11 = @('x%41 {^}.xlsx', '')
}
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$res = [ordered]@{ folder = '<out>'; before = [ordered]@{}; after = [ordered]@{} }
$hide = { param($s) ([string]$s).Replace($Out, '<out>').Replace($parent, '<parent>') }
try {
  $wb = $x.Workbooks.Add()
  $wb.SaveAs($book, 51)
  $s = $wb.Worksheets.Item(1)
  foreach ($k in $cases.Keys) {
    $a, $sub = $cases[$k]
    if ($sub) { [void]$s.Hyperlinks.Add($s.Range($k), $a, $sub) } else { [void]$s.Hyperlinks.Add($s.Range($k), $a) }
  }
  $read = { param($label)
    foreach ($k in $cases.Keys) {
      $h = $s.Range($k).Hyperlinks.Item(1)
      $res[$label][$k] = [ordered]@{ given = (& $hide $cases[$k][0]); address = (& $hide $h.Address); subAddress = [string]$h.SubAddress; text = (& $hide $s.Range($k).Value2) }
    }
  }
  & $read 'before'
  $wb.Save()
  & $read 'after'
  $wb.Close($false)
} finally {
  $x.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($book)
try {
  $part = { param($n) $e = $zip.GetEntry($n); $r = New-Object System.IO.StreamReader($e.Open()); try { $r.ReadToEnd() } finally { $r.Close() } }
  $sheet = & $part 'xl/worksheets/sheet1.xml'
  $res['hyperlinks'] = & $hide ([regex]::Match($sheet, '<hyperlinks>.*</hyperlinks>').Value)
  $res['rels'] = & $hide (& $part 'xl/worksheets/_rels/sheet1.xml.rels')
} finally { $zip.Dispose() }
$res | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 (Join-Path $Out 'filelinks.json')
