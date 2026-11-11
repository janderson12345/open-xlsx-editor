# Excel's Advanced Filter, for advfilter.test.js: in an invisible Excel, a new workbook of a made-up list (A1:C9, with a
# repeated row) and a criteria range; each case Range.AdvancedFilter (in place or copied to I1, with or without a
# criteria range, unique records only or not); the rows hidden, the cells copied, and from the workbook Excel then
# saves, its defined names and whether the sheet has an autoFilter. No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-advfilter.ps1 <out.json> <folder for the saved workbooks>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$LIST = @(@('Name', 'Qty', 'Team'), @('apple', '1', 'red'), @('pear', '2', 'blue'), @('Apple', '3', 'red'), @('fig', '1', 'green'), @('plum', '2.5', 'red'),
  @('kiwi', '10', 'blue'), @('pear', '2', 'blue'), @('avocado', '7', 'red'))
# Each case: criteria cells (F1.., by rows; '' for none), action (1 in place, 2 copy to I1), unique.
$CASES = [ordered]@{
  'rows or, columns and' = @(@('Team|Qty', 'red|>2', 'blue|'), 1, $false)
  'text begins with' = @(@('Name', 'a'), 1, $false)
  'exact text' = @(@('Name', '="=pear"'), 1, $false)
  'two conditions on a column' = @(@('Qty|Qty', '>=2|<7'), 1, $false)
  'unique in place' = @($null, 1, $true)
  'copied to I1' = @(@('Team', 'blue'), 2, $false)
  'copied, unique' = @(@('Team', 'blue'), 2, $true)
}
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$answers = New-Object System.Collections.Generic.List[object]
try {
  $i = 0
  foreach ($name in $CASES.Keys) {
    $i++
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    for ($r = 0; $r -lt $LIST.Count; $r++) { for ($c = 0; $c -lt 3; $c++) { $ws.Range("$([char](65 + $c))$($r + 1)").Formula = $LIST[$r][$c] } }
    $spec = $CASES[$name]
    $crit = $null
    if ($spec[0]) {
      # (Each row of the criteria a string, its cells split at |.)
      $rows = @($spec[0] | ForEach-Object { ,([string]$_ -split '\|') })
      for ($r = 0; $r -lt $rows.Count; $r++) { for ($c = 0; $c -lt $rows[$r].Count; $c++) { if ($rows[$r][$c] -ne '') { $ws.Range("$([char](70 + $c))$($r + 1)").Formula = $rows[$r][$c] } } }
      $crit = $ws.Range("F1:$([char](69 + $rows[0].Count))$($rows.Count)")
    }
    $source = $ws.Range('A1:C9')
    $err = $null
    try {
      if ($spec[1] -eq 1) { if ($crit) { [void]$source.AdvancedFilter(1, $crit, [Type]::Missing, $spec[2]) } else { [void]$source.AdvancedFilter(1, [Type]::Missing, [Type]::Missing, $spec[2]) } }
      else { [void]$source.AdvancedFilter(2, $crit, $ws.Range('I1'), $spec[2]) }
    } catch { $err = [string]$_.Exception.Message }
    $hidden = New-Object System.Collections.Generic.List[int]
    for ($r = 2; $r -le 9; $r++) { if ($ws.Rows.Item($r).Hidden) { $hidden.Add($r) } }
    $copied = New-Object System.Collections.Generic.List[string]
    for ($r = 1; $r -le 9; $r++) { $line = @(); for ($c = 9; $c -le 11; $c++) { $line += [string]$ws.Cells.Item($r, $c).Text }; if (($line -join '') -ne '') { $copied.Add($line -join '|') } }
    $file = Join-Path $Books "advfilter-$i.xlsx"
    if (Test-Path $file) { Remove-Item $file }
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    $read = { param($n) $s = New-Object IO.StreamReader($zip.GetEntry($n).Open()); $t = $s.ReadToEnd(); $s.Close(); $t }
    $book = & $read 'xl/workbook.xml'
    $sheet = & $read 'xl/worksheets/sheet1.xml'
    $zip.Dispose()
    $answers.Add([ordered]@{ name = $name; error = $err; criteria = $spec[0]; action = $spec[1]; unique = $spec[2]; hidden = $hidden; copied = $copied
      names = @([regex]::Matches($book, '<definedName\b[^>]*>[^<]*</definedName>') | ForEach-Object { $_.Value }); autoFilter = [regex]::Match($sheet, '<autoFilter[^>]*?(/>|>.*?</autoFilter>)').Value })
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
[ordered]@{ source = 'tools/excel-compare/excel-advfilter.ps1'; cases = $answers } | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 $Out
"$($answers.Count) cases"
