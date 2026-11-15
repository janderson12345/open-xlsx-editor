# Which texts typed into a Function Arguments box Excel reads as a formula, for fnquote.test.js: the box puts quotes
# round text that doesn't read as one (Microsoft's description of the box; the box itself can't be driven over COM
# without sending keys). In an invisible Excel, a new workbook with 1, 2, 3 in A1:A3, a second sheet "My Data" and a
# defined name Rate (=0.05); each text set as D1's Formula with "=" in front, and whether Excel took it and D1's value
# as text recorded. No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-fnquote.ps1 <out.json>
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$TEXTS = @('>1', '<>', '>=10', '=5', 'abc', 'Abc', 'hello world', 'yes', 'No', 'A1', 'a1', 'A1:A3', '$A$1', 'A:A', '1:1',
  "'My Data'!A1", 'Sheet2!A1', '1.5', '-2', '1,000', '$5', '5%', '1/2', '1e3', 'TRUE', 'false', '"x"', '"', 'x"y',
  'Rate', 'rate*2', 'SUM(A1:A3)', 'sum(1,2)', 'SUMM(1)', 'A1+1', '{1,2,3}', '#N/A', '#DIV/0!', '*', '?a', 'a*', 'Jan',
  'Q1', 'XFD1', 'XFE1', 'R1C1', 'x_y', 'a.b', 'a-b', '(1)', '1 2', 'A1 A2', ' ', 'apple pie', 'N/A', '@', 'a b c', 'ABC123')
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$rows = @()
try {
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $s2 = $wb.Worksheets.Add([Type]::Missing, $ws); $s2.Name = 'My Data'
  [void]$wb.Names.Add('Rate', '=0.05')
  $ws.Range('A1').Value2 = 1; $ws.Range('A2').Value2 = 2; $ws.Range('A3').Value2 = 3
  foreach ($t in $TEXTS) {
    $d = $ws.Range('D1')
    [void]$d.Clear()
    $ok = $true
    try { $d.Formula = "=$t" } catch { $ok = $false }
    $rows += [pscustomobject]@{ typed = $t; ok = $ok; formula = $(if ($ok) { $d.Formula } else { $null }); text = $(if ($ok) { $d.Text } else { $null }) }
  }
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$rows | ConvertTo-Json -Depth 3 | Set-Content -Encoding utf8 $Out
'done'
