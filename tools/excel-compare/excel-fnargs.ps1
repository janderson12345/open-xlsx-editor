# What Excel keeps of a formula as the Function Arguments box would build it, for fnargs.test.js: in an invisible
# Excel, a new workbook with 1, 2, 3 in A1:A3 and "abc" in B1; each formula (arguments left empty at the end, in the
# middle, or filled) set as D1's Range.Formula, and whether Excel took it, the formula it kept and D1's value as
# text recorded. A formula Excel refuses (too few arguments) raises an error over COM. No keys sent, no setting
# changed.
# usage: pwsh tools/excel-compare/excel-fnargs.ps1 <out.json>
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$FORMULAS = @(
  '=SUM()', '=SUM(A1:A3)', '=SUM(A1:A3,)', '=SUM(A1:A3,,5)', '=SUM(,5)', '=SUM(A1,A2,A3)',
  '=IF(A1>1)', '=IF(A1>1,)', '=IF(A1>1,,"no")', '=IF(A1>1,"yes")', '=IF(A1>1,"yes",)', '=IF(,"yes")', '=IF()',
  '=VLOOKUP(2,A1:A3)', '=VLOOKUP(2,A1:A3,1)', '=VLOOKUP(2,A1:A3,1,)', '=VLOOKUP(,A1:A3,1)', '=VLOOKUP(2,,1)', '=VLOOKUP(,,)',
  '=ROUND(1.25)', '=ROUND(1.25,)', '=ROUND(1.25,1)', '=ROUND(,1)',
  '=LEFT(B1)', '=LEFT(B1,)', '=LEFT()', '=LEFT(abc)',
  '=PMT(0.01,12,1000)', '=PMT(0.01,12,1000,,)', '=PMT(0.01,12,1000,,1)', '=PMT(0.01,12)',
  '=TODAY()', '=PI()', '=COUNTIF(A1:A3,">1")', '=COUNTIF(A1:A3,>1)', '=COUNTIF(A1:A3)',
  '=TEXTJOIN(",",TRUE,A1:A3)', '=TEXTJOIN(",",TRUE,A1:A3,)', '=TEXTJOIN(",",TRUE)', '=CONCAT(B1,,B1)',
  '=MAX(A1:A3,,)', '=AVERAGE(,)', '=AVERAGE()', '=COUNT()', '=INDEX(A1:A3,2,)', '=MID(B1,2)', '=MID(B1,2,)'
)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$rows = @()
try {
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Range('A1').Value2 = 1; $ws.Range('A2').Value2 = 2; $ws.Range('A3').Value2 = 3; $ws.Range('B1').Value2 = 'abc'
  foreach ($f in $FORMULAS) {
    $d = $ws.Range('D1')
    [void]$d.Clear()
    $ok = $true
    try { $d.Formula = $f } catch { $ok = $false }
    $rows += [pscustomobject]@{ typed = $f; ok = $ok; formula = $(if ($ok) { $d.Formula } else { $null }); text = $(if ($ok) { $d.Text } else { $null }) }
  }
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$rows | ConvertTo-Json -Depth 3 | Set-Content -Encoding utf8 $Out
'done'
