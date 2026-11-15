# The fewest arguments Excel takes for each function, for fnargs.js's MIN_ARGS: in an invisible Excel, a new workbook;
# for each function of the list (calc/catalog.js's names and how many arguments each names), =F(), =F(B2),
# =F(B2,B2) ... set as A1's Formula2 until Excel takes one (a formula with too few arguments raises an error over COM), and the
# value of the first one taken read, to tell a function this Excel doesn't know (#NAME?, which it takes with any
# number). Also whether it takes one more argument than the syntax names. No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-fnmin.ps1 <list.json> <out.json>
param([Parameter(Mandatory)][string]$List, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$fns = Get-Content $List -Raw | ConvertFrom-Json
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$rows = @()
$take = { param($cell, $f) try { $cell.Formula2 = $f; $true } catch { $false } }
try {
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $a = $ws.Range('A1')
  foreach ($fn in $fns) {
    $min = $null; $known = $true
    for ($n = 0; $n -le [Math]::Max($fn.named, 1) + 1; $n++) {
      # (Each argument B2: an empty one can't be written when there is only one.)
      $f = "=$($fn.name)(" + ((@('B2') * $n) -join ',') + ')'
      if (& $take $a $f) { $min = $n; $known = $a.Text -ne '#NAME?'; break }
    }
    $more = $null
    if ($fn.named -gt 0) { $more = & $take $a ("=$($fn.name)(" + ((@('B2') * ($fn.named + 1)) -join ',') + ')') }
    [void]$a.Clear()
    $rows += [pscustomobject]@{ name = $fn.name; named = $fn.named; min = $min; known = $known; takesMore = $more }
  }
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$rows | ConvertTo-Json -Depth 3 | Set-Content -Encoding utf8 $Out
'done'
