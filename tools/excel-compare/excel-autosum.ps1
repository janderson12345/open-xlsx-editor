# Has Excel's AutoSum button (and its Average, Count Numbers, Max and Min) work on a small sheet and says what it
# wrote (cases: [{ name, setup: { cells, formats: [{ range, numFmt }] }, select, fn, read }]). The ribbon's command is
# run with CommandBars.ExecuteMso, which needs an Excel window: this attaches to the Excel already open (Windows
# PowerShell, for GetActiveObject) and works in a new workbook of its own, closed without saving; no key is sent and
# no other workbook is touched. Only selections of several cells: on one cell AutoSum leaves the cell being edited,
# waiting for Enter, and Excel then answers nothing over COM until someone presses a key.
# usage: powershell.exe -NoProfile -File tools/excel-compare/excel-autosum.ps1 tools/excel-compare/autosum-cases.json test/autosum-excel.json
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$list = Get-Content $Cases -Raw | ConvertFrom-Json
$ids = @{ sum = 'AutoSum'; average = 'AutoSumAverage'; count = 'AutoSumCount'; max = 'AutoSumMax'; min = 'AutoSumMin' }
$x = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
if (-not $x.Ready) { throw 'Excel is busy (a cell being edited?)' }
$answers = New-Object System.Collections.Generic.List[object]
function Props($o) { if ($null -eq $o) { @() } else { $o.PSObject.Properties } }
try {
  foreach ($case in $list) {
    $wb = $x.Workbooks.Add()
    try {
      $ws = $wb.Worksheets.Item(1)
      $s = $case.setup
      foreach ($p in Props $s.cells) { $ws.Range($p.Name).Formula = [string]$p.Value }
      foreach ($f in @($s.formats)) { if ($f -and $f.numFmt) { $ws.Range($f.range).NumberFormat = [string]$f.numFmt } }
      $sel = $ws.Range($case.select)
      if ($sel.Cells.Count -lt 2) { throw "one cell selected in '$($case.name)': AutoSum would wait for Enter" }
      [void]$ws.Activate()
      [void]$sel.Select()
      $x.CommandBars.ExecuteMso($ids[[string]$case.fn])
      $a = [ordered]@{ name = [string]$case.name; selection = [string]$x.Selection.Address($false, $false) }
      $cells = [ordered]@{}
      foreach ($c in $ws.Range($case.read).Cells) {
        $one = [ordered]@{}
        if ($c.HasFormula) { $one.f = [string]$c.Formula }
        $v = $c.Value2
        if ($null -ne $v) { $one.v = if ($v -is [double] -or $v -is [bool]) { $v } elseif ($v -is [int]) { [double]$v } else { "$v" } }
        if ($c.NumberFormat -ne 'General') { $one.fmt = [string]$c.NumberFormat }
        if ($one.Count) { $cells[$c.Address($false, $false)] = $one }
      }
      $a.cells = $cells
      $answers.Add($a)
      # (Written after each case: should Excel stop answering, the cases before are kept.)
      [IO.File]::WriteAllText($Out, ([ordered]@{ excel = 'Excel.Application (the open one): CommandBars.ExecuteMso AutoSum'; cases = $answers } | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding $false))
    } finally { $wb.Close($false) }
  }
} finally {
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
$json = [ordered]@{ excel = 'Excel.Application (the open one): CommandBars.ExecuteMso AutoSum'; cases = $answers } | ConvertTo-Json -Depth 6
[IO.File]::WriteAllText($Out, $json, (New-Object Text.UTF8Encoding $false))
"$($answers.Count) cases"
