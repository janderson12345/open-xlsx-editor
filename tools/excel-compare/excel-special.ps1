# What Excel's Go To Special finds (Home > Find & Select > Go To Special), for special-cases.mjs: test/fixtures/special.xlsx
# opened read-only in an invisible Excel; for each case its ranges selected and its active cell made active, then the
# choice asked as the box asks it (Selection.SpecialCells, ActiveCell.CurrentRegion and CurrentArray, Selection's
# RowDifferences and ColumnDifferences, Precedents and Dependents), the result selected, and its address and the active
# cell then read. A choice that finds nothing gives Excel's "No cells were found." (null). No keys sent, nothing saved.
# usage: pwsh tools/excel-compare/excel-special.ps1 <cases.json> <special.xlsx> <out.json>
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Book, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
trap { Write-Host "LINE $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; break }
$spec = Get-Content -Raw $Cases | ConvertFrom-Json
$Book = (Resolve-Path $Book).Path
$TYPES = @{ constants = 2; formulas = -4123; blanks = 4; notes = -4144; lastCell = 11; visible = 12; cfAll = -4172; cfSame = -4173; dvAll = -4174; dvSame = -4175 }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$results = @()
try {
  $wb = $x.Workbooks.Open($Book, 0, $true)
  $ws = $wb.Worksheets.Item(1)
  $ws.Activate()
  foreach ($case in $spec.cases) {
    $ws.Range($case.select).Select()
    $ws.Range($case.active).Activate()
    $found = $null
    $err = $null
    try {
      $sel = $x.Selection
      $act = $x.ActiveCell
      switch ($case.kind) {
        'region' { $found = $act.CurrentRegion }
        'array' { $found = $act.CurrentArray }
        'rowDiff' { $found = $sel.RowDifferences($act) }
        'colDiff' { $found = $sel.ColumnDifferences($act) }
        'precedents' { if ($case.levels -eq 'all') { $found = $sel.Precedents } else { $found = $sel.DirectPrecedents } }
        'dependents' { if ($case.levels -eq 'all') { $found = $sel.Dependents } else { $found = $sel.DirectDependents } }
        { $_ -in 'constants', 'formulas' } { $found = $sel.SpecialCells($TYPES[$case.kind], [int]$case.types) }
        default { $found = $sel.SpecialCells($TYPES[$case.kind]) }
      }
    } catch { $err = $_.Exception.Message }
    $address = $null; $active = $null
    if ($null -ne $found) {
      $address = $found.Address($false, $false)
      $found.Select()
      $active = $x.ActiveCell.Address($false, $false)
    }
    $results += [pscustomobject]@{ name = $case.name; address = $address; active = $active; error = $err }
  }
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$results | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 $Out
"$($results.Count) cases"
