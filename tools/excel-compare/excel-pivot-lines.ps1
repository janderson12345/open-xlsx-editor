# Where a PivotTable goes when Excel inserts or deletes rows, columns or cells: on its own sheet (Pivot, the pivot at
# C5, rows Region, sum of Units, A1 holding a value beside it) and on its source's (Src, the data of pivot-cases.json
# on A1:E19). For each step a new workbook in an invisible Excel, the step done, and read back: the pivot's range
# (TableRange2) and its source (SourceData), or Excel's refusal. Nothing saved; no keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-pivot-lines.ps1 <pivot-cases.json> <out.json>
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$spec = Get-Content -Raw $Cases | ConvertFrom-Json
$steps = @(
  @('Pivot', 'insert', '1:1'), @('Pivot', 'insert', '6:7'), @('Pivot', 'insert', '20:20'), @('Pivot', 'delete', '1:1'), @('Pivot', 'delete', '6:7'),
  @('Pivot', 'delete', '1:12'), @('Pivot', 'delete', '1:4'), @('Pivot', 'insert', 'A:A'), @('Pivot', 'insert', 'D:D'), @('Pivot', 'delete', 'A:A'),
  @('Pivot', 'delete', 'D:D'), @('Pivot', 'delete', 'A:F'), @('Pivot', 'insert', 'F:F'),
  @('Pivot', 'insert', 'C6', 'down'), @('Pivot', 'insert', 'A2:H2', 'down'), @('Pivot', 'insert', 'A6', 'right'), @('Pivot', 'insert', 'F6', 'right'), @('Pivot', 'delete', 'C12:D12', 'up'),
  @('Src', 'insert', '5:5'), @('Src', 'insert', '1:1'), @('Src', 'delete', '5:6'), @('Src', 'insert', '20:20'), @('Src', 'delete', '19:19'),
  @('Src', 'insert', 'C:C'), @('Src', 'delete', 'C:C'), @('Src', 'insert', 'A:A'), @('Src', 'delete', '1:19'), @('Src', 'insert', 'B5', 'down')
)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$results = @()
try {
  foreach ($s in $steps) {
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1); $ws.Name = 'Src'
    $grid = New-Object 'object[,]' $spec.data.Count, 5
    for ($i = 0; $i -lt $spec.data.Count; $i++) { $row = @($spec.data[$i]); for ($j = 0; $j -lt 5; $j++) { $grid[$i, $j] = $row[$j] } }
    $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($spec.data.Count, 5)).Value2 = $grid
    $pws = $wb.Worksheets.Add(); $pws.Name = 'Pivot'
    $pws.Range('A1').Value2 = 'beside'
    $pc = $wb.PivotCaches().Create(1, "Src!R1C1:R$($spec.data.Count)C5")
    $pt = $pc.CreatePivotTable($pws.Range('C5'), 'Pivot1')
    $pt.PivotFields('Region').Orientation = 1
    [void]$pt.AddDataField($pt.PivotFields('Units'))
    $refused = $null
    try {
      $r = $wb.Worksheets.Item($s[0]).Range($s[2])
      $shift = if ($s.Count -gt 3) { @{ down = -4121; right = -4161; up = -4162; left = -4159 }[$s[3]] } else { $null }
      if ($s[1] -eq 'insert') { if ($null -ne $shift) { [void]$r.Insert($shift) } else { [void]$r.Insert() } }
      else { if ($null -ne $shift) { [void]$r.Delete($shift) } else { [void]$r.Delete() } }
    } catch { $refused = $_.Exception.Message }
    $n = $pws.PivotTables().Count
    $range = if ($n) { $pws.PivotTables(1).TableRange2.Address($false, $false) } else { $null }
    $source = if ($n) { $pws.PivotTables(1).SourceData } else { $null }
    $results += [pscustomobject]@{ sheet = $s[0]; do = $s[1]; range = $s[2]; shift = $(if ($s.Count -gt 3) { $s[3] } else { $null }); refused = $refused; pivot = $range; source = $source; beside = $($hit = $pws.UsedRange.Find('beside'); if ($hit) { $hit.Address($false, $false) } else { $null }) }
    $wb.Close($false)
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$results | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 $Out
"$($results.Count) steps"
