# Pivot tables as Excel builds and refreshes them: for each case of pivot-cases.json, a new workbook in an invisible
# Excel with the data on Data!A1, a pivot table of it on a sheet Pivot at A3 (Excel's own place for a new sheet's),
# its fields put in rows, columns and values; the pivot's cells read (TableRange2: every value and its text); the
# workbook saved (to a folder outside the repository: Excel writes its user's name in the cache) so the parts Excel
# wrote can be read; then the source changed, the pivot refreshed and its cells read again. Closed without saving
# again under another name (the parts a refresh rewrites). No keys sent, no Excel setting changed.
# usage: pwsh tools/excel-compare/excel-pivot.ps1 <pivot-cases.json> <out.json> <folder for the saved workbooks>
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
trap { Write-Host "LINE $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; break }
$spec = Get-Content -Raw $Cases | ConvertFrom-Json
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$FN = @{ sum = -4157; count = -4112; average = -4106; max = -4136; min = -4139 }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
function Read-Pivot($pt) {
  $r = $pt.TableRange2
  $cells = @()
  for ($i = 1; $i -le $r.Rows.Count; $i++) {
    for ($j = 1; $j -le $r.Columns.Count; $j++) {
      $c = $r.Cells.Item($i, $j)
      if ($null -ne $c.Value2 -and "$($c.Value2)" -ne '') { $cells += , @(($c.Row - 1), ($c.Column - 1), $c.Value2, $c.Text) }
    }
  }
  return @{ ref = $r.Address($false, $false); cells = $cells }
}
$results = @()
try {
  $k = 0
  foreach ($case in $spec.cases) {
    $k++
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1); $ws.Name = 'Data'
    # (The table as one block: a two-dimensional array into the range.)
    $grid = New-Object 'object[,]' $spec.data.Count, 5
    for ($i = 0; $i -lt $spec.data.Count; $i++) { $row = @($spec.data[$i]); for ($j = 0; $j -lt 5; $j++) { $grid[$i, $j] = $row[$j] } }
    $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($spec.data.Count, 5)).Value2 = $grid
    $last = $spec.data.Count
    $pws = $wb.Worksheets.Add(); $pws.Name = 'Pivot'
    $pc = $wb.PivotCaches().Create(1, "Data!R1C1:R$($last)C5")
    $pt = $pc.CreatePivotTable($pws.Range('A3'), 'Pivot1')
    $pos = 1; foreach ($f in $case.rows) { $p = $pt.PivotFields($f); $p.Orientation = 1; $p.Position = $pos; $pos++ }
    $pos = 1; foreach ($f in $case.cols) { $p = $pt.PivotFields($f); $p.Orientation = 2; $p.Position = $pos; $pos++ }
    foreach ($v in $case.values) { $df = $pt.AddDataField($pt.PivotFields($v.field)); $df.Function = $FN[$v.fn] }
    $before = Read-Pivot $pt
    $file = Join-Path $Books ("pivot-case-$k.xlsx")
    if (Test-Path $file) { Remove-Item $file }
    $wb.SaveAs($file, 51)
    foreach ($ch in $spec.change) { $ws.Cells.Item($ch.r + 1, $ch.c + 1).Value2 = $ch.value }
    [void]$pt.RefreshTable()
    $after = Read-Pivot $pt
    # (Saved again after the refresh, to compare the parts a refresh rewrites.)
    $fileAfter = Join-Path $Books ("pivot-case-$k-after.xlsx")
    if (Test-Path $fileAfter) { Remove-Item $fileAfter }
    $wb.SaveAs($fileAfter, 51)
    $results += [pscustomobject]@{ name = $case.name; file = (Split-Path $file -Leaf); before = $before; after = $after }
    $wb.Close($false)
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$results | ConvertTo-Json -Depth 8 | Set-Content -Encoding utf8 $Out
"$($results.Count) cases"
