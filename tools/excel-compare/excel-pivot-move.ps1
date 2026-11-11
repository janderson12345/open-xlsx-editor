# A PivotTable's fields moved after it is made, as the field list moves them, for pivot-move.test.js: for each move
# (one case of pivot-cases.json's layout to another's) a new workbook in an invisible Excel, the pivot built at Pivot!A3
# with the first layout as excel-pivot.ps1 builds it, then its fields taken off and put on as the second has them; its
# cells read, the workbook saved to a folder outside the repository and the three pivot parts read back, the cache's
# refreshedBy (the user's name) taken out. No keys sent, no Excel setting changed.
# usage: pwsh tools/excel-compare/excel-pivot-move.ps1 <pivot-cases.json> <out.json> <folder for the saved workbooks>
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
trap { Write-Host "LINE $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; break }
Add-Type -AssemblyName System.IO.Compression.FileSystem
$spec = Get-Content -Raw $Cases | ConvertFrom-Json
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$CAPTION = @{ sum = 'Sum'; count = 'Count'; average = 'Average'; max = 'Max'; min = 'Min' }
$FN = @{ sum = -4157; count = -4112; average = -4106; max = -4136; min = -4139 }
$MOVES = @(@(0, 1), @(1, 0), @(0, 3), @(3, 0), @(0, 6), @(6, 0), @(4, 0), @(2, 7), @(7, 2), @(5, 2))
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
function Read-Pivot($pt) {
  $r = $pt.TableRange2
  $cells = @()
  for ($i = 1; $i -le $r.Rows.Count; $i++) { for ($j = 1; $j -le $r.Columns.Count; $j++) {
    $c = $r.Cells.Item($i, $j)
    if ($null -ne $c.Value2 -and "$($c.Value2)" -ne '') { $cells += , @(($c.Row - 1), ($c.Column - 1), $c.Value2) }
  } }
  return @{ ref = $r.Address($false, $false); cells = $cells }
}
function Lay-Out($pt, $case) {
  $pos = 1; foreach ($f in $case.rows) { $p = $pt.PivotFields($f); $p.Orientation = 1; $p.Position = $pos; $pos++ }
  $pos = 1; foreach ($f in $case.cols) { $p = $pt.PivotFields($f); $p.Orientation = 2; $p.Position = $pos; $pos++ }
  foreach ($v in $case.values) { $df = $pt.AddDataField($pt.PivotFields($v.field)); $df.Function = $FN[$v.fn] }
}
$results = @()
try {
  $k = 0
  foreach ($m in $MOVES) {
    $k++
    $from = $spec.cases[$m[0]]; $to = $spec.cases[$m[1]]
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1); $ws.Name = 'Data'
    $grid = New-Object 'object[,]' $spec.data.Count, 5
    for ($i = 0; $i -lt $spec.data.Count; $i++) { $row = @($spec.data[$i]); for ($j = 0; $j -lt 5; $j++) { $grid[$i, $j] = $row[$j] } }
    $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($spec.data.Count, 5)).Value2 = $grid
    $pws = $wb.Worksheets.Add(); $pws.Name = 'Pivot'
    $pc = $wb.PivotCaches().Create(1, "Data!R1C1:R$($spec.data.Count)C5")
    $pt = $pc.CreatePivotTable($pws.Range('A3'), 'Pivot1')
    Lay-Out $pt $from
    # (Everything taken off: the values first, then the row and column fields; then the second layout put on.)
    foreach ($v in $from.values) { $pt.PivotFields($CAPTION[$v.fn] + ' of ' + $v.field).Orientation = 0 }
    foreach ($f in @($from.rows) + @($from.cols)) { if ($f) { $pt.PivotFields($f).Orientation = 0 } }
    Lay-Out $pt $to
    $cells = Read-Pivot $pt
    $file = Join-Path $Books "pivot-move-$k.xlsx"
    if (Test-Path $file) { Remove-Item $file }
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    $read = { param($n) $s = New-Object IO.StreamReader($zip.GetEntry($n).Open()); $t = $s.ReadToEnd(); $s.Close(); $t }
    $parts = @{ table = (& $read 'xl/pivotTables/pivotTable1.xml'); cache = ((& $read 'xl/pivotCache/pivotCacheDefinition1.xml') -replace ' refreshedBy="[^"]*"', ''); records = (& $read 'xl/pivotCache/pivotCacheRecords1.xml') }
    $zip.Dispose()
    $results += [pscustomobject]@{ from = $m[0]; to = $m[1]; cells = $cells; parts = $parts }
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$results | ConvertTo-Json -Depth 8 | Set-Content -Encoding utf8 $Out
"$($results.Count) moves"
