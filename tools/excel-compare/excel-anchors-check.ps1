# Excel opens each workbook anchors-ours.mjs saved and reads its chart, rectangle and notes as excel-anchors.ps1 reads
# them, for comparison with what Excel did itself (test/anchors-excel.json). Invisible, read-only, closed unsaved.
# usage: pwsh tools/excel-compare/excel-anchors-check.ps1 <folder> <out.json>
param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
function Box($o) { return @{ from = $o.TopLeftCell.Address($false, $false); to = $o.BottomRightCell.Address($false, $false); width = $o.Width; height = $o.Height } }
$results = @()
try {
  Get-ChildItem (Join-Path (Resolve-Path $Folder).Path 'case-*.xlsx') | Sort-Object { [int]($_.BaseName -replace '\D', '') } | ForEach-Object {
    $wb = $x.Workbooks.Open($_.FullName, 0, $true)
    $ws = $wb.Worksheets.Item(1)
    $charts = @(); foreach ($c in $ws.ChartObjects()) { $charts += , (Box $c) }
    $shapes = @(); foreach ($p in $ws.Shapes) { if ($p.Type -eq 1) { $shapes += , (Box $p) } }
    $notes = @(); foreach ($c in $ws.Comments) { $notes += $c.Parent.Address($false, $false) }
    $table = if ($ws.ListObjects.Count) { $ws.ListObjects.Item(1).Range.Address($false, $false) } else { $null }
    $columns = if ($ws.ListObjects.Count) { @($ws.ListObjects.Item(1).ListColumns | ForEach-Object { $_.Name }) } else { @() }
    $hit = $ws.UsedRange.Find('SUM(', [Type]::Missing, -4123, 2)
    $formula = if ($hit) { "$($hit.Address($false, $false)) $($hit.Formula) = $($hit.Text)" } else { $null }
    $results += [pscustomobject]@{ file = $_.Name; charts = $charts; shapes = $shapes; notes = $notes; table = $table; columns = $columns; formula = $formula }
    $wb.Close($false)
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$results | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 $Out
"$($results.Count) opened"
