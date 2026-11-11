# Excel opens each workbook pivot-lines-ours.mjs saved, reads its pivot table's range and source (or none), refreshes
# it and reads the range again: for comparison with what Excel did itself (test/pivot-lines-excel.json). Invisible,
# read-only, closed without saving.
# usage: pwsh tools/excel-compare/excel-pivot-lines-check.ps1 <folder> <out.json>
param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$results = @()
try {
  Get-ChildItem (Join-Path (Resolve-Path $Folder).Path 'step-*.xlsx') | Sort-Object { [int]($_.BaseName -replace '\D', '') } | ForEach-Object {
    $wb = $x.Workbooks.Open($_.FullName, 0, $true)
    $ws = $wb.Worksheets.Item('Pivot')
    $pivot = $null; $source = $null; $refreshed = $null
    if ($ws.PivotTables().Count) {
      $pt = $ws.PivotTables(1)
      $pivot = $pt.TableRange2.Address($false, $false); $source = $pt.SourceData
      [void]$pt.RefreshTable(); $refreshed = $pt.TableRange2.Address($false, $false)
    }
    $results += [pscustomobject]@{ file = $_.Name; pivot = $pivot; source = $source; refreshed = $refreshed }
    $wb.Close($false)
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$results | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 $Out
"$($results.Count) opened"
