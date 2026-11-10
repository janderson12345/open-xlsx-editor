# Excel opens each workbook pivot-ours.mjs saved (the editor's refresh), reads its pivot table's cells, refreshes it
# itself and reads them again: the two the same means Excel takes the cache, records and layout written there as its
# own. Invisible, read-only, closed without saving.
# usage: pwsh tools/excel-compare/excel-pivot-check.ps1 <folder>
param([Parameter(Mandatory)][string]$Folder)
$ErrorActionPreference = 'Stop'
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  Get-ChildItem (Join-Path (Resolve-Path $Folder).Path 'ours-*.xlsx') | Sort-Object Name | ForEach-Object {
    $wb = $x.Workbooks.Open($_.FullName, 0, $true)
    $pt = $wb.Worksheets.Item('Pivot').PivotTables(1)
    $read = { $r = $pt.TableRange2; $s = $r.Address($false, $false) + ':'; foreach ($c in $r.Cells) { $s += "|$($c.Value2)" }; $s }
    $before = & $read
    [void]$pt.RefreshTable()
    $after = & $read
    "$($_.Name): opened, $($pt.TableRange2.Address($false, $false)), Excel's own refresh changed nothing: $($before -eq $after)"
    $wb.Close($false)
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
