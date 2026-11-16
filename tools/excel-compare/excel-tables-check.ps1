# Excel reads back the tables and cell styles the editor wrote: each .xlsx in a folder (test/out/tablemake, written by
# test/tableedit.test.js) opened read-only in an invisible Excel, and each table of its first sheet read through COM
# (name, range, style, the Table Style Options, the filter, its columns and their totals, the total row's values), with
# cell B10's style name. For tableedit.test.js (test/tablemake-readback-excel.json). Nothing saved, no keys sent; the
# Excel started here is quit.
# usage: pwsh tools/excel-compare/excel-tables-check.ps1 <folder> <out.json>
param([Parameter(Mandatory)][string]$In, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$In = (Resolve-Path $In).Path
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$files = New-Object System.Collections.Generic.List[object]
try {
  foreach ($f in Get-ChildItem $In -Filter *.xlsx | Sort-Object Name) {
    $entry = [ordered]@{ file = $f.Name }
    try {
      $wb = $x.Workbooks.Open($f.FullName, 0, $true)
      $ws = $wb.Worksheets.Item(1)
      $entry.tables = @(foreach ($lo in $ws.ListObjects) {
        $totals = if ($lo.ShowTotals) { @($lo.TotalsRowRange.Cells | ForEach-Object { [string]$_.Text }) } else { @() }
        [ordered]@{ name = $lo.Name; range = $lo.Range.Address($false, $false); style = [string]$lo.TableStyle.Name
          headers = $lo.ShowHeaders; totals = $lo.ShowTotals; rows = $lo.ShowTableStyleRowStripes; cols = $lo.ShowTableStyleColumnStripes
          first = $lo.ShowTableStyleFirstColumn; last = $lo.ShowTableStyleLastColumn; filter = $lo.ShowAutoFilter; buttons = $(if ($lo.ShowAutoFilter) { $lo.ShowAutoFilterDropDown } else { $false })
          columns = @($lo.ListColumns | ForEach-Object { $_.Name }); calc = @($lo.ListColumns | ForEach-Object { $_.TotalsCalculation }); totalsText = $totals }
      })
      $entry.b10 = [string]$ws.Range('B10').Style.Name
      $entry.g3 = [string]$ws.Range('G3').Formula
      $wb.Close($false)
    } catch { $entry.error = $_.Exception.Message }
    $files.Add($entry)
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
[IO.File]::WriteAllText(([IO.Path]::GetFullPath($Out, (Get-Location).Path)), (ConvertTo-Json -InputObject $files.ToArray() -Depth 6) + "`n", (New-Object Text.UTF8Encoding($false)))
