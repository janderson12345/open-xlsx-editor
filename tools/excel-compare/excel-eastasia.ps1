# How Excel shows the East Asian built-in number formats (ids 27-36, 50-58). Part 1: a workbook giving only the ids
# (eastasia-book.mjs) opened and every cell's Range.Text and Range.NumberFormat read, which is what this machine's
# Excel (en-US) makes of them. Part 2: each locale's code from ECMA-376's table (src/ooxml/eastasia.js, passed in as
# JSON) set on a new sheet with each value, Range.Text read: how Excel draws eras, ROC years and AM/PM words.
# An invisible instance, macros disabled for it only, workbooks closed without saving, no keys sent.
# usage: pwsh tools/excel-compare/excel-eastasia.ps1 <ids.xlsx> <codes.json> <out.json>
param([Parameter(Mandatory)][string]$Book, [Parameter(Mandatory)][string]$Codes, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$spec = Get-Content -Raw -Encoding utf8 $Codes | ConvertFrom-Json
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$ids = New-Object System.Collections.Generic.List[object]
$codesOut = New-Object System.Collections.Generic.List[object]
try {
  $wb = $x.Workbooks.Open((Resolve-Path $Book).Path, 0, $true)
  $ws = $wb.Worksheets.Item(1)
  for ($r = 1; $r -le $spec.ids.Count; $r++) {
    for ($c = 1; $c -le $spec.values.Count; $c++) {
      $cell = $ws.Cells.Item($r, $c)
      $ids.Add([ordered]@{ id = [int]$spec.ids[$r - 1]; value = [double]$spec.values[$c - 1]; code = [string]$cell.NumberFormat; local = [string]$cell.NumberFormatLocal; text = [string]$cell.Text })
    }
  }
  $wb.Close($false)
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $ws.Columns.Item(1).ColumnWidth = 80
  $row = 0
  foreach ($k in $spec.codes) {
    foreach ($v in $spec.values) {
      $row++
      $cell = $ws.Cells.Item($row, 1)
      $cell.NumberFormat = [string]$k.code
      $cell.Value2 = [double]$v
      $codesOut.Add([ordered]@{ code = [string]$k.code; value = [double]$v; text = [string]$cell.Text })
    }
  }
  $wb.Close($false)
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
$json = [ordered]@{ excel = 'Excel.Application (en-US Windows), Range.Text'; ids = $ids; codes = $codesOut } | ConvertTo-Json -Depth 4
[IO.File]::WriteAllText($Out, $json, (New-Object Text.UTF8Encoding $false))
"$($ids.Count) id cells, $($codesOut.Count) code cells"
