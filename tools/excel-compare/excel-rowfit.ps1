# The row heights Excel works out for rowfit-book.mjs's workbook (no customHeight anywhere): Range.RowHeight of rows
# 1-62 as Excel opens it, written as test/rowfit-excel.json { heights: [pt, ...] } (row 1 first).
# An invisible instance, read-only, macros disabled for it only, no keys sent.
# usage: node tools/excel-compare/rowfit-book.mjs <book.xlsx>; pwsh tools/excel-compare/excel-rowfit.ps1 <book.xlsx> <out.json>
param([Parameter(Mandatory)][string]$Book, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$Book = (Resolve-Path $Book).Path
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
try {
  $wb = $x.Workbooks.Open($Book, 0, $true)
  $ws = $wb.Worksheets.Item(1)
  $heights = foreach ($r in 1..62) { [double]$ws.Rows.Item($r).RowHeight }
  $wb.Close($false)
  # (UTF-8 without a byte order mark, which Windows PowerShell's Set-Content would add.)
  $json = [ordered]@{ excel = $x.Version; heights = $heights } | ConvertTo-Json -Compress
  [IO.File]::WriteAllText($Out, $json + "`n", (New-Object Text.UTF8Encoding $false))
  Write-Output "rows 1-62: $($heights -join ' ')"
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
