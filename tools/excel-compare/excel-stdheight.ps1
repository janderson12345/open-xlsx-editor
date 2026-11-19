# Excel's standard row height and row heights for stdheight-book.mjs's workbooks: Worksheet.StandardHeight and
# Range.RowHeight of rows 1-13 as Excel opens each, written as test/stdheight-excel.json
# { excel, books: { name: { std, heights: [pt, ...] } } } (row 1 first; row 13 has no element).
# An invisible instance, read-only, macros disabled for it only, no keys sent.
# usage: node tools/excel-compare/stdheight-book.mjs <dir>; pwsh tools/excel-compare/excel-stdheight.ps1 <dir> <out.json>
param([Parameter(Mandatory)][string]$Dir, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$Dir = (Resolve-Path $Dir).Path
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
try {
  $books = [ordered]@{}
  foreach ($f in Get-ChildItem -Path $Dir -Filter '*.xlsx' | Sort-Object Name) {
    $wb = $x.Workbooks.Open($f.FullName, 0, $true)
    $ws = $wb.Worksheets.Item(1)
    $heights = foreach ($r in 1..13) { [double]$ws.Rows.Item($r).RowHeight }
    $books[$f.BaseName] = [ordered]@{ std = [double]$ws.StandardHeight; heights = @($heights) }
    $wb.Close($false)
    Write-Output ("{0}: std {1} rows {2}" -f $f.BaseName, $books[$f.BaseName].std, ($heights -join ' '))
  }
  # (UTF-8 without a byte order mark, which Windows PowerShell's Set-Content would add.)
  $json = [ordered]@{ excel = $x.Version; books = $books } | ConvertTo-Json -Compress -Depth 4
  [IO.File]::WriteAllText($Out, $json + "`n", (New-Object Text.UTF8Encoding $false))
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
