# Excel opens each workbook dynamic-ours.mjs saved (dynamic array formulas typed in the editor) and reads, as
# excel-dynamic.ps1 reads its own: D1's spill range and the values, texts and formulas of D1:H8 on the first sheet.
# Invisible, read-only, closed without saving. Compare the output with test/dynamic-excel.json.
# usage: pwsh tools/excel-compare/excel-dynamic-check.ps1 <folder> <out.json>
param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
trap { Write-Host "LINE $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; break }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$results = @()
try {
  Get-ChildItem (Join-Path (Resolve-Path $Folder).Path 'ours-*.xlsx') | Sort-Object { [int]($_.BaseName -replace '\D', '') } | ForEach-Object {
    $wb = $x.Workbooks.Open($_.FullName, 0, $true)
    $ws = $wb.Worksheets.Item(1)
    $d1 = $ws.Range('D1')
    $spill = $null
    if ($d1.HasSpill) { $spill = $d1.SpillingToRange.Address($false, $false) }
    $cells = @()
    for ($r = 1; $r -le 8; $r++) { for ($c = 4; $c -le 8; $c++) {
      $cell = $ws.Cells.Item($r, $c)
      if ($null -ne $cell.Value2 -or $cell.HasFormula) { $cells += [pscustomobject]@{ ref = $cell.Address($false, $false); value = $cell.Value2; text = $cell.Text; formula2 = $cell.Formula2 } }
    } }
    $results += [pscustomobject]@{ file = $_.Name; spill = $spill; cells = $cells }
    $wb.Close($false)
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$results | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 $Out
"$($results.Count) opened"
