# Excel's reading of the data bars and icon sets written here (cfbars.test.js): test/out/cf/cfbars.xlsx opened read
# only in an invisible Excel; for each column of the probe, its first rule's type, data bar colour, fill and border, or
# icon set. Nothing saved, no keys sent.
# usage: pwsh tools/excel-compare/excel-cf-bars-check.ps1 <workbook> <columns.json> <out.json>
param([Parameter(Mandatory)][string]$Book, [Parameter(Mandatory)][string]$Columns, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$cols = (Get-Content -Raw $Columns | ConvertFrom-Json).columns
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$found = [ordered]@{}
try {
  $wb = $x.Workbooks.Open((Resolve-Path $Book).Path, 0, $true)
  $ws = $wb.Worksheets.Item(1)
  foreach ($p in $cols.PSObject.Properties) {
    $fc = $ws.Range("$($p.Value)1:$($p.Value)10").FormatConditions
    if ($fc.Count -lt 1) { $found[$p.Name] = $null; continue }
    $c = $fc.Item(1)
    $found[$p.Name] = $(if ($c.Type -eq 4) { [pscustomobject]@{ type = 'dataBar'; color = $c.BarColor.Color; solid = ($c.BarFillType -eq 0); border = $c.BarBorder.Type; min = $c.MinPoint.Type; max = $c.MaxPoint.Type } }
      elseif ($c.Type -eq 6) { [pscustomobject]@{ type = 'iconSet'; set = $c.IconSet.ID; count = $c.IconSet.Count } } else { [pscustomobject]@{ type = $c.Type } })
  }
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
[pscustomobject]@{ source = 'tools/excel-compare/excel-cf-bars-check.ps1 on test/out/cf/cfbars.xlsx (cfbars.test.js)'; rules = $found } | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 $Out
'done'
