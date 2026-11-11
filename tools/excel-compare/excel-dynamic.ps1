# What Excel 365 does with dynamic array formulas, for dynamic-cases.mjs: for each case a new workbook in an invisible
# Excel, numbers in A1:A6 and words in B1:B6, the formula entered in D1 with Formula2 (as typed in Excel 365), then its
# spill range (HasSpill, SpillingToRange), the values and texts of D1:H8, and the formula as Formula and Formula2
# read; the workbook saved to a folder outside the repository and the sheet's cells in D:H, its metadata part and the
# content types read back (no user name in them). No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-dynamic.ps1 <cases.json> <out.json> <folder for the saved workbooks>
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
trap { Write-Host "LINE $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; break }
Add-Type -AssemblyName System.IO.Compression.FileSystem
$spec = Get-Content -Raw $Cases | ConvertFrom-Json
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$results = @()
try {
  $k = 0
  foreach ($case in $spec.cases) {
    $k++
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $grid = New-Object 'object[,]' 6, 2
    $nums = @(5, 3, 8, 3, 1, 8); $words = @('pear', 'apple', 'fig', 'apple', 'kiwi', 'plum')
    for ($i = 0; $i -lt 6; $i++) { $grid[$i, 0] = $nums[$i]; $grid[$i, 1] = $words[$i] }
    $ws.Range('A1:B6').Value2 = $grid
    if ($case.block) { $ws.Range($case.block.ref).Value2 = $case.block.value }
    $ws.Range('D1').Formula2 = $case.formula
    if ($case.then) { $ws.Range($case.then.ref).Formula2 = $case.then.formula }
    $d1 = $ws.Range('D1')
    $spill = $null
    if ($d1.HasSpill) { $spill = $d1.SpillingToRange.Address($false, $false) }
    $cells = @()
    for ($r = 1; $r -le 8; $r++) { for ($c = 4; $c -le 8; $c++) {
      $cell = $ws.Cells.Item($r, $c)
      if ($null -ne $cell.Value2 -or $cell.HasFormula) { $cells += [pscustomobject]@{ ref = $cell.Address($false, $false); value = $cell.Value2; text = $cell.Text; formula = $cell.Formula; formula2 = $cell.Formula2 } }
    } }
    $file = Join-Path $Books "dynamic-$k.xlsx"
    if (Test-Path $file) { Remove-Item $file }
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    $read = { param($n) $e = $zip.GetEntry($n); if (-not $e) { return $null }; $s = New-Object IO.StreamReader($e.Open()); $t = $s.ReadToEnd(); $s.Close(); $t }
    $sheet = & $read 'xl/worksheets/sheet1.xml'
    $metadata = & $read 'xl/metadata.xml'
    $types = & $read '[Content_Types].xml'
    $rels = & $read 'xl/_rels/workbook.xml.rels'
    $zip.Dispose()
    $xml = ([regex]::Matches($sheet, '<c r="[D-H]\d+"[^>]*(?:/>|>[\s\S]*?</c>)') | ForEach-Object { $_.Value }) -join ''
    $results += [pscustomobject]@{ name = $case.name; spill = $spill; cells = $cells; xml = $xml; metadata = $metadata
      metadataType = ([regex]::Match($types, '<Override PartName="/xl/metadata.xml"[^>]*/>')).Value; metadataRel = ([regex]::Match($rels, '<Relationship [^>]*metadata[^>]*/>')).Value }
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$results | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 $Out
"$($results.Count) cases"
