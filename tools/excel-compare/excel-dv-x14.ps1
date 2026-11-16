# Which data validation rules Excel writes in its 2010 form (x14:dataValidation in the sheet's extLst), and how: in an
# invisible Excel, one new workbook per rule (made-up values on a second sheet "Data" and a defined name), the rule
# added with Range.Validation.Add, the workbook saved to %TEMP% and the sheet's rules kept as text: the main
# dataValidations, and the extLst with its ext elements' uris in their order. No keys sent, no setting changed.
# With -Open <folder>, Excel opens the editor's own results instead (<folder>\<rule number>.xlsx, from
# test/dvx14.test.js), read-only, and reads B2's rule back over COM (type, operator, formulas), `repaired` saying
# whether it had to repair the file (a repair log, error*.xml, appeared in %TEMP%).
# usage: pwsh tools/excel-compare/excel-dv-x14.ps1 <out.json> [-Open <folder>]
param([Parameter(Mandatory)][string]$Out, [string]$Open = '')
$ErrorActionPreference = 'Stop'
if ($Open) {
  $x = New-Object -ComObject Excel.Application
  $x.Visible = $false; $x.DisplayAlerts = $false
  $results = @()
  try {
    for ($n = 1; Test-Path (Join-Path $Open "$n.xlsx"); $n++) {
      $logs = @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue).Count
      $wb = $x.Workbooks.Open((Join-Path (Resolve-Path $Open).Path "$n.xlsx"), 0, $true)
      $repaired = @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue).Count -ne $logs
      $v = $wb.Worksheets.Item(1).Range('B2').Validation
      $f2 = try { $v.Formula2 } catch { $null }
      $results += [pscustomobject]@{ file = "$n.xlsx"; repaired = $repaired; type = $v.Type; operator = $v.Operator; formula1 = $v.Formula1; formula2 = $f2; range = $wb.Worksheets.Item(1).Cells.SpecialCells(-4174).Address($false, $false) }
      $wb.Close($false)
    }
  } finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
  [IO.File]::WriteAllText($Out, ($results | ConvertTo-Json -Depth 4), (New-Object Text.UTF8Encoding $false))
  "$($results.Count) opened"
  return
}
Add-Type -AssemblyName System.IO.Compression.FileSystem
# [name, type (1 whole, 2 decimal, 3 list, 4 date, 6 text length, 7 custom), operator (1 between, 3 equal, 5 greater),
#  formula1, formula2, sparkline too]
$RULES = @(
  @('list of another sheet', 3, 1, '=Data!$A$1:$A$3', $null, $false),
  @('list of this sheet', 3, 1, '=$H$1:$H$3', $null, $false),
  @('list of a name', 3, 1, '=Choices', $null, $false),
  @('whole, max from another sheet', 1, 1, '1', '=Data!$B$1', $false),
  @('whole, min from another sheet', 1, 5, '=Data!$B$1', $null, $false),
  @('custom reading another sheet', 7, 1, '=B2>Data!B1', $null, $false),
  @('custom reading this sheet', 7, 1, '=B2>C2', $null, $false),
  @('length from another sheet', 6, 3, '=Data!$B$1', $null, $false),
  @('list of another sheet, with a sparkline', 3, 1, '=Data!$A$1:$A$3', $null, $true)
)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$results = @()
$n = 0
try {
  foreach ($rule in $RULES) {
    $n++
    $wb = $x.Workbooks.Add()
    while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $ws = $wb.Worksheets.Item(1)
    $data = $wb.Worksheets.Item(2); $data.Name = 'Data'
    $data.Range('A1').Formula = 'p'; $data.Range('A2').Formula = 'q'; $data.Range('A3').Formula = 'r'; $data.Range('B1').Value2 = [double]5
    $ws.Range('H1').Formula = 'a'; $ws.Range('H2').Formula = 'b'; $ws.Range('H3').Formula = 'c'
    [void]$wb.Names.Add('Choices', '=Data!$A$1:$A$3')
    [void]$ws.Activate()
    if ($rule[5]) { for ($c = 2; $c -le 5; $c++) { $ws.Cells.Item(10, $c).Value2 = [double]$c }; [void]$ws.Range('G10').SparklineGroups.Add(1, 'Sheet1!B10:E10') }
    $val = $ws.Range('B2:B5').Validation
    if ($null -ne $rule[4]) { $val.Add($rule[1], 1, $rule[2], $rule[3], $rule[4]) } else { $val.Add($rule[1], 1, $rule[2], $rule[3]) }
    $file = Join-Path $env:TEMP "oxe-dvx$n.xlsx"
    if (Test-Path $file) { Remove-Item $file }
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    try { $s = New-Object IO.StreamReader($zip.GetEntry('xl/worksheets/sheet1.xml').Open()); $t = $s.ReadToEnd(); $s.Close() } finally { $zip.Dispose() }
    Remove-Item $file
    $results += [pscustomobject]@{
      name = $rule[0]
      main = [regex]::Match($t, '<dataValidations.*?</dataValidations>', 'Singleline').Value -replace ' xr:uid="[^"]*"', ''
      ext = [regex]::Match($t, '<x14:dataValidations.*?</x14:dataValidations>', 'Singleline').Value -replace ' xr:uid="[^"]*"', ''
      uris = @([regex]::Matches($t, '<ext uri="([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
      root = [regex]::Match($t, '<worksheet[^>]*>').Value -replace ' xr:uid="[^"]*"', ''
      after = [regex]::Match($t, '</sheetData>(.*?)<extLst>', 'Singleline').Groups[1].Value
    }
    Write-Host "$n $($rule[0])"
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
[IO.File]::WriteAllText($Out, ($results | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding $false))
"$($results.Count) rules"
