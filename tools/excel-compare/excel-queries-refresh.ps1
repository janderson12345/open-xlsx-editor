# What Excel's Refresh does to a query's table when the query's result changes: each case opens the fixture
# test/queries-excel.xlsx (excel-queries.ps1), types its cells (things beside and under the table), gives the query
# Sales a new formula (Workbook.Queries("Sales").Formula), refreshes its table (QueryTable.Refresh, not in the
# background) and saves; written out are the sheet's sheetData, the table, queryTable and connections parts, the
# workbook's definedNames and the table's range as COM reports it. Cases: [{ name, csv (the rows, "|" between them),
# cells: { A7: "below" } (none beside the table: typing there would make the table wider, Excel's AutoExpansion),
# column: { name, formula } (a column added to the table first, with ListColumns.Add, a column of the user's) }].
# An invisible instance, macros disabled for it only; no Excel setting changed, no keys sent.
# usage: pwsh tools/excel-compare/excel-queries-refresh.ps1 test/queries-excel.xlsx test/queries-refresh-excel.json
param([Parameter(Mandatory)][string]$Fixture, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$Fixture = [IO.Path]::GetFullPath($Fixture, (Get-Location).Path)
$cases = @(
  @{ name = 'grow'; csv = 'Region,Item,Qty|North,kiwi,30|South,apple,10|East,pear,20|West,date,50|North,fig,40|South,lime,5'; cells = @{ A7 = 'below'; E2 = 'beside' } }
  @{ name = 'shrink'; csv = 'Region,Item,Qty|North,kiwi,30|South,apple,10'; cells = @{ A7 = 'below'; E2 = 'beside' } }
  @{ name = 'columns'; csv = 'Region,Qty,Price|North,30,1.5|South,10,0.5|East,20,2|West,50,3'; cells = @{ A7 = 'below' } }
  @{ name = 'wider'; csv = 'Region,Item,Qty,Price|North,kiwi,30,1.5|South,apple,10,0.5'; cells = @{ G2 = 'beside'; A7 = 'below' } }
  @{ name = 'narrower'; csv = 'Region,Qty|North,30|South,10|East,20|West,50|North,40'; cells = @{ G2 = 'beside'; A7 = 'below'; C7 = 'under' } }
  @{ name = 'unbound'; csv = 'Region,Item,Qty,Price|North,kiwi,30,1.5|South,apple,10,0.5|East,pear,20,2|West,date,50,3|North,fig,40,1|South,lime,5,4'; cells = @{ G2 = 'beside'; A8 = 'below' }; column = @{ name = 'Double'; formula = '=[@Qty]*2' } }
  @{ name = 'same'; csv = 'Region,Item,Qty|North,kiwi,30|South,apple,10|East,pear,20|West,date,50'; cells = @{} }
)
$read = { param($zip, $name) $e = $zip.GetEntry($name); if ($e) { (New-Object IO.StreamReader($e.Open())).ReadToEnd() } else { $null } }
$pick = { param($xml, $pat) if ($xml -and $xml -match $pat) { $Matches[0] } else { $null } }
$dir = Join-Path ([IO.Path]::GetTempPath()) ("oxe-queries-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $dir | Out-Null
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$answers = New-Object System.Collections.Generic.List[object]
try {
  foreach ($case in $cases) {
    $a = [ordered]@{ name = $case.name; csv = $case.csv; cells = $case.cells; errors = @() }
    $wb = $x.Workbooks.Open($Fixture, 0, $true)
    try {
      $ws = $wb.Worksheets.Item(1)
      $lo = $ws.ListObjects.Item(1)
      if ($case.column) { $lc = $lo.ListColumns.Add(); $lc.Name = $case.column.name; $lc.DataBodyRange.Formula = $case.column.formula }
      foreach ($k in $case.cells.Keys) { $ws.Range($k).Value2 = $case.cells[$k] }
      $cols = ($case.csv.Split('|')[0].Split(',')).Count
      $text = $case.csv.Replace('|', '#(lf)')
      $wb.Queries.Item('Sales').Formula = "let`n    Source = Csv.Document(`"$text`", [Delimiter=`",`", Columns=$cols, QuoteStyle=QuoteStyle.None]),`n    #`"Promoted Headers`" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])`nin`n    #`"Promoted Headers`""
      try { [void]$lo.QueryTable.Refresh($false) } catch { $a.errors += "refresh: $($_.Exception.Message)" }
      $a.range = $lo.Range.Address($false, $false)
      $a.values = @($ws.UsedRange.Rows | ForEach-Object { , @($_.Cells | ForEach-Object { $_.Value2 }) })
      $a.used = $ws.UsedRange.Address($false, $false)
      $file = Join-Path $dir "$($case.name).xlsx"
      $wb.SaveAs($file, 51)
    } finally { $wb.Close($false) }
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    try {
      $a.sheetData = & $pick (& $read $zip 'xl/worksheets/sheet1.xml') '(?s)<sheetData\b[^>]*?(/>|>.*?</sheetData>)'
      $a.strings = @([regex]::Matches([string](& $read $zip 'xl/sharedStrings.xml'), '<t[^>]*>([^<]*)</t>') | ForEach-Object { $_.Groups[1].Value })
      $a.table = & $read $zip 'xl/tables/table1.xml'
      $a.queryTable = & $read $zip 'xl/queryTables/queryTable1.xml'
      $a.connections = & $read $zip 'xl/connections.xml'
      $a.definedNames = & $pick (& $read $zip 'xl/workbook.xml') '(?s)<definedNames\b.*?</definedNames>'
    } finally { $zip.Dispose() }
    $answers.Add($a)
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  Remove-Item -Recurse -Force $dir -ErrorAction SilentlyContinue
}
$json = ConvertTo-Json -InputObject $answers.ToArray() -Depth 6
[IO.File]::WriteAllText(([IO.Path]::GetFullPath($Out, (Get-Location).Path)), $json + "`n", (New-Object Text.UTF8Encoding($false)))
