# Has Excel make a workbook with Power Query queries and says what it saves (the fixture test/queries-excel.xlsx and
# test/queries-excel.json): Sales (a CSV text read with Csv.Document, its headers promoted and its types changed) loaded
# to a table at Sheet1!A1 as Data > From Table/Range's Close & Load does (ListObjects.Add over the Mashup OLE DB
# provider, refreshed); Big (Sales's rows with Qty over 15) as a connection only (Connections.Add2); Model (Sales again)
# loaded to the data model only (Connections.Add2 with CreateModelConnection). The M reads no file, so it refreshes
# anywhere. Written out: the file, the docProps names put back as the editor's (Excel writes the user's name there),
# and what COM reports: each query's Formula, each connection's name, type and command text, the table's range and
# its values.
# An invisible instance, macros disabled for it only; no Excel setting changed, no keys sent.
# usage: pwsh tools/excel-compare/excel-queries.ps1 test/queries-excel.xlsx test/queries-excel.json
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Json)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$Out = [IO.Path]::GetFullPath($Out, (Get-Location).Path)
$Json = [IO.Path]::GetFullPath($Json, (Get-Location).Path)
$sales = @'
let
    Source = Csv.Document("Region,Item,Qty#(lf)North,kiwi,30#(lf)South,apple,10#(lf)East,pear,20#(lf)West,date,50", [Delimiter=",", Columns=3, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Region", type text}, {"Item", type text}, {"Qty", Int64.Type}})
in
    #"Changed Type"
'@
$big = @'
let
    Source = Sales,
    #"Filtered Rows" = Table.SelectRows(Source, each [Qty] > 15)
in
    #"Filtered Rows"
'@
$model = @'
let
    Source = Sales
in
    Source
'@
$conn = { param($name) "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook`$;Location=$name;Extended Properties=`"`"" }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$answer = [ordered]@{ errors = @() }
try {
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  [void]$wb.Queries.Add('Sales', $sales.Replace("`r`n", "`n"))
  [void]$wb.Queries.Add('Big', $big.Replace("`r`n", "`n"))
  [void]$wb.Queries.Add('Model', $model.Replace("`r`n", "`n"))
  $lo = $ws.ListObjects.Add(0, (& $conn 'Sales'), [Type]::Missing, 1, $ws.Range('A1'))
  $qt = $lo.QueryTable
  $qt.CommandType = 2
  $qt.CommandText = 'SELECT * FROM [Sales]'
  $qt.RowNumbers = $false; $qt.FillAdjacentFormulas = $false; $qt.PreserveFormatting = $true
  $qt.RefreshOnFileOpen = $false; $qt.BackgroundQuery = $true; $qt.RefreshStyle = 1; $qt.SavePassword = $false
  $qt.SaveData = $true; $qt.AdjustColumnWidth = $true; $qt.RefreshPeriod = 0; $qt.PreserveColumnInfo = $true
  $qt.ListObject.DisplayName = 'Sales'
  [void]$qt.Refresh($false)
  # (Named as Close & Load names it.)
  $qt.WorkbookConnection.Name = 'Query - Sales'
  $qt.WorkbookConnection.Description = "Connection to the 'Sales' query in the workbook."
  # (The arguments as Excel's macro recorder writes them for Only Create Connection and for Add to the Data Model.)
  [void]$wb.Connections.Add2('Query - Big', "Connection to the 'Big' query in the workbook.", (& $conn 'Big'), 'SELECT * FROM [Big]', 2, $false, $false)
  try { [void]$wb.Connections.Add2('Query - Model', "Connection to the 'Model' query in the workbook.", (& $conn 'Model'), '"Model"', 6, $true, $false) }
  catch { $answer.errors += "model: $($_.Exception.Message)" }
  # (Queries and Connections are walked by index: PowerShell can't enumerate Queries.)
  $answer.queries = @(1..$wb.Queries.Count | ForEach-Object { $q = $wb.Queries.Item($_); [ordered]@{ name = $q.Name; formula = $q.Formula; description = $q.Description } })
  $answer.connections = @(1..$wb.Connections.Count | ForEach-Object { $c = $wb.Connections.Item($_); [ordered]@{ name = $c.Name; type = $c.Type; inModel = $c.InModel
    command = $(try { [string]$c.OLEDBConnection.CommandText } catch { '' }) } })
  $answer.table = [ordered]@{ name = $lo.Name; range = $lo.Range.Address($false, $false)
    values = @($lo.Range.Rows | ForEach-Object { , @($_.Cells | ForEach-Object { $_.Value2 }) }) }
  if (Test-Path $Out) { Remove-Item $Out }
  $wb.SaveAs($Out, 51)
  $wb.Close($false)
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
# (docProps/core.xml holds the user's name as creator and last modified by: the editor's name put there instead.)
$zip = [IO.Compression.ZipFile]::Open($Out, 'Update')
try {
  $e = $zip.GetEntry('docProps/core.xml')
  $r = New-Object IO.StreamReader($e.Open()); $core = $r.ReadToEnd(); $r.Dispose()
  $core = [regex]::Replace($core, '<dc:creator>[^<]*</dc:creator>', '<dc:creator>open-xlsx-editor</dc:creator>')
  $core = [regex]::Replace($core, '<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>', '<cp:lastModifiedBy>open-xlsx-editor</cp:lastModifiedBy>')
  $e.Delete()
  $n = $zip.CreateEntry('docProps/core.xml')
  $w = New-Object IO.StreamWriter($n.Open(), (New-Object Text.UTF8Encoding($false))); $w.Write($core); $w.Dispose()
} finally { $zip.Dispose() }
$text = ConvertTo-Json -InputObject $answer -Depth 8
[IO.File]::WriteAllText($Json, $text + "`n", (New-Object Text.UTF8Encoding($false)))
