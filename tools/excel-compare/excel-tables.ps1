# Has Excel make and change tables (ListObjects) and says what it saves: each case is a new workbook holding one of a
# few made-up lists, its steps run over COM (ListObjects.Add as Insert > Table and Format as Table make one; TableStyle,
# ShowTotals, ShowHeaders and the Table Style Options; Name; Resize; a column's TotalsCalculation; Unlist as Convert to
# Range; a value typed), then saved to a temporary folder; written out are the table parts, the sheet's sheetData,
# tableParts and relationships, the table content types, the workbook's definedNames and styles.xml's dxfs, tableStyles
# and cellXfs (the rest, docProps with the user's name, is not), and the table as COM reports it. Cases:
# [{ name, data: "basic" | "numbers" | "heads", steps: [{ do: "add", range, headers: "yes" | "no" | "guess" },
#   { do: "style", value }, { do: "set", prop, value }, { do: "name", value }, { do: "resize", range },
#   { do: "calc", col, value }, { do: "unlist" }, { do: "type", cell, value }] }].
# An invisible instance, macros disabled for it only; no Excel setting changed, no keys sent.
# usage: pwsh tools/excel-compare/excel-tables.ps1 tools/excel-compare/table-cases.json test/tablemake-excel.json
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$list = @(Get-Content $Cases -Raw | ConvertFrom-Json)
$missing = [Type]::Missing
$dir = Join-Path ([IO.Path]::GetTempPath()) ("oxe-tables-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $dir | Out-Null
# (The made-up lists, each put at B2: basic has a header row of words, numbers a first row of numbers, heads a header
# row with a number, a blank, a repeated name and a formula.)
$data = @{
  basic = @(@('Region', 'Item', 'Qty', 'Price'), @('North', 'kiwi', 30, 1.5), @('South', 'apple', 10, 0.5),
    @('East', 'pear', 20, 2), @('West', 'date', 50, 3), @('North', 'fig', 40, 2.5))
  numbers = @(@(1, 2, 3), @(4, 5, 6), @(7, 8, 9))
  heads = @(@('Qty', 2024, '', 'Qty', '=1+1'), @('a', 1, 2, 3, 4), @('b', 5, 6, 7, 8))
}
$read = { param($zip, $name) $e = $zip.GetEntry($name); if ($e) { (New-Object IO.StreamReader($e.Open())).ReadToEnd() } else { $null } }
$pick = { param($xml, $pat) if ($xml -and $xml -match $pat) { $Matches[0] } else { $null } }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$answers = New-Object System.Collections.Generic.List[object]
try {
  foreach ($case in $list) {
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $rows = $data[[string]$case.data]
    for ($i = 0; $i -lt $rows.Count; $i++) { for ($j = 0; $j -lt $rows[$i].Count; $j++) { $v = [string]$rows[$i][$j]; if ($v -ne '') { $ws.Cells.Item($i + 2, $j + 2).Formula = $v } } }
    $a = [ordered]@{ name = $case.name; errors = @() }
    $lo = $null
    foreach ($st in @($case.steps)) {
      try {
        switch ($st.do) {
          'add' { $h = @{ yes = 1; no = 2; guess = 0 }[[string]$st.headers]; $lo = $ws.ListObjects.Add(1, $ws.Range($st.range), $missing, $h) }
          'style' { $lo.TableStyle = [string]$st.value }
          'set' { $lo.($st.prop) = [bool]$st.value }
          'name' { $lo.Name = [string]$st.value }
          'resize' { $lo.Resize($ws.Range($st.range)) }
          'calc' { $lo.ListColumns.Item([int]$st.col).TotalsCalculation = [int]$st.value }
          'unlist' { $lo.Unlist(); $lo = $null }
          'type' { $ws.Range($st.cell).Formula = [string]$st.value }
        }
      } catch { $a.errors += "$($st.do): $($_.Exception.Message)" }
    }
    if ($lo) {
      $a.com = [ordered]@{ name = $lo.Name; range = $lo.Range.Address($false, $false); style = [string]$lo.TableStyle.Name
        headers = $lo.ShowHeaders; totals = $lo.ShowTotals; columns = @($lo.ListColumns | ForEach-Object { $_.Name })
        totalsCalc = @($lo.ListColumns | ForEach-Object { $_.TotalsCalculation }) }
    }
    $file = Join-Path $dir "$($case.name).xlsx"
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    try {
      $sheet = & $read $zip 'xl/worksheets/sheet1.xml'
      $a.sheetData = & $pick $sheet '(?s)<sheetData\b[^>]*?(/>|>.*?</sheetData>)'
      # (The shared strings' texts, in order, so the cells' strings read as text.)
      $a.strings = @([regex]::Matches([string](& $read $zip 'xl/sharedStrings.xml'), '<t[^>]*>([^<]*)</t>') | ForEach-Object { $_.Groups[1].Value })
      $a.tableParts = & $pick $sheet '(?s)<tableParts\b[^>]*?(/>|>.*?</tableParts>)'
      $a.sheetRels = & $read $zip 'xl/worksheets/_rels/sheet1.xml.rels'
      $types = & $read $zip '[Content_Types].xml'
      $a.contentTypes = @([regex]::Matches($types, '<Override [^>]*table[^>]*/>') | ForEach-Object { $_.Value })
      $a.definedNames = & $pick (& $read $zip 'xl/workbook.xml') '(?s)<definedNames\b.*?</definedNames>'
      $styles = & $read $zip 'xl/styles.xml'
      $a.dxfs = & $pick $styles '(?s)<dxfs\b[^>]*?(/>|>.*?</dxfs>)'
      $a.tableStyles = & $pick $styles '(?s)<tableStyles\b[^>]*?(/>|>.*?</tableStyles>)'
      $a.cellXfs = & $pick $styles '(?s)<cellXfs\b.*?</cellXfs>'
      foreach ($part in 'fonts', 'fills', 'borders') { $a[$part] = & $pick $styles ("(?s)<$part" + '\b.*?</' + "$part>") }
      $a.tables = @($zip.Entries | Where-Object { $_.FullName -like 'xl/tables/*' } | ForEach-Object { [ordered]@{ path = $_.FullName; xml = (& $read $zip $_.FullName) } })
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
