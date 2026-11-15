# Has Excel filter a small sheet (AutoFilter) and says what came of it (cases: [{ name, setup: { cells, formats },
# steps: [{ do: "on", cell } (Range.AutoFilter with no arguments: the filter turned on, or off when it is on) or
# { do: "filter", range, field (1-based), values: [shown texts, "=" for blanks] } (xlFilterValues) or { do: "clear",
# range, field } (that column's filter taken off) or { do: "sort", key, order } (AutoFilter.Sort, by the column of
# cell key, while filtered) or { do: "dates", range, field, groups: [[level, "m/d/yyyy"]] } (xlFilterValues with
# Criteria2: level 0 a year, 1 a month, 2 a day)], read }]): the cells' values (`values`, when the case asks), the
# rows hidden, the filter's range, and, from the workbook Excel then saves, the sheet's autoFilter element and the
# _xlnm._FilterDatabase name as Excel writes them. An invisible instance, macros disabled for it only; each workbook
# saved to a temporary folder, read, and deleted; no Excel setting is changed and no keys are sent.
# usage: pwsh tools/excel-compare/excel-filter.ps1 tools/excel-compare/filter-cases.json test/filter-excel.json
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$list = Get-Content $Cases -Raw | ConvertFrom-Json
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('oxe-filter-' + [guid]::NewGuid().ToString('n').Substring(0, 8))
New-Item -ItemType Directory $tmp | Out-Null
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3 # msoAutomationSecurityForceDisable: this instance runs no macros
$answers = New-Object System.Collections.Generic.List[object]
function Props($o) { if ($null -eq $o) { @() } else { $o.PSObject.Properties } }
function PartText($zip, $name) { $e = $zip.GetEntry($name); $r = New-Object IO.StreamReader($e.Open()); try { $r.ReadToEnd() } finally { $r.Dispose() } }
try {
  $i = 0
  foreach ($case in $list) {
    $i++
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    foreach ($p in Props $case.setup.cells) { $ws.Range($p.Name).Formula = [string]$p.Value }
    foreach ($f in @($case.setup.formats)) { if ($f -and $f.numFmt) { $ws.Range($f.range).NumberFormat = [string]$f.numFmt } }
    $a = [ordered]@{ name = [string]$case.name }
    try {
      foreach ($st in $case.steps) {
        if ($st.do -eq 'on') { [void]$ws.Range($st.cell).AutoFilter() }
        elseif ($st.do -eq 'clear') { [void]$ws.Range($st.range).AutoFilter([int]$st.field) }
        elseif ($st.do -eq 'sort') {
          $so = $ws.AutoFilter.Sort; $so.SortFields.Clear()
          [void]$so.SortFields.Add($ws.Range([string]$st.key), 0, $(if ($st.order -eq 'desc') { 2 } else { 1 }))
          $so.Header = 1; $so.Apply()
        }
        elseif ($st.do -eq 'dates') {
          $crit = New-Object 'object[]' ($st.groups.Count * 2)
          for ($g = 0; $g -lt $st.groups.Count; $g++) { $crit[2 * $g] = [int]$st.groups[$g][0]; $crit[2 * $g + 1] = [string]$st.groups[$g][1] }
          [void]$ws.Range($st.range).AutoFilter([int]$st.field, [Type]::Missing, 7, $crit)
        }
        else { [void]$ws.Range($st.range).AutoFilter([int]$st.field, [string[]]@($st.values), 7) }
      }
    } catch { $a.refused = $true; $a.error = [string]$_.Exception.Message }
    $a.filtering = [bool]$ws.AutoFilterMode
    if ($ws.AutoFilterMode) { $a.range = [string]$ws.AutoFilter.Range.Address($false, $false) }
    $hidden = New-Object System.Collections.Generic.List[int]
    foreach ($row in $ws.Range($case.read).Rows) { if ($row.EntireRow.Hidden) { $hidden.Add([int]$row.Row) } }
    $a.hidden = $hidden
    if ($case.values) { $vals = [ordered]@{}; foreach ($c in $ws.Range($case.read).Cells) { if ($null -ne $c.Value2) { $vals[$c.Address($false, $false)] = [string]$c.Value2 } }; $a.values = $vals }
    $file = Join-Path $tmp "$i.xlsx"
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    try {
      $sheet = PartText $zip 'xl/worksheets/sheet1.xml'
      $m = [regex]::Match($sheet, '<autoFilter[^>]*?(/>|>.*?</autoFilter>)')
      $a.autoFilter = if ($m.Success) { $m.Value } else { $null }
      $book = PartText $zip 'xl/workbook.xml'
      $n = [regex]::Match($book, '<definedName name="_xlnm._FilterDatabase"[^>]*>[^<]*</definedName>')
      $a.filterName = if ($n.Success) { $n.Value } else { $null }
      $rows = [regex]::Matches($sheet, '<row [^>]*hidden="1"[^>]*>') | ForEach-Object { $_.Value -replace '<c .*', '' }
      $a.hiddenRows = @($rows)
    } finally { $zip.Dispose() }
    $answers.Add($a)
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
}
[ordered]@{ excel = 'Excel.Application: Range.AutoFilter, then the saved sheet and workbook XML'; cases = $answers } | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 $Out
"$($answers.Count) cases"
