# Has Excel sort a small made-up table and says what it saves of the sort: the worksheet's <sortState> and its
# <autoFilter> (where a filtered range keeps its sortState), as the sheet's XML holds them after a save. Cases:
# [{ name, filter: bool (AutoFilter on A1's region first), legacy: bool (Range.Sort, as the A to Z button sorts),
#    range, header: "yes" | "no", matchCase, leftToRight, keys: [{ key, order: "asc" | "desc", on: "cellColor" |
#    "fontColor", color, customOrder }] }]. Each case is a new workbook in an invisible instance (macros disabled for it
# only), saved to a temporary folder and deleted; only the two elements are written out (the rest of the package, which
# holds the user's name in docProps, is not). No Excel setting is changed and no keys are sent.
# usage: pwsh tools/excel-compare/excel-sortstate.ps1 tools/excel-compare/sortstate-cases.json test/sortstate-excel.json
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$list = Get-Content $Cases -Raw | ConvertFrom-Json
$missing = [Type]::Missing
$dir = Join-Path ([IO.Path]::GetTempPath()) ("oxe-sortstate-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $dir | Out-Null
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$answers = New-Object System.Collections.Generic.List[object]
# (The same made-up table each time: a header row, five rows of fruit, a count and a colour word.)
$rows = @(@('Fruit', 'Count', 'Colour'), @('kiwi', 30, 'green'), @('Apple', 10, 'red'), @('pear', 20, 'green'),
  @('Date', 50, 'brown'), @('fig', 40, 'purple'))
try {
  foreach ($case in $list) {
   try {
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    for ($i = 0; $i -lt $rows.Count; $i++) { for ($j = 0; $j -lt 3; $j++) { $v = $rows[$i][$j]; $ws.Cells.Item($i + 1, $j + 1).Formula = [string]$v } }
    # (Colours for Sort On: rows 3 and 5 filled red, row 4's text blue.)
    $ws.Range('B3').Interior.Color = 255; $ws.Range('B5').Interior.Color = 255; $ws.Range('B4').Font.Color = 16711680
    $a = [ordered]@{ name = $case.name }
    $hdr = if ($case.header -eq 'no') { 2 } else { 1 }
    $way = if ($case.leftToRight) { 2 } else { 1 }
    $keys = @($case.keys)
    if ($case.filter) { [void]$ws.Range('A1').AutoFilter() }
    $r = $ws.Range($case.range)
    if ($case.legacy) {
      $k1 = $ws.Range($keys[0].key); $o1 = if ($keys[0].order -eq 'desc') { 2 } else { 1 }
      [void]$r.Sort($k1, $o1, $missing, $missing, 1, $missing, 1, $hdr, 1, [bool]$case.matchCase, $way)
    } else {
      # (A filtered range is sorted through its AutoFilter's Sort, as its buttons and the Sort box sort it.)
      $so = if ($case.filter) { $ws.AutoFilter.Sort } else { $ws.Sort }
      $so.SortFields.Clear()
      foreach ($k in $keys) {
        $order = if ($k.order -eq 'desc') { 2 } else { 1 }
        $on = @{ cellColor = 1; fontColor = 2 }[[string]$k.on]
        $custom = if ($k.customOrder) { [string]$k.customOrder } else { $missing }
        $sf = $so.SortFields.Add($ws.Range($k.key), $(if ($on) { $on } else { 0 }), $order, $custom, 0)
        if ($on) { $sf.SortOnValue.Color = [int]$k.color }
      }
      if (-not $case.filter) { $so.SetRange($r) }
      $so.Header = $hdr
      $so.MatchCase = [bool]$case.matchCase
      $so.Orientation = $way
      $so.Apply()
    }
    $file = Join-Path $dir "$($case.name).xlsx"
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    try {
      $sheet = (New-Object IO.StreamReader($zip.GetEntry('xl/worksheets/sheet1.xml').Open())).ReadToEnd()
      $styles = (New-Object IO.StreamReader($zip.GetEntry('xl/styles.xml').Open())).ReadToEnd()
    } finally { $zip.Dispose() }
    $a.autoFilter = if ($sheet -match '(?s)<autoFilter\b.*?(</autoFilter>|/>)') { $Matches[0] } else { $null }
    $a.sortState = if ($sheet -match '(?s)<sortState\b.*?</sortState>') { $Matches[0] } else { $null }
    $a.sortStates = [regex]::Matches($sheet, '<sortState' + [char]92 + 'b').Count
    # (Where the sortState sits among the sheet's elements: the names before and after it.)
    $names = [regex]::Matches($sheet, '<(\w+)[ >/]') | ForEach-Object { $_.Groups[1].Value } | Where-Object { $_ -in 'sheetData','autoFilter','sortState','mergeCells','conditionalFormatting','pageMargins','pageSetup','sheetProtection','dataValidations' }
    $a.order = @($names | Select-Object -Unique)
    $a.dxfs = if ($styles -match '(?s)<dxfs count="0"/>|<dxfs' + [char]92 + 'b.*?</dxfs>') { $Matches[0] } else { $null }
    $answers.Add($a)
    Remove-Item $file
   } catch { $answers.Add([ordered]@{ name = $case.name; error = [string]$_.Exception.Message; at = [string]$_.InvocationInfo.ScriptLineNumber }); try { $wb.Close($false) } catch {} }
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Remove-Item -Recurse -Force $dir -ErrorAction SilentlyContinue
}
[ordered]@{ excel = 'Excel.Application: Worksheet.Sort, AutoFilter.Sort and Range.Sort, saved; the sheet''s sortState and autoFilter'; cases = $answers } | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 $Out
