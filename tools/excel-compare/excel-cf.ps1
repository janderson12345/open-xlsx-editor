# Has Excel apply conditional formats to a small sheet and says what each cell shows (cases: [{ name, setup: { cells,
# formats }, rules: [{ range, kind, ... }], read }]): each cell's fill, font colour and bold as Excel draws them
# (Range.DisplayFormat), and, from the workbook Excel then saves, its conditionalFormatting elements and dxfs, which
# the tests put into a sheet of their own. Rules (FormatConditions):
#   { kind: "cell", op: "between" | "notBetween" | "equal" | "notEqual" | "greater" | "less" | "greaterEqual" |
#     "lessEqual", f1, f2 } (xlCellValue); { kind: "expr", f1 } (xlExpression); { kind: "text", op: "contains" |
#     "notContains" | "begins" | "ends", text }; { kind: "top", rank, percent, bottom } (AddTop10); { kind: "average",
#     how: 0..3 } (AddAboveAverage: above, below, equal or above, equal or below); { kind: "unique", dup } (AddUniqueValues);
#     { kind: "blanks" | "noBlanks" | "errors" }; { kind: "scale", colors: [RGB, RGB(, RGB)] } (AddColorScale: lowest,
#     (50th percentile,) highest); each with fill, color (RGB as Excel's number), bold and stop.
# An invisible instance, macros disabled for it only; each workbook saved to a temporary folder, read, and deleted.
# usage: pwsh tools/excel-compare/excel-cf.ps1 tools/excel-compare/cf-cases.json test/cf-excel.json
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$list = Get-Content $Cases -Raw | ConvertFrom-Json
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('oxe-cf-' + [guid]::NewGuid().ToString('n').Substring(0, 8))
New-Item -ItemType Directory $tmp | Out-Null
$ops = @{ between = 1; notBetween = 2; equal = 3; notEqual = 4; greater = 5; less = 6; greaterEqual = 7; lessEqual = 8 }
$textOps = @{ contains = 0; notContains = 1; begins = 2; ends = 3 }
$m = [Type]::Missing
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3 # msoAutomationSecurityForceDisable: this instance runs no macros
$answers = New-Object System.Collections.Generic.List[object]
function Props($o) { if ($null -eq $o) { @() } else { $o.PSObject.Properties } }
function PartText($zip, $name) { $e = $zip.GetEntry($name); if (-not $e) { return '' }; $r = New-Object IO.StreamReader($e.Open()); try { $r.ReadToEnd() } finally { $r.Dispose() } }
try {
  $i = 0
  foreach ($case in $list) {
    $i++
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    foreach ($p in Props $case.setup.cells) { $ws.Range($p.Name).Formula = [string]$p.Value }
    foreach ($f in @($case.setup.formats)) { if ($f -and $f.numFmt) { $ws.Range($f.range).NumberFormat = [string]$f.numFmt } }
    foreach ($rule in $case.rules) {
      $fc = $ws.Range($rule.range).FormatConditions
      switch ($rule.kind) {
        'cell' { $f2 = if ($rule.f2) { [string]$rule.f2 } else { $m }; $c = $fc.Add(1, $ops[$rule.op], [string]$rule.f1, $f2) }
        'expr' { $c = $fc.Add(2, $m, [string]$rule.f1) }
        'text' { $c = $fc.Add(9, $m, $m, $m, [string]$rule.text, $textOps[$rule.op]) }
        'top' { $c = $fc.AddTop10(); $c.TopBottom = $(if ($rule.bottom) { 0 } else { 1 }); $c.Rank = [int]$rule.rank; $c.Percent = [bool]$rule.percent }
        'average' { $c = $fc.AddAboveAverage(); $c.AboveBelow = [int]$rule.how }
        'unique' { $c = $fc.AddUniqueValues(); $c.DupeUnique = $(if ($rule.dup) { 1 } else { 0 }) }
        'blanks' { $c = $fc.Add(10) }
        'noBlanks' { $c = $fc.Add(13) }
        'errors' { $c = $fc.Add(16) }
        'scale' {
          $c = $fc.AddColorScale(@($rule.colors).Count)
          $k = 1; foreach ($col in $rule.colors) { $c.ColorScaleCriteria.Item($k).FormatColor.Color = [int]$col; $k++ }
        }
      }
      if ($rule.kind -ne 'scale') {
        if ($null -ne $rule.fill) { $c.Interior.Color = [int]$rule.fill }
        if ($null -ne $rule.color) { $c.Font.Color = [int]$rule.color }
        if ($rule.bold) { $c.Font.Bold = $true }
        if ($rule.stop) { $c.StopIfTrue = $true }
      }
    }
    $a = [ordered]@{ name = [string]$case.name }
    $cells = [ordered]@{}
    foreach ($cell in $ws.Range($case.read).Cells) {
      $d = $cell.DisplayFormat
      $one = [ordered]@{}
      if ($d.Interior.ColorIndex -ne -4142) { $one.fill = [int]$d.Interior.Color }
      if ([int]$d.Font.Color -ne 0) { $one.color = [int]$d.Font.Color }
      if ($d.Font.Bold) { $one.bold = $true }
      if ($one.Count) { $cells[$cell.Address($false, $false)] = $one }
    }
    $a.cells = $cells
    $file = Join-Path $tmp "$i.xlsx"
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    try {
      $sheet = PartText $zip 'xl/worksheets/sheet1.xml'
      $a.conditionalFormatting = @([regex]::Matches($sheet, '<conditionalFormatting .*?</conditionalFormatting>') | ForEach-Object { $_.Value })
      $dx = [regex]::Match((PartText $zip 'xl/styles.xml'), '<dxfs[^>]*?(/>|>.*?</dxfs>)')
      $a.dxfs = if ($dx.Success) { $dx.Value } else { $null }
    } finally { $zip.Dispose() }
    $answers.Add($a)
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
}
[ordered]@{ excel = 'Excel.Application: FormatConditions, Range.DisplayFormat, then the saved sheet and styles XML'; cases = $answers } | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 $Out
"$($answers.Count) cases"
