# Has Excel do editing steps on a small sheet and says what came of them (cases: [{ name, setup, steps, read }]):
# inserting and deleting cells (Range.Insert and Delete, with the shift given or left to Excel), Paste Special
# (Range.Copy, then PasteSpecial with what to paste, the operation, Skip blanks and Transpose), Paste Link
# (Worksheet.Paste with Link), Fill (Range.FillDown, FillRight, FillUp, FillLeft) and the fill handle
# (Range.AutoFill). Each case is a new workbook in an invisible instance (macros disabled for it only), closed without
# saving; no Excel setting is changed and no keys are sent. For each case the cells of `read` on
# Sheet1 (and the cells given on Sheet2) are written out: formula, value, number format, bold, fill colour, bottom
# border, with the merged ranges, column widths and defined names, or the refusal when Excel refused a step.
#
# setup: { cells: { "B1": "1", "E1": "=B5" }, arrays: { "E20:E21": "=B1:B2*2" }, sheet2: { "A1": "=Sheet1!B5" },
#   formats: [{ range, bold, color (RGB as Excel's number), numFmt, border (true: a thin bottom border) }],
#   merges: ["B6:C6"], widths: { "B": 20 }, names: { "Four": "=Sheet1!$B$4" } }
# steps: [{ do: "insert" | "delete", range, shift: "down" | "right" | "up" | "left" (omitted: Excel decides) },
#   { do: "paste", from, to, paste: "all" | "formulas" | "values" | "formats" | "widths" | "formulasFormats" |
#     "valuesFormats" | "noBorders", op: "add" | "subtract" | "multiply" | "divide", skip: bool, transpose: bool },
#   { do: "link", from, to },
#   { do: "fill", range, dir: "down" | "right" | "up" | "left" } (Range.FillDown...),
#   { do: "autofill", from, to, type: "default" | "copy" | "series" | "formats" | "values" } (the fill handle: AutoFill)]
# usage: pwsh tools/excel-compare/excel-steps.ps1 tools/excel-compare/shift-cases.json test/shift-excel.json
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
# (Not $cases or $out: PowerShell's names ignore case, and those are the paths.)
$list = Get-Content $Cases -Raw | ConvertFrom-Json
$shifts = @{ down = -4121; right = -4161; up = -4162; left = -4159 }
$pastes = @{ all = -4104; formulas = -4123; values = -4163; formats = -4122; widths = 8; formulasFormats = 11; valuesFormats = 12; noBorders = 7 }
$fills = @{ default = 0; copy = 1; series = 2; formats = 3; values = 4 }
$ops = @{ none = -4142; add = 2; subtract = 3; multiply = 4; divide = 5 }
$missing = [Type]::Missing
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3 # msoAutomationSecurityForceDisable: this instance runs no macros
$answers = New-Object System.Collections.Generic.List[object]
function Props($o) { if ($null -eq $o) { @() } else { $o.PSObject.Properties } }
function CellOf($c) {
  $a = [ordered]@{}
  if ($c.HasFormula) { $a.f = [string]$c.Formula }
  $v = $c.Value2
  if ($null -ne $v) { $a.v = if ($v -is [double] -or $v -is [bool]) { $v } elseif ($v -is [int]) { [double]$v } else { "$v" } }
  if ($c.NumberFormat -ne 'General') { $a.fmt = [string]$c.NumberFormat }
  if ($c.Font.Bold) { $a.bold = $true }
  if ($c.Interior.ColorIndex -ne -4142) { $a.color = [int]$c.Interior.Color }
  if ($c.Borders.Item(9).LineStyle -ne -4142) { $a.border = $true }
  $a
}
try {
  foreach ($case in $list) {
    $wb = $x.Workbooks.Add()
    while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $ws = $wb.Worksheets.Item(1); $ws2 = $wb.Worksheets.Item(2)
    $s = $case.setup
    foreach ($p in Props $s.cells) { $ws.Range($p.Name).Formula = [string]$p.Value }
    foreach ($p in Props $s.arrays) { $ws.Range($p.Name).FormulaArray = [string]$p.Value }
    foreach ($p in Props $s.sheet2) { $ws2.Range($p.Name).Formula = [string]$p.Value }
    foreach ($f in @($s.formats)) {
      if ($null -eq $f) { continue }
      $r = $ws.Range($f.range)
      if ($f.bold) { $r.Font.Bold = $true }
      if ($null -ne $f.color) { $r.Interior.Color = [int]$f.color }
      if ($f.numFmt) { $r.NumberFormat = [string]$f.numFmt }
      if ($f.border) { $r.Borders.Item(9).LineStyle = 1 }
    }
    foreach ($m in @($s.merges)) { if ($m) { [void]$ws.Range($m).Merge() } }
    foreach ($p in Props $s.widths) { $ws.Columns.Item($p.Name).ColumnWidth = [double]$p.Value }
    foreach ($p in Props $s.names) { [void]$wb.Names.Add($p.Name, [string]$p.Value) }
    $a = [ordered]@{ name = [string]$case.name }
    try {
      foreach ($st in $case.steps) {
        switch ($st.do) {
          'insert' { if ($st.shift) { [void]$ws.Range($st.range).Insert($shifts[$st.shift]) } else { [void]$ws.Range($st.range).Insert() } }
          'delete' { if ($st.shift) { [void]$ws.Range($st.range).Delete($shifts[$st.shift]) } else { [void]$ws.Range($st.range).Delete() } }
          'paste' {
            [void]$ws.Range($st.from).Copy()
            $what = if ($st.paste) { $pastes[$st.paste] } else { -4104 }
            $op = if ($st.op) { $ops[$st.op] } else { -4142 }
            [void]$ws.Range($st.to).PasteSpecial($what, $op, [bool]$st.skip, [bool]$st.transpose)
            $x.CutCopyMode = 0
          }
          'fill' {
            $r = $ws.Range($st.range)
            switch ($st.dir) { 'down' { [void]$r.FillDown() } 'right' { [void]$r.FillRight() } 'up' { [void]$r.FillUp() } 'left' { [void]$r.FillLeft() } }
          }
          'autofill' {
            $t = if ($st.type) { $fills[$st.type] } else { 0 }
            [void]$ws.Range($st.from).AutoFill($ws.Range($st.to), $t)
          }
          'link' {
            [void]$ws.Range($st.from).Copy()
            [void]$ws.Activate(); [void]$ws.Range($st.to).Select()
            [void]$ws.Paste($missing, $true)
            $x.CutCopyMode = 0
          }
        }
      }
    } catch { $a.refused = $true; $a.error = [string]$_.Exception.Message }
    $cells = [ordered]@{}
    foreach ($c in $ws.Range($case.read).Cells) {
      $one = CellOf $c
      if ($one.Count) { $cells[$c.Address($false, $false)] = $one }
    }
    $a.cells = $cells
    $two = [ordered]@{}
    foreach ($p in Props $s.sheet2) { $c = $ws2.Range($p.Name); $two[$p.Name] = [string]$c.Formula }
    if ($two.Count) { $a.sheet2 = $two }
    $merged = New-Object System.Collections.Generic.List[string]
    foreach ($c in $ws.Range($case.read).Cells) { if ($c.MergeCells) { $ad = $c.MergeArea.Address($false, $false); if (-not $merged.Contains($ad)) { $merged.Add($ad) } } }
    if ($merged.Count) { $a.merges = $merged }
    $w = [ordered]@{}
    $cols = $ws.Range($case.read).Columns
    foreach ($col in $cols) { if ($col.ColumnWidth -ne $ws.StandardWidth) { $w[($col.Address($false, $false) -replace '\d', '' -replace ':.*', '')] = [double]$col.ColumnWidth } }
    if ($w.Count) { $a.widths = $w }
    $n = [ordered]@{}
    foreach ($nm in $wb.Names) { $n[$nm.Name] = [string]$nm.RefersTo }
    if ($n.Count) { $a.names = $n }
    $answers.Add($a)
    $wb.Close($false)
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
[ordered]@{ excel = 'Excel.Application: Range.Insert, Range.Delete, Range.Copy and PasteSpecial, Worksheet.Paste with Link'; cases = $answers } | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 $Out
"$($answers.Count) cases"
