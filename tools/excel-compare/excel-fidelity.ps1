# The save-fidelity benchmark, second half: Excel opens each original workbook and the editor's saved copy
# (fidelity.mjs) and compares what it sees, sheet by sheet: every cell of the original's used range (formula, value,
# number format; the cell the editor typed into left aside), and how many charts, pivot tables, tables, shapes, notes,
# threaded comments, hyperlinks, conditional formats, sparkline groups and validated cells each sheet has, with the
# workbook's names. An invisible instance, macros disabled for it only, both files opened read-only and closed without
# saving.
# usage: pwsh tools/excel-compare/excel-fidelity.ps1 <folder of originals> <out folder of fidelity.mjs>
param([Parameter(Mandatory)][string]$Books, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$Books = (Resolve-Path $Books).Path; $Out = (Resolve-Path $Out).Path
$parts = Get-Content (Join-Path $Out 'parts.json') -Raw | ConvertFrom-Json
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3 # msoAutomationSecurityForceDisable: this instance runs no macros
function Count($block) { try { & $block } catch { 0 } }
function Look($ws) {
  [ordered]@{
    charts = Count { $ws.ChartObjects().Count }; pivots = Count { $ws.PivotTables().Count }; tables = Count { $ws.ListObjects.Count }
    shapes = Count { $ws.Shapes.Count }; notes = Count { $ws.Comments.Count }; threads = Count { $ws.CommentsThreaded.Count }
    links = Count { $ws.Hyperlinks.Count }; condFormats = Count { $ws.Cells.FormatConditions.Count }
    sparklines = Count { $ws.Cells.SparklineGroups.Count }; validated = Count { $ws.Cells.SpecialCells(-4174).Count }
    protected = [bool]$ws.ProtectContents; printTitles = [string]$ws.PageSetup.PrintTitleRows
  }
}
$report = New-Object System.Collections.Generic.List[object]
try {
  foreach ($p in $parts) {
    $a = $x.Workbooks.Open((Join-Path $Books $p.name), 0, $true)
    $one = [ordered]@{ name = $p.name; cells = 0; differ = @(); features = @() }
    # (A copy Excel won't open is the worst loss of all: said, and the next workbook taken.)
    try { $b = $x.Workbooks.Open((Join-Path $Out "saved-$($p.name)"), 0, $true) }
    catch { $a.Close($false); $one.unopened = $true; $report.Add($one); "$($p.name.PadRight(22)) EXCEL WOULD NOT OPEN THE SAVED COPY"; continue }
    if ($a.Names.Count -ne $b.Names.Count) { $one.features += "names $($a.Names.Count) -> $($b.Names.Count)" }
    foreach ($wa in $a.Worksheets) {
      $wbS = $b.Worksheets.Item($wa.Name)
      $la = Look $wa; $lb = Look $wbS
      foreach ($k in $la.Keys) { if ("$($la[$k])" -ne "$($lb[$k])") { $one.features += "$($wa.Name) $k $($la[$k]) -> $($lb[$k])" } }
      foreach ($c in $wa.UsedRange.Cells) {
        $addr = $c.Address($false, $false)
        if ("$($wa.Name)!$addr" -eq $p.edit) { continue }
        $d = $wbS.Range($addr)
        $one.cells++
        if ($c.Formula -ne $d.Formula -or "$($c.Value2)" -ne "$($d.Value2)" -or $c.NumberFormat -ne $d.NumberFormat) {
          if ($one.differ.Count -lt 5) { $one.differ += "$($wa.Name)!$addr" }
        }
      }
    }
    $b.Close($false); $a.Close($false)
    $report.Add($one)
    "$($p.name.PadRight(22)) cells $($one.cells.ToString().PadLeft(5)) as Excel had them: $(if ($one.differ.Count) { 'no: ' + ($one.differ -join ', ') } else { 'all' })   features: $(if ($one.features.Count) { $one.features -join '; ' } else { 'all the same' })"
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
$report | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 (Join-Path $Out 'excel.json')
