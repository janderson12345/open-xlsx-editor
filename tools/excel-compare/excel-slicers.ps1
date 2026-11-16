# What Excel does to slicers and timelines when rows, columns or cells are inserted or deleted on their sheet
# (slicer-cases.json, from slicer-cases.mjs): for each case a new workbook in an invisible Excel with a table Sales
# (made-up dates, regions, items, units), a slicer on its Region column, a PivotTable of it with a slicer on Item and
# a timeline on Date, all on one sheet; the step done, then where each slicer, the table and the pivot are read over
# COM, and the workbook saved and its slicer, timeline, cache, table, pivot and drawing parts kept as text.
# No keys sent, no setting changed.
# With -Open <folder>, each case's workbook is the editor's own result (<folder>\<case number>.xlsx, from
# test/slicers.test.js) opened instead and read back the same way; `repaired` says whether Excel had to repair it (a
# repair log, error*.xml, appeared in %TEMP%).
# usage: pwsh tools/excel-compare/excel-slicers.ps1 <slicer-cases.json> <out.json> [<folder to keep the books>] [-Open <folder>]
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out, [string]$Books = '', [string]$Open = '')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);'
$spec = Get-Content -Raw $Cases | ConvertFrom-Json
if ($Books) { New-Item -ItemType Directory -Force $Books | Out-Null; $Books = (Resolve-Path $Books).Path }
$partsOf = { param($file)
  $zip = [IO.Compression.ZipFile]::OpenRead($file)
  try {
    $parts = [ordered]@{}
    foreach ($e in $zip.Entries) {
      if ($e.FullName -notmatch '^xl/(slicers|slicerCaches|timelines|timelineCaches|tables|pivotTables|drawings|worksheets)/[^/]*\.xml$') { continue }
      $s = New-Object IO.StreamReader($e.Open()); $t = $s.ReadToEnd(); $s.Close()
      # (A worksheet only for its dimension and extLst; the rest is cells.)
      if ($e.FullName -match '^xl/worksheets/') { $t = [regex]::Match($t, '<extLst>.*</extLst>', 'Singleline').Value }
      $parts[$e.FullName] = $t
    }
    $parts
  } finally { $zip.Dispose() }
}
$x = New-Object -ComObject Excel.Application
$mine = [uint32]0
[void][W.U]::GetWindowThreadProcessId([IntPtr]$x.Hwnd, [ref]$mine)
Write-Host "Excel PID $mine"
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$results = @()
$n = 0
try {
  foreach ($case in $spec.cases) {
    $n++
    $refused = $null
    $repaired = $false
    if ($Open) {
      $logs = @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue).Count
      $wb = $x.Workbooks.Open((Join-Path (Resolve-Path $Open).Path "$n.xlsx"), 0, $true)
      $repaired = @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue).Count -ne $logs
      $ws = $wb.Worksheets.Item(1)
    } else {
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    [void]$ws.Activate()
    $rows = $spec.sheet.rows
    $top = $ws.Range($spec.sheet.at)
    for ($i = 0; $i -lt $rows.Count; $i++) {
      $row = @($rows[$i])
      for ($j = 0; $j -lt $row.Count; $j++) {
        $v = $row[$j]
        # (Two call sites: PowerShell's COM binder keeps the first value type a call site saw.)
        if ($v -is [string]) { $top.Offset($i, $j).Formula = [string]$v } else { $top.Offset($i, $j).Value2 = [double]$v }
      }
    }
    $top.Offset(1, 0).Resize($rows.Count - 1, 1).NumberFormat = 'yyyy-mm-dd'
    $lo = $ws.ListObjects.Add(1, $top.Resize($rows.Count, @($rows[0]).Count), [Type]::Missing, 1); $lo.Name = 'Sales'
    $place = { param($at) $r = $ws.Range($at); @($r.Top, $r.Left, $r.Width, $r.Height) }
    $p = & $place $spec.sheet.tableSlicer
    $sc = $wb.SlicerCaches.Add2($lo, 'Region')
    [void]$sc.Slicers.Add($ws, [Type]::Missing, 'Region', 'Region', $p[0], $p[1], $p[2], $p[3])
    $pc = $wb.PivotCaches().Create(1, 'Sales')
    $pt = $pc.CreatePivotTable($ws.Range($spec.sheet.pivot), 'Pivot1')
    $pt.PivotFields('Region').Orientation = 1
    [void]$pt.AddDataField($pt.PivotFields('Units'))
    $p = & $place $spec.sheet.pivotSlicer
    $ps = $wb.SlicerCaches.Add2($pt, 'Item')
    [void]$ps.Slicers.Add($ws, [Type]::Missing, 'Item', 'Item', $p[0], $p[1], $p[2], $p[3])
    $p = & $place $spec.sheet.timeline
    $tc = $wb.SlicerCaches.Add2($pt, 'Date', [Type]::Missing, 2)
    [void]$tc.Slicers.Add($ws, [Type]::Missing, 'Date', 'Date', $p[0], $p[1], $p[2], $p[3])
    $s = $case.step
    try {
      $r = $ws.Range($s.range)
      $shift = if ($s.shift -eq 'down') { -4121 } elseif ($s.shift -eq 'right') { -4161 } elseif ($s.shift -eq 'up') { -4162 } elseif ($s.shift -eq 'left') { -4159 } else { $null }
      if ($s.do -eq 'insert') { if ($null -ne $shift) { [void]$r.Insert($shift) } else { [void]$r.Insert() } }
      if ($s.do -eq 'delete') { if ($null -ne $shift) { [void]$r.Delete($shift) } else { [void]$r.Delete() } }
    } catch { $refused = ($_.Exception.InnerException ?? $_.Exception).Message }
    }
    $slicers = @()
    foreach ($cache in $wb.SlicerCaches) {
      foreach ($sl in $cache.Slicers) { $slicers += "$($sl.Name) $($sl.Shape.TopLeftCell.Address($false, $false)):$($sl.Shape.BottomRightCell.Address($false, $false)) $($sl.Top),$($sl.Left) $($sl.Width)x$($sl.Height)" }
    }
    $table = if ($ws.ListObjects.Count) { $ws.ListObjects.Item(1).Range.Address($false, $false) } else { $null }
    $pivot = if ($ws.PivotTables().Count) { $ws.PivotTables(1).TableRange1.Address($false, $false) } else { $null }
    $file = Join-Path $env:TEMP "oxe-sl$n.xlsx"
    if (Test-Path $file) { Remove-Item $file }
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $parts = & $partsOf $file
    if ($Books) { Move-Item -Force $file (Join-Path $Books "slicer$n.xlsx") } else { Remove-Item $file }
    $results += [pscustomobject]@{ name = $case.name; refused = $refused; repaired = $repaired; slicers = $slicers; table = $table; pivot = $pivot; parts = $parts }
    Write-Host "$n $($case.name)$(if ($refused) { " (refused: $refused)" })"
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$json = $results | ConvertTo-Json -Depth 8
[IO.File]::WriteAllText($Out, $json, (New-Object Text.UTF8Encoding $false))
"$($results.Count) cases"
