# Where sparklines, data validation rules, Circle Invalid Data's ovals, notes and threaded comments go when Excel
# inserts or deletes rows, columns or cells, or sorts (refmove-cases.json, from refmove-cases.mjs): for each case a new
# workbook in an invisible Excel, the sheet built (made-up numbers), the step done, the ovals read from the sheet's
# shapes, the notes and threads from COM, then the workbook saved to %TEMP% and its sheets' XML for these features kept
# (the persons part and the notes' authors are not read: they carry Excel's user name). No keys sent, no setting changed.
# With -Books <folder>, each case's workbook is the editor's own result (<folder>\<case number>.xlsx, from
# test/refmoves.test.js) opened instead, read back the same way; `repaired` says whether Excel had to repair it (a
# repair log, error*.xml, appeared in %TEMP%).
# usage: pwsh tools/excel-compare/excel-refmoves.ps1 <refmove-cases.json> <out.json> [-Books <folder>]
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out, [string]$Books = '')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);'
$spec = Get-Content -Raw $Cases | ConvertFrom-Json
$partsOf = { param($file)
  $zip = [IO.Compression.ZipFile]::OpenRead($file)
  $read = { param($name) $e = $zip.GetEntry($name); if (-not $e) { return $null }; $s = New-Object IO.StreamReader($e.Open()); $t = $s.ReadToEnd(); $s.Close(); $t }
  try {
    $sheet = & $read 'xl/worksheets/sheet1.xml'
    $data = & $read 'xl/worksheets/sheet2.xml'
    $names = @($zip.Entries | ForEach-Object { $_.FullName })
    $comments = $names | Where-Object { $_ -match '^xl/comments\d*\.xml$' } | Select-Object -First 1
    $threads = $names | Where-Object { $_ -match '^xl/threadedComments/' } | Select-Object -First 1
    $vml = $names | Where-Object { $_ -match '^xl/drawings/vmlDrawing\d*\.vml$' } | Select-Object -First 1
    $c = if ($comments) { & $read $comments } else { '' }
    $t = if ($threads) { & $read $threads } else { '' }
    $v = if ($vml) { & $read $vml } else { '' }
    [ordered]@{
      sparklines = [regex]::Match($sheet, '<x14:sparklineGroups.*?</x14:sparklineGroups>', 'Singleline').Value
      dataValidations = [regex]::Match($sheet, '<dataValidations.*?</dataValidations>', 'Singleline').Value
      x14DataValidations = [regex]::Match($sheet, '<x14:dataValidations.*?</x14:dataValidations>', 'Singleline').Value
      conditional = @([regex]::Matches($sheet, '<conditionalFormatting.*?</conditionalFormatting>', 'Singleline') | ForEach-Object { $_.Value })
      dataSheetExt = [regex]::Match($data, '<extLst>.*?</extLst>', 'Singleline').Value
      commentRefs = @([regex]::Matches($c, '<comment ref="([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
      threadRefs = @([regex]::Matches($t, '<threadedComment ref="([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
      vmlNotes = @([regex]::Matches($v, '<x:Anchor>([^<]*)</x:Anchor>[\s\S]*?<x:Row>(\d+)</x:Row>[\s\S]*?<x:Column>(\d+)</x:Column>') | ForEach-Object { "$($_.Groups[2].Value),$($_.Groups[3].Value): $($_.Groups[1].Value.Trim())" })
    }
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
    if ($Books) {
      $logs = @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue).Count
      $wb = $x.Workbooks.Open((Join-Path (Resolve-Path $Books).Path "$n.xlsx"), 0, $true)
      $repaired = @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue).Count -ne $logs
      $ws = $wb.Worksheets.Item(1)
    } else {
    $wb = $x.Workbooks.Add()
    while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $ws = $wb.Worksheets.Item(1)
    $data = $wb.Worksheets.Item(2); $data.Name = 'Data'
    # (Numbers and text set from two call sites: PowerShell's COM binder keeps the first value type a call site saw and
    # refuses the other, "Unable to cast ... String to Double".)
    $put = { param($on, $cells)
      foreach ($cell in $cells.PSObject.Properties) {
        $v = $cell.Value
        if ($v -is [string]) { $on.Range($cell.Name).Formula = [string]$v } else { $on.Range($cell.Name).Value2 = [double]$v }
      }
    }
    & $put $ws $spec.sheet.values
    & $put $data $spec.data.values
    [void]$ws.Activate()
    foreach ($s in $spec.sheet.sparklines) { [void]$ws.Range($s.at).SparklineGroups.Add($s.type, $s.data) }
    foreach ($v in $spec.sheet.validations) {
      $val = $ws.Range($v.at).Validation
      $val.Delete()
      if ($null -ne $v.formula2) { $val.Add($v.type, 1, $v.op, $v.formula1, $v.formula2) } else { $val.Add($v.type, 1, $v.op, $v.formula1) }
    }
    foreach ($c in $spec.sheet.notes) { [void]$ws.Range($c).AddComment("note $c") }
    foreach ($c in $spec.sheet.threads) { [void]$ws.Range($c).AddCommentThreaded("thread $c") }
    [void]$ws.Activate()
    $ws.CircleInvalid()
    $s = $case.step
    $on = if ($s.sheet -eq 'Data') { $data } else { $ws }
    try {
      $r = $on.Range($s.range)
      $shift = if ($s.shift -eq 'down') { -4121 } elseif ($s.shift -eq 'right') { -4161 } elseif ($s.shift -eq 'up') { -4162 } elseif ($s.shift -eq 'left') { -4159 } else { $null }
      if ($s.do -eq 'insert') { if ($null -ne $shift) { [void]$r.Insert($shift) } else { [void]$r.Insert() } }
      if ($s.do -eq 'delete') { if ($null -ne $shift) { [void]$r.Delete($shift) } else { [void]$r.Delete() } }
      if ($s.do -eq 'sort') {
        $sort = $on.Sort
        $sort.SortFields.Clear()
        [void]$sort.SortFields.Add($on.Range($s.key), 0, $(if ($s.desc) { 2 } else { 1 }))
        $sort.SetRange($r)
        $sort.Header = 2
        $sort.Orientation = $(if ($s.leftToRight) { 2 } else { 1 })
        $sort.Apply()
      }
    } catch { $refused = ($_.Exception.InnerException ?? $_.Exception).Message }
    }
    # (The ovals have no TopLeftCell over COM: the cell is the one whose box holds the oval's centre.)
    $ovals = @(); foreach ($p in $ws.Shapes) {
      if ($p.Type -ne 1 -and $p.Type -ne 5) { continue }
      $cx = $p.Left + $p.Width / 2; $cy = $p.Top + $p.Height / 2
      $col = 1; while ($ws.Cells.Item(1, $col + 1).Left -le $cx) { $col++ }
      $row = 1; while ($ws.Cells.Item($row + 1, 1).Top -le $cy) { $row++ }
      $ovals += "$($ws.Cells.Item($row, $col).Address($false, $false)) type $($p.Type) $($p.Left),$($p.Top) $($p.Width)x$($p.Height)"
    }
    $notes = @(); foreach ($c in $ws.Comments) { $notes += $c.Parent.Address($false, $false) }
    $threads = @(); foreach ($c in $ws.CommentsThreaded) { $threads += $c.Parent.Address($false, $false) }
    $sparks = @(); $all = $ws.Cells.SparklineGroups
    for ($k = 1; $k -le $all.Count; $k++) { $g = $all.Item($k); for ($j = 1; $j -le $g.Count; $j++) { $sp = $g.Item($j); $sparks += "$($sp.Location.Address($false, $false)) <- $($sp.SourceData)" } }
    $file = Join-Path $env:TEMP "oxe-rm$n.xlsx"
    if (Test-Path $file) { Remove-Item $file }
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $xml = & $partsOf $file
    Remove-Item $file
    $results += [pscustomobject]@{ name = $case.name; refused = $refused; repaired = $repaired; ovals = $ovals; notes = $notes; threads = $threads; sparklines = $sparks; xml = $xml }
    Write-Host "$n $($case.name)$(if ($refused) { " (refused: $refused)" })"
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$json = $results | ConvertTo-Json -Depth 8
[IO.File]::WriteAllText($Out, $json, (New-Object Text.UTF8Encoding $false))
"$($results.Count) cases"
