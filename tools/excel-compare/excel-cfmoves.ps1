# Where conditional formats go when Excel inserts or deletes rows, columns or cells, or sorts (cfmove-cases.json, from
# cfmove-cases.mjs), the Excel 2010 form's (x14:conditionalFormatting in the sheet's extLst) above all: for each case
# a new workbook in an invisible Excel, Sheet1 built (made-up numbers in B2:E10, 4 in Data!A1) with data bars on
# B2:B10 (written in both forms, linked by an x14:id), 3 stars on C2:C10 (the 2010 form only), a formula rule reading
# another sheet on D2:D10 (=D2>Data!$A$1, the 2010 form) and a plain "greater than 5" on E2:E10; the step done; the
# workbook saved to %TEMP% and its rules kept as text. With -Books <folder>, the editor's results (<folder>\<n>.xlsx,
# from test/cfmoves.test.js) are opened instead and read back the same way, `repaired` saying whether Excel had to
# repair one (a repair log, error*.xml, appeared in %TEMP%). No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-cfmoves.ps1 <cfmove-cases.json> <out.json> [-Books <folder>]
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out, [string]$Books = '')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$spec = Get-Content -Raw $Cases | ConvertFrom-Json
$x = New-Object -ComObject Excel.Application
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
    } else {
      $wb = $x.Workbooks.Add()
      while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
      $ws = $wb.Worksheets.Item(1)
      $data = $wb.Worksheets.Item(2); $data.Name = 'Data'
      $data.Range('A1').Value2 = [double]4
      for ($r = 2; $r -le 10; $r++) { for ($c = 2; $c -le 5; $c++) { $ws.Cells.Item($r, $c).Value2 = [double]((($r * 3 + $c * 5) % 9) + 1) } }
      [void]$ws.Activate()
      $db = $ws.Range('B2:B10').FormatConditions.AddDatabar()
      $is = $ws.Range('C2:C10').FormatConditions.AddIconSetCondition(); $is.IconSet = $wb.IconSets.Item(18)
      $fx = $ws.Range('D2:D10').FormatConditions.Add(2, [Type]::Missing, '=D2>Data!$A$1'); $fx.Interior.Color = 0xCEC7FF
      $gt = $ws.Range('E2:E10').FormatConditions.Add(1, 5, '5'); $gt.Interior.Color = 0x9CEBFF
      $s = $case.step
      $on = if ($s.sheet -eq 'Data') { $data } else { $ws }
      try {
        $rg = $on.Range($s.range)
        $shift = if ($s.shift -eq 'down') { -4121 } elseif ($s.shift -eq 'right') { -4161 } elseif ($s.shift -eq 'up') { -4162 } elseif ($s.shift -eq 'left') { -4159 } else { $null }
        if ($s.do -eq 'insert') { if ($null -ne $shift) { [void]$rg.Insert($shift) } else { [void]$rg.Insert() } }
        if ($s.do -eq 'delete') { if ($null -ne $shift) { [void]$rg.Delete($shift) } else { [void]$rg.Delete() } }
        if ($s.do -eq 'sort') {
          $sort = $on.Sort; $sort.SortFields.Clear()
          [void]$sort.SortFields.Add($on.Range($s.key), 0, $(if ($s.desc) { 2 } else { 1 }))
          $sort.SetRange($rg); $sort.Header = 2; $sort.Apply()
        }
      } catch { $refused = ($_.Exception.InnerException ?? $_.Exception).Message }
    }
    $file = Join-Path $env:TEMP "oxe-cfm$n.xlsx"
    if (Test-Path $file) { Remove-Item $file }
    $wb.SaveAs($file, 51)
    $wb.Close($false)
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    try { $sr = New-Object IO.StreamReader($zip.GetEntry('xl/worksheets/sheet1.xml').Open()); $t = $sr.ReadToEnd(); $sr.Close() } finally { $zip.Dispose() }
    Remove-Item $file
    $results += [pscustomobject]@{
      name = $case.name; refused = $refused; repaired = $repaired
      main = @([regex]::Matches($t, '<conditionalFormatting\b.*?</conditionalFormatting>', 'Singleline') | ForEach-Object { $_.Value })
      ext = [regex]::Match($t, '<x14:conditionalFormattings>.*?</x14:conditionalFormattings>', 'Singleline').Value
    }
    Write-Host "$n $($case.name)$(if ($refused) { " (refused: $refused)" })"
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
[IO.File]::WriteAllText($Out, ($results | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding $false))
"$($results.Count) cases"
