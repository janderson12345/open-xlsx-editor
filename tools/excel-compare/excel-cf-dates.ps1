# What Excel writes and draws for Home > Conditional Formatting > Highlight Cells Rules > A Date Occurring: in an
# invisible Excel, a new workbook whose columns A to J each hold the dates from 45 days before today to 45 days after,
# one of the ten periods added to each column (Light Red Fill with Dark Red Text), the workbook saved to a folder
# outside the repository and its conditionalFormatting read back, and the cells Excel highlights read from their
# DisplayFormat. No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-cf-dates.ps1 <out.json> <folder for the saved workbook>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
trap { Write-Host "LINE $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; break }
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$m = [Type]::Missing
# XlTimePeriods, in the order of the menu's list.
$PERIODS = [ordered]@{ yesterday = 1; today = 0; tomorrow = 6; last7Days = 2; lastWeek = 4; thisWeek = 3; nextWeek = 7; lastMonth = 5; thisMonth = 9; nextMonth = 8 }
$N = 91
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
try {
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $today = [int]$x.WorksheetFunction.Floor([double]$ws.Evaluate('TODAY()'), 1)
  $grid = New-Object 'object[,]' $N, $PERIODS.Count
  for ($i = 0; $i -lt $N; $i++) { for ($j = 0; $j -lt $PERIODS.Count; $j++) { $grid[$i, $j] = $today - 45 + $i } }
  $all = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($N, $PERIODS.Count))
  $all.Value2 = $grid
  $all.NumberFormat = 'yyyy-mm-dd'
  $j = 0
  foreach ($p in $PERIODS.Keys) {
    $j++
    $col = $ws.Range($ws.Cells.Item(1, $j), $ws.Cells.Item($N, $j))
    $fc = $col.FormatConditions.Add(11, $m, $m, $m, $m, $m, $PERIODS[$p])
    $fc.Interior.Color = 13551615; $fc.Font.Color = 393372
  }
  $hits = [ordered]@{}
  $j = 0
  foreach ($p in $PERIODS.Keys) {
    $j++
    $hits[$p] = @(for ($i = 0; $i -lt $N; $i++) { if ($ws.Cells.Item($i + 1, $j).DisplayFormat.Interior.Color -eq 13551615) { $i - 45 } })
  }
  $file = Join-Path $Books 'cf-dates.xlsx'
  if (Test-Path $file) { Remove-Item $file }
  $wb.SaveAs($file, 51)
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x); [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
$zip = [IO.Compression.ZipFile]::OpenRead($file)
$s = New-Object IO.StreamReader($zip.GetEntry('xl/worksheets/sheet1.xml').Open()); $sheet = $s.ReadToEnd(); $s.Close()
$s = New-Object IO.StreamReader($zip.GetEntry('xl/styles.xml').Open()); $styles = $s.ReadToEnd(); $s.Close()
$zip.Dispose()
$cf = @([regex]::Matches($sheet, '<conditionalFormatting\b[\s\S]*?</conditionalFormatting>') | ForEach-Object { $_.Value })
$dxfs = ([regex]::Match($styles, '<dxfs\b[\s\S]*?</dxfs>|<dxfs\b[^>]*/>')).Value
[pscustomobject]@{ today = $today; weekday = [int]([datetime]"1899-12-30").AddDays($today).DayOfWeek + 1; periods = $hits; conditionalFormatting = $cf; dxfs = $dxfs } | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 $Out
"$($PERIODS.Count) periods, today $today"
