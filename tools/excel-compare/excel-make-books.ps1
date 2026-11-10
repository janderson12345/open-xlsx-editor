# Has Excel make workbooks full of what editors tend to lose, for the save-fidelity benchmark (fidelity.mjs): charts,
# a pivot table and its cache, a table, sparklines, data bars and icon sets, notes and threaded comments, hyperlinks,
# defined names, data validation, conditional formats, freeze panes, grouped rows, a protected sheet, print settings,
# pictures of shapes. Written to a folder outside the repository (they hold Excel's own theme and parts); nothing is
# sent anywhere. An invisible instance, macros disabled for it only.
# usage: pwsh tools/excel-compare/excel-make-books.ps1 <folder>
param([Parameter(Mandatory)][string]$Folder)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $Folder | Out-Null
$Folder = (Resolve-Path $Folder).Path # (Excel saves a relative path against its own folder.)
$m = [Type]::Missing
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3 # msoAutomationSecurityForceDisable: this instance runs no macros
function Add-Rows($ws) {
  $ws.Range('A1:D1').Value2 = [object[]]@('Region', 'Month', 'Sales', 'Units')
  $regions = 'North', 'South', 'East', 'West'
  for ($i = 0; $i -lt 24; $i++) {
    $r = $i + 2
    $ws.Cells.Item($r, 1).Value2 = $regions[$i % 4]
    $ws.Cells.Item($r, 2).Formula = "=DATE(2026,$([int][math]::Floor($i / 4) + 1),1)"
    $ws.Cells.Item($r, 3).Value2 = [double](100 + ($i * 37) % 250)
    $ws.Cells.Item($r, 4).Value2 = [double](5 + ($i * 11) % 40)
  }
  $ws.Range('B2:B25').NumberFormat = 'mmm yyyy'
  $ws.Range('C2:C25').NumberFormat = '$#,##0.00'
  $ws.Range('A1:D1').Font.Bold = $true
}
$made = @()
try {
  # 1. Charts of three kinds over a sheet of data, a total row and a defined name.
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = 'Sales'; Add-Rows $ws
  $ws.Range('C26').Formula = '=SUM(C2:C25)'
  [void]$wb.Names.Add('SalesTotal', '=Sales!$C$26')
  $kinds = 51, 4, 5 # column, line, pie
  $k = 0
  foreach ($kind in $kinds) {
    $co = $ws.ChartObjects().Add(320 + $k * 20, 20 + $k * 220, 360, 200)
    $co.Chart.ChartType = $kind
    $co.Chart.SetSourceData($ws.Range('A1:A9,C1:C9'))
    $k++
  }
  $wb.SaveAs((Join-Path $Folder 'charts.xlsx'), 51); $wb.Close($false); $made += 'charts.xlsx'

  # 2. A pivot table, its cache, and a table (ListObject) with a total row.
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = 'Data'; Add-Rows $ws
  $lo = $ws.ListObjects.Add(1, $ws.Range('A1:D25'), $m, 1)
  $lo.Name = 'SalesTable'; $lo.ShowTotals = $true
  $pws = $wb.Worksheets.Add(); $pws.Name = 'Pivot'
  $pc = $wb.PivotCaches().Create(1, 'SalesTable')
  $pt = $pc.CreatePivotTable($pws.Range('A3'), 'SalesPivot')
  $pt.PivotFields('Region').Orientation = 1
  [void]$pt.AddDataField($pt.PivotFields('Sales'), 'Sum of Sales', -4157)
  $wb.SaveAs((Join-Path $Folder 'pivot.xlsx'), 51); $wb.Close($false); $made += 'pivot.xlsx'

  # 3. Conditional formats of every family: values, a colour scale, data bars, an icon set; sparklines; validation.
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = 'Formats'; Add-Rows $ws
  [void]$ws.Range('C2:C25').FormatConditions.AddColorScale(3)
  [void]$ws.Range('D2:D25').FormatConditions.AddDatabar()
  [void]$ws.Range('E2:E25').FormatConditions.AddIconSetCondition()
  for ($r = 2; $r -le 25; $r++) { $ws.Cells.Item($r, 5).Formula = "=D$r*2" }
  $c = $ws.Range('C2:C25').FormatConditions.Add(1, 5, '300'); $c.Font.Bold = $true
  [void]$ws.Range('F2:F7').SparklineGroups.Add(1, 'C2:D7')
  $v = $ws.Range('A2:A25').Validation; $v.Delete(); [void]$v.Add(3, 1, 1, 'North,South,East,West')
  $v.InputTitle = 'Region'; $v.InputMessage = 'Pick a region'
  $wb.SaveAs((Join-Path $Folder 'formats.xlsx'), 51); $wb.Close($false); $made += 'formats.xlsx'

  # 4. Notes, a threaded comment, hyperlinks, shapes, freeze panes, grouped rows, a protected sheet, print settings.
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = 'Review'; Add-Rows $ws
  [void]$ws.Range('C2').AddComment('Check this figure')
  try { [void]$ws.Range('C3').AddCommentThreaded('Is this right?') } catch { }
  [void]$ws.Hyperlinks.Add($ws.Range('A27'), 'https://example.com/report', $m, 'Open the report', 'Report')
  [void]$ws.Hyperlinks.Add($ws.Range('A28'), '', 'Sales!A1', $m, 'Back to the top')
  [void]$ws.Shapes.AddShape(1, 300, 30, 120, 60)
  [void]$ws.Shapes.AddTextbox(1, 300, 110, 160, 40)
  [void]$ws.Activate(); $x.ActiveWindow.FreezePanes = $false; [void]$ws.Range('B2').Select(); $x.ActiveWindow.FreezePanes = $true
  [void]$ws.Range('A6:A9').Rows.Group()
  $ws.PageSetup.Orientation = 2; $ws.PageSetup.PrintTitleRows = '$1:$1'; $ws.PageSetup.CenterFooter = 'Page &P of &N'
  $ps = $wb.Worksheets.Add(); $ps.Name = 'Locked'; $ps.Range('A1').Value2 = 'Protected'; $ps.Protect('', $true, $true)
  $wb.SaveAs((Join-Path $Folder 'review.xlsx'), 51); $wb.Close($false); $made += 'review.xlsx'
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
"made: $($made -join ', ')"
