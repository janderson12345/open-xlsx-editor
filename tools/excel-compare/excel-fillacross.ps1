# Has Excel fill a range across grouped sheets (Sheets(Array(...)).FillAcrossSheets with xlFillWithAll,
# xlFillWithContents and xlFillWithFormats, as Home > Fill > Across Worksheets does) and reads back what each sheet's
# cells hold. Each case is a new workbook of made-up values in an invisible instance (macros disabled for it only),
# closed without saving. Writes JSON: { case: { sheet: { cell: {...} }, widths, merges } }.
# usage: pwsh tools/excel-compare/excel-fillacross.ps1 <out.json>
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$types = [ordered]@{ all = -4104; contents = 2; formats = -4122 }
$res = [ordered]@{}
$cells = 'A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3', 'A4', 'B4', 'C4', 'D1', 'A5'
$read = { param($s)
  $o = [ordered]@{}
  foreach ($a in $cells) {
    $c = $s.Range($a)
    $o[$a] = [ordered]@{
      formula = [string]$c.Formula; value = $c.Value2; numFmt = [string]$c.NumberFormat
      bold = ($c.Font.Bold -eq $true); italic = ($c.Font.Italic -eq $true)
      fill = [int]$c.Interior.ColorIndex; bottom = [int]$c.Borders.Item(9).LineStyle
      merged = ($c.MergeCells -eq $true); note = ($null -ne $c.Comment)
      links = [int]$c.Hyperlinks.Count; cf = [int]$c.FormatConditions.Count
      dv = $(try { [int]$c.Validation.Type } catch { -1 })
    }
  }
  $o['_widthA'] = [double]$s.Columns.Item(1).ColumnWidth
  $o['_heightR1'] = [double]$s.Rows.Item(1).RowHeight
  $o
}
try {
  foreach ($k in $types.Keys) {
    $wb = $x.Workbooks.Add()
    while ($wb.Worksheets.Count -lt 3) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $s1 = $wb.Worksheets.Item(1)
    # The source: a bold number with a format, a fill and a border, formulas (relative, absolute, to another sheet),
    # text, a blank cell, a merge, a note, a hyperlink, validation, a conditional format, a wide column, a tall row.
    $s1.Range('A1').Value2 = 1234.5; $s1.Range('A1').NumberFormat = '#,##0.00'; $s1.Range('A1').Font.Bold = $true
    $s1.Range('A1').Interior.ColorIndex = 6; $s1.Range('A1').Borders.Item(9).LineStyle = 1
    $s1.Range('B1').Formula = '=A1*2'
    $s1.Range('C1').Formula = '=$A$1+Sheet2!A1'
    $s1.Range('A2').Value2 = 'text'
    # B2 left blank
    $s1.Range('C2').Formula = '=SUM(A1:B1)'
    $s1.Range('A3').Value2 = 'noted'; [void]$s1.Range('A3').AddComment('a note')
    $s1.Range('B3').Value2 = 'link'; [void]$s1.Hyperlinks.Add($s1.Range('B3'), 'https://example.com/')
    $s1.Range('C3').Value2 = 5
    $v = $s1.Range('C3').Validation; $v.Add(1, 1, 1, '1', '10')
    [void]$s1.Range('A4').FormatConditions.Add(1, 5, '=0')
    $s1.Range('A4').Value2 = 7
    $s1.Range('B4:C4').Merge(); $s1.Range('B4').Value2 = 'merged'
    $s1.Columns.Item(1).ColumnWidth = 20
    $s1.Rows.Item(1).RowHeight = 30
    # Each other sheet: something already in the range (italic), and a cell outside it.
    foreach ($i in 2, 3) {
      $s = $wb.Worksheets.Item($i)
      foreach ($a in 'A1', 'B2', 'C3', 'B4') { $s.Range($a).Value2 = "old$i"; $s.Range($a).Font.Italic = $true }
      $s.Range('D1').Value2 = "outside$i"
      $s.Range('A5').Value2 = "below$i"
    }
    $wb.Worksheets.Item(1).Activate()
    $grp = $wb.Sheets.Item([object[]]@('Sheet1', 'Sheet2', 'Sheet3'))
    $grp.FillAcrossSheets($s1.Range('A1:C4'), $types[$k])
    $o = [ordered]@{}
    foreach ($i in 1..3) { $o["Sheet$i"] = & $read $wb.Worksheets.Item($i) }
    $res[$k] = $o
    $wb.Close($false)
  }
} finally {
  $x.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
$res | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 $Out
