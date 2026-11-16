# Has Excel copy a formatted range of made-up values and saves the HTML it puts on the clipboard ("HTML Format"),
# to see how it writes fonts, colours, fills, borders, alignment, number formats (mso-number-format), merged cells
# (colspan, rowspan) and column widths. A new workbook in an invisible instance (macros disabled for it only),
# closed without saving. Range.Copy uses the Windows clipboard: its text is put back afterwards.
# usage: pwsh -STA tools/excel-compare/excel-clip-html.ps1 <folder outside the repo>
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
$keep = [System.Windows.Forms.Clipboard]::GetText()
New-Item -ItemType Directory -Force $Out | Out-Null
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
try {
  $b = $x.Workbooks.Add()
  $s = $b.Worksheets.Item(1)
  $s.Range('A1').Value2 = 'Item'; $s.Range('B1').Value2 = 'Qty'; $s.Range('C1').Value2 = 'Price'; $s.Range('D1').Value2 = 'Share'
  $s.Range('A1:D1').Font.Bold = $true
  $s.Range('A1:D1').Interior.Color = 0x00C0FF # RGB(255,192,0) as BGR
  $s.Range('A2').Value2 = 'Pears'; $s.Range('B2').Value2 = 1234.5; $s.Range('C2').Value2 = 3.25; $s.Range('D2').Value2 = 0.125
  $s.Range('B2').NumberFormat = '#,##0.00'
  $s.Range('C2').NumberFormat = '$#,##0.00'
  $s.Range('D2').NumberFormat = '0.0%'
  $s.Range('A3').Value2 = 'Plums, long text'; $s.Range('A3').WrapText = $true
  $s.Range('B3').Value2 = 45000; $s.Range('B3').NumberFormat = 'yyyy-mm-dd'
  $s.Range('C3').Formula = '=B2*C2'
  $s.Range('D3').Value2 = '007'; $s.Range('D3').NumberFormat = '@'
  $s.Range('A2').Font.Italic = $true; $s.Range('A2').Font.Name = 'Georgia'; $s.Range('A2').Font.Size = 14
  $s.Range('A2').Font.Color = 0x0000FF # red
  $s.Range('B2').Font.Underline = 2; $s.Range('C2').Font.Strikethrough = $true
  $s.Range('B2:B3').HorizontalAlignment = -4108 # center
  $s.Range('A3').VerticalAlignment = -4160 # top
  $s.Range('A4:B4').Merge(); $s.Range('A4').Value2 = 'Merged across'
  $s.Range('C4:C5').Merge(); $s.Range('C4').Value2 = 'Down'
  $s.Range('A5').Value2 = -7; $s.Range('A5').NumberFormat = '0;[Red](0)'
  $s.Range('B5').Value2 = $true
  $s.Range('D4').Value2 = 'indented'; $s.Range('D4').IndentLevel = 2
  $bd = $s.Range('A2:B3').Borders
  foreach ($i in 7, 8, 9, 10) { $bd.Item($i).LineStyle = 1; $bd.Item($i).Weight = 2 }
  $s.Range('D2').Borders.Item(9).LineStyle = -4119 # double bottom
  $s.Range('D2').Borders.Item(9).Color = 0xFF0000 # blue
  $s.Columns.Item(1).ColumnWidth = 20
  $s.Columns.Item(3).ColumnWidth = 12
  $s.Rows.Item(3).RowHeight = 30
  [void]$s.Range('A1:D5').Copy()
  $html = [System.Windows.Forms.Clipboard]::GetData('HTML Format')
  if ($html -is [System.IO.MemoryStream]) { $html = [System.Text.Encoding]::UTF8.GetString($html.ToArray()) }
  $text = [System.Windows.Forms.Clipboard]::GetText()
  $x.CutCopyMode = 0
  [System.IO.File]::WriteAllText((Join-Path $Out 'excel-clip.html'), [string]$html)
  [System.IO.File]::WriteAllText((Join-Path $Out 'excel-clip.txt'), [string]$text)
  $b.Close($false)
} finally {
  $x.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  if ($keep) { [System.Windows.Forms.Clipboard]::SetText($keep) }
}
