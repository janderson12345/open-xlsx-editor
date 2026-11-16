# How Excel shows numbers with other decimal and thousands separators: Excel's own Options > Advanced > "Use system
# separators" turned off and the separators set (Application.DecimalSeparator / ThousandsSeparator), each value in each
# format put in a cell and its Range.Text read. The three settings changed are put back as they were before Excel
# quits (finally), so the computer's Excel is left as it was; an invisible instance, macros disabled for it only, the
# workbook closed without saving, no keys sent.
# usage: pwsh tools/excel-compare/excel-separators.ps1 test/separators-excel.json
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$values = @(1234.5, -0.5, 1234567.891, 0.000123, 1.5E+20, 45356.75, 0, 12.25)
$codes = @('General', '0.00', '#,##0', '#,##0.00', '0.0%', '0.00E+00', '#,##0.00 "kr"', '#,##0,"k"', 'mm:ss.0', '# ?/?', '"v1.5 "0.0', 'm/d/yyyy', '#,##0.00_);[Red](#,##0.00)')
$settings = @(@{ name = 'de'; dec = ','; grp = '.' }, @{ name = 'fr'; dec = ','; grp = [string][char]0xA0 }, @{ name = 'ch'; dec = '.'; grp = "'" })
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$was = @{ sys = $x.UseSystemSeparators; dec = $x.DecimalSeparator; grp = $x.ThousandsSeparator }
$answers = New-Object System.Collections.Generic.List[object]
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $ws.Columns.Item(1).ColumnWidth = 60
  # (Every format set first, with the separators as they are: Range.NumberFormat is read with the separators Excel
  # has, and refused "mm:ss.0" and "# ?/?" once they were changed.)
  $row = 0
  foreach ($code in $codes) {
    foreach ($v in $values) {
      $row++
      $c = $ws.Cells.Item($row, 1)
      $c.NumberFormat = $code
      $c.Value2 = [double]$v
    }
  }
  foreach ($s in $settings) {
    $x.UseSystemSeparators = $false
    $x.DecimalSeparator = $s.dec
    $x.ThousandsSeparator = $s.grp
    $row = 0
    foreach ($code in $codes) {
      foreach ($v in $values) {
        $row++
        $answers.Add([ordered]@{ setting = $s.name; decimal = $s.dec; group = $s.grp; code = $code; value = [double]$v; text = [string]$ws.Cells.Item($row, 1).Text })
      }
    }
  }
  $wb.Close($false)
} finally {
  $x.DecimalSeparator = $was.dec
  $x.ThousandsSeparator = $was.grp
  $x.UseSystemSeparators = $was.sys
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
$json = [ordered]@{ excel = 'Excel.Application, UseSystemSeparators off, Range.Text (United States settings otherwise)'; shown = $answers } | ConvertTo-Json -Depth 4
[IO.File]::WriteAllText((Join-Path (Get-Location) $Out), $json, (New-Object Text.UTF8Encoding $false))
"$($answers.Count) cells"
