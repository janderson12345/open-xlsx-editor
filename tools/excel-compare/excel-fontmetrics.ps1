# The row height and maximum digit width Excel gives each font and size: in a new workbook the Normal style is given
# the font and size (and bold), then the sheet's StandardHeight (points) is read, and, regular only, column A's
# Width (points) at ColumnWidth 1 and 10, the digit width being a ninth of the difference in pixels. Sizes 6 to 36 in
# half points. Excel's heights come from Windows' rendering of the font at the size in whole pixels, which no formula
# of the font's metrics gives (rowheights.js, fonts.js); this table is what the editor looks them up in.
# An invisible instance, blank workbooks only (no one's file), macros disabled for it only, no keys sent.
# usage: pwsh tools/excel-compare/excel-fontmetrics.ps1 <out.json>
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$fonts = 'Calibri', 'Calibri Light', 'Arial', 'Arial Narrow', 'Arial Black', 'Times New Roman', 'Aptos Narrow', 'Aptos',
  'Tahoma', 'Book Antiqua', 'Cambria', 'Verdana', 'Lato', 'Trebuchet MS', 'Courier New', 'Segoe UI', 'Georgia',
  'Century Gothic', 'Garamond', 'Consolas', 'Helv', 'Geneva', 'MS Sans Serif', 'Tms Rmn', 'Courier'
$sizes = for ($s = 6; $s -le 36; $s += 0.5) { $s }
# (A reading Excel gives no number for, such as DBNull, is recorded as null.)
function num($v) { if ($v -is [double] -or $v -is [int] -or $v -is [single] -or $v -is [decimal]) { [double]$v } else { $null } }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$x.AutomationSecurity = 3
$res = [ordered]@{}
try {
  foreach ($f in $fonts) {
    $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
    $font = $wb.Styles.Item('Normal').Font
    $font.Name = $f
    $row = [ordered]@{}
    foreach ($s in $sizes) {
      try {
        # (A plain double: a size from the loop is a PSObject, which COM can't take.)
        $font.Size = [double]$s
        $font.Bold = $false
        $h = num $ws.StandardHeight
        $col = $ws.Columns.Item(1)
        $col.ColumnWidth = 1; $w1 = num $col.Width
        $col.ColumnWidth = 10; $w10 = num $col.Width
        $font.Bold = $true
        $hb = num $ws.StandardHeight
        $row["$s"] = @($h, $hb, $w1, $w10)
      } catch {
        Write-Output "$f $s : $($_.InvocationInfo.Line.Trim()) :: $($_.Exception.Message)"
        $row["$s"] = $null
      }
    }
    $res[$f] = [ordered]@{ face = $font.Name; sizes = $row }
    $wb.Close($false)
  }
  $json = [ordered]@{ excel = $x.Version; fonts = $res } | ConvertTo-Json -Depth 5 -Compress
  [IO.File]::WriteAllText($Out, $json + "`n", (New-Object Text.UTF8Encoding $false))
  Write-Output "$($fonts.Count) fonts, $($sizes.Count) sizes"
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
