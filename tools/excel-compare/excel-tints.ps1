# Has Excel say what colour a tint gives: for each base colour (the fixtures' theme colours, the 2023 Office theme's
# accents, greys and a few others) and each tint (Excel's ribbon's lighter 80/60/40% and darker 25/50%, and the table
# styles' 15, 35 and 45%), a cell's Interior.Color set to the base and TintAndShade to the tint, and Interior.Color read
# back as RRGGBB. An invisible instance, macros disabled for it only; nothing saved, no setting changed, no keys sent.
# usage: pwsh tools/excel-compare/excel-tints.ps1 test/tints-excel.json
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$bases = @('000000', 'FFFFFF', '1F3864', 'E7E6E6', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47',
  '0E2841', 'E8E8E8', '156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E', '808080', 'C00000', '7030A0', '00B0F0')
$tints = @(0.8, 0.6, 0.4, 0.25, 0.15, 0.45, -0.15, -0.25, -0.35, -0.5)
$hex = { param($v) $i = [int64]$v; '{0:X2}{1:X2}{2:X2}' -f ($i -band 255), (($i -shr 8) -band 255), (($i -shr 16) -band 255) }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$answers = [ordered]@{}
try {
  $wb = $x.Workbooks.Add()
  $c = $wb.Worksheets.Item(1).Range('B2')
  foreach ($b in $bases) {
    $bgr = [Convert]::ToInt32($b.Substring(4, 2) + $b.Substring(2, 2) + $b.Substring(0, 2), 16)
    $row = [ordered]@{}
    foreach ($t in $tints) {
      $c.Interior.Color = $bgr
      $c.Interior.TintAndShade = $t
      $row["$t"] = & $hex $c.Interior.Color
    }
    $answers[$b] = $row
  }
  $wb.Close($false)
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
[IO.File]::WriteAllText(([IO.Path]::GetFullPath($Out, (Get-Location).Path)), ($answers | ConvertTo-Json -Depth 4) + "`n", (New-Object Text.UTF8Encoding($false)))
