# How Excel draws the rules: for each case, ours-<k>.xlsx (cf-write-ours.mjs, its sheet "Rules") and Excel's own
# cf-write-<k>.xlsx (excel-cf-write.ps1, its first sheet), each cell of A1:B10's DisplayFormat fill and font colour read
# from both; the cells that differ listed. Invisible, read-only, closed without saving.
# usage: pwsh tools/excel-compare/excel-cf-look.ps1 <ours folder> <Excel's folder> <count>
param([Parameter(Mandatory)][string]$Ours, [Parameter(Mandatory)][string]$Theirs, [Parameter(Mandatory)][int]$Count)
$ErrorActionPreference = 'Stop'
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
function Looks($path, $sheet) {
  $wb = $x.Workbooks.Open($path, 0, $true)
  $ws = $wb.Worksheets.Item($sheet)
  $out = @()
  foreach ($c in $ws.Range('A1:B10').Cells) { $d = $c.DisplayFormat; $out += "$($d.Interior.Color)/$($d.Font.Color)" }
  $wb.Close($false)
  return $out
}
try {
  $same = 0
  for ($k = 1; $k -le $Count; $k++) {
    $a = Looks (Join-Path (Resolve-Path $Ours).Path "ours-$k.xlsx") 'Rules'
    $b = Looks (Join-Path (Resolve-Path $Theirs).Path "cf-write-$k.xlsx") 1
    $diff = @(); for ($i = 0; $i -lt $a.Count; $i++) { if ($a[$i] -ne $b[$i]) { $diff += "cell $i ours $($a[$i]) Excel $($b[$i])" } }
    if ($diff.Count) { "case $($k): $($diff.Count) cells differ: $($diff[0..2] -join '; ')" } else { $same++ }
  }
  "$same of $Count cases drawn the same by Excel"
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
