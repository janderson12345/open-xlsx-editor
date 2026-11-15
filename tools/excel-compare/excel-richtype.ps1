# Where Excel puts text added inside a cell of rich text: which run's font the new characters take. A cell is given
# runs (Characters(start, length).Font set), then Characters(start, length).Insert(text) adds or replaces text, and
# each character's font is read back (bold, italic, size, colour). Typing can't be driven over COM without
# keystrokes, so this is the object model's own insert, the nearest Excel offers to typing at a caret.
# Starts its own hidden Excel and quits it; the owner's Excel is not touched.
# usage: powershell.exe -NoProfile -File tools/excel-compare/excel-richtype.ps1 test/richtype-excel.json
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
# (Each case: the text, its formatted runs [start, length, what], then Insert(text) at [start, length].)
$cases = @(
  @{ name = 'inside a bold run'; text = 'plainBOLDplain'; runs = ,@(6, 4, 'b'); at = @(8, 0); add = 'x' },
  @{ name = 'just after a bold run (bold before the caret)'; text = 'plainBOLDplain'; runs = ,@(6, 4, 'b'); at = @(10, 0); add = 'x' },
  @{ name = 'just before a bold run (plain before the caret)'; text = 'plainBOLDplain'; runs = ,@(6, 4, 'b'); at = @(6, 0); add = 'x' },
  @{ name = 'at the start, the first run bold'; text = 'BOLDplain'; runs = ,@(1, 4, 'b'); at = @(1, 0); add = 'x' },
  @{ name = 'at the end, the last run bold'; text = 'plainBOLD'; runs = ,@(6, 4, 'b'); at = @(10, 0); add = 'x' },
  @{ name = 'after a red 16pt run'; text = 'plainREDplain'; runs = ,@(6, 3, 'red16'); at = @(9, 0); add = 'xy' },
  @{ name = 'before a red 16pt run'; text = 'plainREDplain'; runs = ,@(6, 3, 'red16'); at = @(6, 0); add = 'xy' },
  @{ name = 'replacing plain and bold (selection starts plain)'; text = 'plainBOLDplain'; runs = ,@(6, 4, 'b'); at = @(4, 4); add = 'xy' },
  @{ name = 'replacing bold and plain (selection starts bold)'; text = 'plainBOLDplain'; runs = ,@(6, 4, 'b'); at = @(8, 4); add = 'xy' },
  @{ name = 'replacing all of a bold run'; text = 'plainBOLDplain'; runs = ,@(6, 4, 'b'); at = @(6, 4); add = 'xy' }
)
Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);'
$x = New-Object -ComObject Excel.Application
# (This Excel's own process, from its window handle, so only it is stopped if Quit leaves it running.)
$mine = [uint32]0
[void][W.U]::GetWindowThreadProcessId([IntPtr]$x.Hwnd, [ref]$mine)
Write-Host "Excel PID $mine"
$x.Visible = $false
$x.DisplayAlerts = $false
$answers = New-Object System.Collections.Generic.List[object]
try {
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  foreach ($case in $cases) {
    $c = $ws.Range('A1')
    $c.Clear()
    $c.Value2 = $case.text
    foreach ($r in $case.runs) {
      $f = $c.Characters($r[0], $r[1]).Font
      if ($r[2] -eq 'b') { $f.Bold = $true }
      if ($r[2] -eq 'red16') { $f.Color = 255; $f.Size = 16 }
    }
    [void]$c.Characters($case.at[0], $case.at[1]).Insert($case.add)
    $text = [string]$c.Value2
    $chars = @()
    for ($i = 1; $i -le $text.Length; $i++) {
      $f = $c.Characters($i, 1).Font
      $chars += [ordered]@{ ch = $text.Substring($i - 1, 1); b = [bool]$f.Bold; size = [double]$f.Size; color = [double]$f.Color }
    }
    $answers.Add([ordered]@{ name = $case.name; text = $case.text; runs = $case.runs; at = $case.at; add = $case.add; result = $text; chars = $chars })
  }
  $wb.Close($false)
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  Start-Sleep -Milliseconds 500
  $p = Get-Process -Id $mine -ErrorAction SilentlyContinue
  if ($p -and $p.MainWindowHandle -eq 0) { Stop-Process -Id $mine -Force }
}
$json = ConvertTo-Json -InputObject $answers.ToArray() -Depth 6
[IO.File]::WriteAllText((Join-Path (Get-Location) $Out), $json + "`n", (New-Object Text.UTF8Encoding $false))
