# The colours each of Excel's 36 sparkline styles (the Sparkline tab's style gallery) writes, for sparklinestyles.js
# and docs/SPARKLINES.md. A style can only be applied from the ribbon (a sparkline group has no Style property in COM
# and the gallery isn't a command ExecuteMso can run), so this one Excel of its own is made visible and each gallery
# item is invoked through Windows UI Automation: no keystrokes, no mouse, no other window touched (the owner agreed to
# a visible Excel for this on 2026-11-15). For each style: a copy of the workbook saved and its group's eight colour
# elements kept as Excel wrote them, with the gallery item's name. Excel is quit at the end.
# usage: pwsh tools/excel-compare/excel-sparkstyles.ps1 <out.json> <folder for the saved copies>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.IO.Compression.FileSystem
Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);'
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$AE = [System.Windows.Automation.AutomationElement]
$Scope = [System.Windows.Automation.TreeScope]
$CT = [System.Windows.Automation.ControlType]
$byType = { param($type) New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $type) }
$x = New-Object -ComObject Excel.Application
$mine = [uint32]0
[void][W.U]::GetWindowThreadProcessId([IntPtr]$x.Hwnd, [ref]$mine)
Write-Host "Excel PID $mine"
$x.DisplayAlerts = $false
$styles = @()
try {
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  for ($c = 1; $c -le 6; $c++) { $ws.Cells.Item(1, $c).Value2 = $c * ($c % 2 * 2 - 1) }
  [void]$ws.Range('H1').SparklineGroups.Add(1, 'A1:F1')
  $x.Visible = $true
  $ws.Range('H1').Select()
  Start-Sleep -Seconds 2
  $win = $AE::FromHandle([IntPtr]$x.Hwnd)
  $tab = $win.FindAll($Scope::Descendants, (& $byType $CT::TabItem)) | Where-Object { $_.Current.Name -eq 'Sparkline' } | Select-Object -First 1
  if (-not $tab) { throw 'no Sparkline tab' }
  $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()
  Start-Sleep -Seconds 2
  $items = { $win.FindAll($Scope::Descendants, (& $byType $CT::ListItem)) | Where-Object { $_.Current.ClassName -eq 'NetUIGalleryButton' -and $_.Current.Name -match 'Sparkline Style' } }
  $count = @(& $items).Count
  Write-Host "$count styles"
  for ($k = 0; $k -lt $count; $k++) {
    $item = @(& $items)[$k]
    $name = $item.Current.Name
    $p = $null
    if ($item.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$p)) { $p.Invoke() }
    elseif ($item.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref]$p)) { $p.Select() }
    else { throw "no pattern for $name" }
    Start-Sleep -Milliseconds 700
    $file = Join-Path $Books ("style-{0:d2}.xlsx" -f ($k + 1))
    if (Test-Path $file) { Remove-Item $file }
    $wb.SaveCopyAs($file)
    $zip = [IO.Compression.ZipFile]::OpenRead($file)
    try { $s = New-Object IO.StreamReader($zip.GetEntry('xl/worksheets/sheet1.xml').Open()); $xml = $s.ReadToEnd(); $s.Close() } finally { $zip.Dispose() }
    $colors = [ordered]@{}
    foreach ($m in [regex]::Matches($xml, '<x14:color(\w+) ([^/]*)/>')) { $colors[$m.Groups[1].Value.ToLower()] = $m.Groups[2].Value }
    $styles += [ordered]@{ n = $k + 1; name = $name; colors = $colors }
  }
  $wb.Close($false)
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  Start-Sleep -Milliseconds 800
  $proc = Get-Process -Id $mine -ErrorAction SilentlyContinue
  if ($proc) { Stop-Process -Id $mine -Force }
}
$json = ConvertTo-Json -InputObject $styles -Depth 6
[IO.File]::WriteAllText($Out, $json.Replace("`r`n", "`n") + "`n", (New-Object Text.UTF8Encoding $false))
