# Watches one Excel process (one this repo's scripts started, its window off the screen) for the boxes it puts up: a
# message box or a dialog. Each one's title, its text and its buttons go to a log (a JSON line each), then the box is
# closed by window messages to that process alone (BM_CLICK on a message box's first button, else WM_CLOSE): no key
# or click goes through the desktop. Runs until the stop file appears.
# usage: pwsh -NoProfile -File tools/excel-compare/dialog-watch.ps1 <excel pid> <log> <stop file>
param([Parameter(Mandatory)][int]$ExcelPid, [Parameter(Mandatory)][string]$Log, [Parameter(Mandatory)][string]$Stop)
$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class DW {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc f, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc f, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, StringBuilder l);
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  public static string Cls(IntPtr h) { var s = new StringBuilder(256); GetClassName(h, s, 256); return s.ToString(); }
  public static string Text(IntPtr h) {
    int n = (int)SendMessage(h, 0x000E, IntPtr.Zero, IntPtr.Zero);
    var s = new StringBuilder(Math.Max(n + 1, 2));
    SendMessage(h, 0x000D, (IntPtr)s.Capacity, s);
    return s.ToString();
  }
  public static List<IntPtr> Tops(uint pid) {
    var o = new List<IntPtr>();
    EnumWindows((h, l) => { uint p; GetWindowThreadProcessId(h, out p); if (p == pid && IsWindowVisible(h)) o.Add(h); return true; }, IntPtr.Zero);
    return o;
  }
  public static List<IntPtr> Kids(IntPtr top) {
    var o = new List<IntPtr>();
    EnumChildWindows(top, (h, l) => { o.Add(h); return true; }, IntPtr.Zero);
    return o;
  }
}
'@
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$UA = [Windows.Automation.AutomationElement]
$CT = [Windows.Automation.ControlType]
# (Excel's message boxes are DirectUI, no child windows: their text and buttons are read through UI Automation.)
function Uia($h) {
  $el = $UA::FromHandle($h)
  $all = $el.FindAll([Windows.Automation.TreeScope]::Descendants, [Windows.Automation.Condition]::TrueCondition)
  $texts = @(); $buttons = @(); $inputs = 0
  foreach ($e in $all) {
    $t = $e.Current.ControlType; $n = $e.Current.Name
    if ($t -in @($CT::Edit, $CT::ComboBox, $CT::List, $CT::CheckBox, $CT::RadioButton, $CT::Tree, $CT::Tab)) { $inputs++ }
    if (-not $n) { continue }
    if ($t -eq $CT::Button) { $buttons += $e } elseif ($t -eq $CT::Text) { $texts += $n }
  }
  @{ texts = $texts; buttons = $buttons; inputs = $inputs }
}
$seen = @{}
while (-not (Test-Path $Stop)) {
  foreach ($h in [DW]::Tops([uint32]$ExcelPid)) {
    $cls = [DW]::Cls($h)
    if ($cls -notin @('#32770', 'bosa_sdm_XL9', 'NUIDialog')) { continue }
    $kids = [DW]::Kids($h)
    $texts = @(); $buttons = @()
    foreach ($k in $kids) {
      $kc = [DW]::Cls($k); $kt = [DW]::Text($k)
      if ($kc -eq 'Button') { $buttons += $kt } elseif ($kc -eq 'Static' -and $kt) { $texts += $kt }
    }
    $u = $null
    if ($cls -eq 'NUIDialog') { try { $u = Uia $h; $texts += $u.texts; $buttons += @($u.buttons | ForEach-Object { $_.Current.Name }) } catch { $texts += "(uia: $_)" } }
    $entry = [ordered]@{ at = [DateTime]::UtcNow.ToString('o'); class = $cls; title = [DW]::Text($h); text = ($texts -join ' | '); buttons = $buttons }
    $key = "$h"
    if (-not $seen.ContainsKey($key)) { Add-Content -Path $Log -Value ($entry | ConvertTo-Json -Compress -Depth 3) -Encoding utf8; $seen[$key] = 1 }
    $btn = $kids | Where-Object { [DW]::Cls($_) -eq 'Button' } | Select-Object -First 1
    $ub = if ($u) { @($u.buttons | Where-Object { $_.Current.Name -notmatch '^(Close|Help|Show Help)$' }) } else { @() }
    # (A message box, nothing to fill in: its first button, as Enter would. A dialog: closed, as Esc would.)
    $isBox = $u -and $ub.Count -and -not $u.inputs
    if ($isBox) { try { $ub[0].GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke() } catch { [void][DW]::PostMessage($h, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) } }
    elseif ($cls -eq '#32770' -and $btn) { [void][DW]::SendMessage($btn, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero) } else { [void][DW]::PostMessage($h, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) }
  }
  Start-Sleep -Milliseconds 150
}
