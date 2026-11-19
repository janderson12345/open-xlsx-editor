# The width in whole pixels Windows (GDI) gives each character of each font Excel's table has (excel-fontmetrics.ps1),
# regular and bold, at each half point from 6 to 36pt on a 96 DPI screen: Excel breaks wrapped lines at the sum of
# these, which the browser's fractional widths miss by a varying few percent (rowheights.js). Each font is made with
# CreateFont at its size in whole pixels (a point is 4/3 px, rounded: 11pt is 15px), characters 32 to 255 read with
# GetCharWidth32. Office's cloud fonts (Aptos, Aptos Narrow, Lato: %LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts)
# are loaded for this process only, nothing installed. A font Windows doesn't have (GetTextFace gives another
# name) is left out.
# Not in use: wrapping at these widths was tried and set aside (docs/LAYOUT.md, Tried and dropped).
# usage: pwsh tools/excel-compare/gdi-charwidths.ps1 <out.json>
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class Gdi {
  [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr hwnd, IntPtr dc);
  [DllImport("gdi32.dll")] public static extern IntPtr SelectObject(IntPtr dc, IntPtr obj);
  [DllImport("gdi32.dll")] public static extern bool DeleteObject(IntPtr obj);
  [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern bool GetCharWidth32W(IntPtr dc, uint first, uint last, int[] widths);
  [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern int GetTextFaceW(IntPtr dc, int n, StringBuilder face);
  [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr CreateFontW(int h, int w, int esc, int orient, int weight, uint italic, uint underline, uint strike, uint charset, uint outPrec, uint clipPrec, uint quality, uint pitch, string face);
  [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern int AddFontResourceExW(string path, uint fl, IntPtr res);
}
'@
$fonts = 'Calibri', 'Calibri Light', 'Arial', 'Arial Narrow', 'Arial Black', 'Times New Roman', 'Aptos Narrow', 'Aptos',
  'Tahoma', 'Book Antiqua', 'Cambria', 'Verdana', 'Lato', 'Trebuchet MS', 'Courier New', 'Segoe UI', 'Georgia',
  'Century Gothic', 'Garamond', 'Consolas', 'Helv', 'Geneva', 'MS Sans Serif', 'Tms Rmn', 'Courier'
$cloud = Join-Path $env:LOCALAPPDATA 'Microsoft\FontCache\4\CloudFonts'
foreach ($n in 'Aptos', 'Aptos Narrow', 'Lato') {
  # (FR_PRIVATE, 0x10: for this process only.)
  Get-ChildItem (Join-Path $cloud $n) -File -ErrorAction SilentlyContinue | ForEach-Object { [void][Gdi]::AddFontResourceExW($_.FullName, 0x10, [IntPtr]::Zero) }
}
$dc = [Gdi]::GetDC([IntPtr]::Zero)
$res = [ordered]@{}
try {
  foreach ($f in $fonts) {
    $entry = [ordered]@{}
    foreach ($bold in $false, $true) {
      $sizes = [ordered]@{}
      for ($s = 6; $s -le 36; $s += 0.5) {
        # (DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH.)
        $h = [Gdi]::CreateFontW(-[int][Math]::Round($s * 4 / 3), 0, 0, 0, $(if ($bold) { 700 } else { 400 }), 0, 0, 0, 1, 0, 0, 0, 0, $f)
        $old = [Gdi]::SelectObject($dc, $h)
        $face = New-Object System.Text.StringBuilder 64
        [void][Gdi]::GetTextFaceW($dc, 64, $face)
        $w = New-Object int[] 224
        [void][Gdi]::GetCharWidth32W($dc, 32, 255, $w)
        [void][Gdi]::SelectObject($dc, $old); [void][Gdi]::DeleteObject($h)
        if ($face.ToString() -ne $f) { $sizes = $null; Write-Output "$f : Windows gives $($face.ToString())"; break }
        $sizes["$s"] = $w
      }
      if (-not $sizes) { $entry = $null; break }
      $entry[$(if ($bold) { 'bold' } else { 'regular' })] = $sizes
    }
    if ($entry) { $res[$f.ToLower()] = $entry }
  }
} finally { [void][Gdi]::ReleaseDC([IntPtr]::Zero, $dc) }
[IO.File]::WriteAllText($Out, ($res | ConvertTo-Json -Depth 5 -Compress), (New-Object System.Text.UTF8Encoding $false))
"$($res.Count) fonts"
