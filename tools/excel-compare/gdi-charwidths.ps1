# The width in whole pixels Windows (GDI) gives each character of each font Excel's table has (excel-fontmetrics.ps1),
# regular, bold, italic and bold italic, at each half point from 6 to 36pt on a 96 DPI screen, and the font's descent
# there (TEXTMETRIC's tmDescent): Excel breaks wrapped lines at the sum of these widths, which the browser's
# fractional widths miss by a varying few percent, and a row's height counts the descents (rowheights.js). Each font
# is made with CreateFont at its size in whole pixels (a point is 4/3 px, rounded: 11pt is 15px), characters 32 to
# 255 read with GetCharWidth32. Office's cloud fonts (Aptos, Aptos Narrow, Lato:
# %LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts) are loaded for this process only, nothing installed. A font
# Windows doesn't have (GetTextFace gives another name) is left out. Under 8pt (10px and less) Excel lays every font
# out in Windows' bitmap font Small Fonts (excel-wrapwidths.ps1 measured it): its widths and descents at 5 to 10px are
# written as "small fonts", by pixel size.
# Written as { font: { regular|bold|italic|boldItalic: { pt: [widths] }, desc: { regular|...: { pt: px } } } }.
# charwidths-table.mjs makes src/ooxml/excelwidths.js of the result.
# usage: pwsh tools/excel-compare/gdi-charwidths.ps1 <out.json>
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct TextMetric {
  public int tmHeight, tmAscent, tmDescent, tmInternalLeading, tmExternalLeading, tmAveCharWidth, tmMaxCharWidth, tmWeight, tmOverhang, tmDigitizedAspectX, tmDigitizedAspectY;
  public char tmFirstChar, tmLastChar, tmDefaultChar, tmBreakChar;
  public byte tmItalic, tmUnderlined, tmStruckOut, tmPitchAndFamily, tmCharSet;
}
public static class Gdi {
  [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr hwnd, IntPtr dc);
  [DllImport("gdi32.dll")] public static extern IntPtr SelectObject(IntPtr dc, IntPtr obj);
  [DllImport("gdi32.dll")] public static extern bool DeleteObject(IntPtr obj);
  [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern bool GetCharWidth32W(IntPtr dc, uint first, uint last, int[] widths);
  [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern bool GetTextMetricsW(IntPtr dc, out TextMetric tm);
  [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern int GetTextFaceW(IntPtr dc, int n, StringBuilder face);
  [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr CreateFontW(int h, int w, int esc, int orient, int weight, uint italic, uint underline, uint strike, uint charset, uint outPrec, uint clipPrec, uint quality, uint pitch, string face);
  [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern int AddFontResourceExW(string path, uint fl, IntPtr res);
}
'@
$fonts = 'Calibri', 'Calibri Light', 'Arial', 'Arial Narrow', 'Arial Black', 'Times New Roman', 'Aptos Narrow', 'Aptos',
  'Tahoma', 'Book Antiqua', 'Cambria', 'Verdana', 'Lato', 'Trebuchet MS', 'Courier New', 'Segoe UI', 'Georgia',
  'Century Gothic', 'Garamond', 'Consolas', 'Helv', 'Geneva', 'MS Sans Serif', 'Tms Rmn', 'Courier'
$styles = [ordered]@{ regular = @(400, 0); bold = @(700, 0); italic = @(400, 1); boldItalic = @(700, 1) }
$cloud = Join-Path $env:LOCALAPPDATA 'Microsoft\FontCache\4\CloudFonts'
foreach ($n in 'Aptos', 'Aptos Narrow', 'Lato') {
  # (FR_PRIVATE, 0x10: for this process only.)
  Get-ChildItem (Join-Path $cloud $n) -File -ErrorAction SilentlyContinue | ForEach-Object { [void][Gdi]::AddFontResourceExW($_.FullName, 0x10, [IntPtr]::Zero) }
}
$dc = [Gdi]::GetDC([IntPtr]::Zero)
# The font's face, its characters' widths, its descent and its overhang (what a made-up bold or italic adds a line).
function GdiFont($face, $px, $weight, $italic) {
  # (DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH.)
  $h = [Gdi]::CreateFontW(-$px, 0, 0, 0, $weight, $italic, 0, 0, 1, 0, 0, 0, 0, $face)
  $old = [Gdi]::SelectObject($dc, $h)
  $got = New-Object System.Text.StringBuilder 64
  [void][Gdi]::GetTextFaceW($dc, 64, $got)
  $w = New-Object int[] 224
  [void][Gdi]::GetCharWidth32W($dc, 32, 255, $w)
  $tm = New-Object TextMetric
  [void][Gdi]::GetTextMetricsW($dc, [ref]$tm)
  [void][Gdi]::SelectObject($dc, $old); [void][Gdi]::DeleteObject($h)
  return @{ face = $got.ToString(); widths = $w; descent = $tm.tmDescent; overhang = $tm.tmOverhang }
}
$res = [ordered]@{}
try {
  foreach ($f in $fonts) {
    $entry = [ordered]@{}
    $desc = [ordered]@{}
    foreach ($style in $styles.Keys) {
      $sizes = [ordered]@{}
      $descs = [ordered]@{}
      for ($s = 6; $s -le 36; $s += 0.5) {
        $m = GdiFont $f ([int][Math]::Round($s * 4 / 3)) $styles[$style][0] $styles[$style][1]
        if ($m.face -ne $f) { $sizes = $null; Write-Output "$f : Windows gives $($m.face)"; break }
        $sizes["$s"] = $m.widths
        $descs["$s"] = $m.descent
      }
      if (-not $sizes) { $entry = $null; break }
      $entry[$style] = $sizes
      $desc[$style] = $descs
    }
    if ($entry) { $entry.desc = $desc; $res[$f.ToLower()] = $entry }
  }
  $small = [ordered]@{}
  foreach ($style in $styles.Keys) {
    $sizes = [ordered]@{}; $descs = [ordered]@{}; $over = [ordered]@{}
    for ($px = 5; $px -le 10; $px++) {
      $m = GdiFont 'Small Fonts' $px $styles[$style][0] $styles[$style][1]
      if ($m.face -ne 'Small Fonts') { throw "Small Fonts: Windows gives $($m.face)" }
      $sizes["$px"] = $m.widths; $descs["$px"] = $m.descent; $over["$px"] = $m.overhang
    }
    $small[$style] = [ordered]@{ px = $sizes; desc = $descs; overhang = $over }
  }
  $res['small fonts'] = $small
} finally { [void][Gdi]::ReleaseDC([IntPtr]::Zero, $dc) }
[IO.File]::WriteAllText($Out, ($res | ConvertTo-Json -Depth 6 -Compress), (New-Object System.Text.UTF8Encoding $false))
"$($res.Count) fonts"
