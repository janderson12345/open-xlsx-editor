# Rich text in Excel (cases from richruns-cases.mjs, written to richruns-cases.json): what Format Cells' Font tab does
# to part of a cell's text (Characters(start, length).Font, as the box sets it for the characters selected in the cell
# editor), what Copy and Paste (and Paste Special's choices, Transpose and a cut) do with a cell's runs, and what
# pasting another program's HTML with mixed formatting in a cell makes. Each cell is read back character by character
# (font, size, bold, italic, underline, strikethrough, colour, super- and subscript, run together where they agree),
# and the workbook saved so its string's XML (the `si` or `is`) is kept too. Done in an invisible instance (macros
# disabled for it only); the clipboard's text is put back afterwards.
# usage: pwsh -STA tools/excel-compare/excel-richruns.ps1 tools/excel-compare/richruns-cases.json test/richruns-excel.json
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem
$keep = [System.Windows.Forms.Clipboard]::GetText()
$spec = Get-Content $Cases -Raw | ConvertFrom-Json
$utf8 = New-Object System.Text.UTF8Encoding($false)
function Hex($n) { $n = [int64]$n; '{0:X2}{1:X2}{2:X2}' -f ($n -band 255), (($n -shr 8) -band 255), (($n -shr 16) -band 255) }
function CfHtml([string]$html) {
  $head = "Version:0.9`r`nStartHTML:{0:D10}`r`nEndHTML:{1:D10}`r`nStartFragment:{2:D10}`r`nEndFragment:{3:D10}`r`n"
  $len = $utf8.GetByteCount(($head -f 0, 0, 0, 0))
  $s = $html.IndexOf('<!--StartFragment-->')
  $e = $html.IndexOf('<!--EndFragment-->')
  if ($s -lt 0) { $s = 0; $e = $html.Length } else { $s += '<!--StartFragment-->'.Length }
  $fs = $len + $utf8.GetByteCount($html.Substring(0, $s))
  $fe = $len + $utf8.GetByteCount($html.Substring(0, $e))
  $all = ($head -f $len, ($len + $utf8.GetByteCount($html)), $fs, $fe) + $html
  $utf8.GetBytes($all)
}
# A cell's text with each character's font, run together where they agree; the cell's own font beside it.
function Runs($c) {
  $text = [string]$c.Value2
  $runs = New-Object System.Collections.Generic.List[object]
  for ($k = 1; $k -le $text.Length; $k++) {
    $cf = $c.Characters($k, 1).Font
    $key = [ordered]@{ name = [string]$cf.Name; size = [double]$cf.Size; bold = [bool]$cf.Bold; italic = [bool]$cf.Italic; underline = [int]$cf.Underline; strike = [bool]$cf.Strikethrough; color = (Hex $cf.Color); sup = [bool]$cf.Superscript; sub = [bool]$cf.Subscript }
    $ch = $text.Substring($k - 1, 1)
    if ($runs.Count -and (($runs[$runs.Count - 1].font | ConvertTo-Json -Compress) -eq ($key | ConvertTo-Json -Compress))) { $runs[$runs.Count - 1].text += $ch }
    else { $runs.Add([ordered]@{ text = $ch; font = $key }) }
  }
  [ordered]@{ value = $c.Value2; formula = [string]$c.Formula; runs = $runs }
}
# A source cell: `text`, then each { start, len, font } of `parts` set on its characters.
function Fill($c, $src) {
  if ($src.number -ne $null) { $c.Value2 = [double]$src.number; return }
  $c.Value2 = [string]$src.text
  foreach ($p in @($src.parts)) { if ($p) { SetFont $c.Characters([int]$p.start, [int]$p.len).Font $p.font } }
}
function SetFont($f, $font) {
  foreach ($prop in $font.PSObject.Properties) {
    switch ($prop.Name) {
      'color' { $h = [string]$prop.Value; $f.Color = [Convert]::ToInt32($h.Substring(4, 2) + $h.Substring(2, 2) + $h.Substring(0, 2), 16) }
      default { $f.($prop.Name) = $prop.Value }
    }
  }
}
$scratch = Join-Path ([System.IO.Path]::GetTempPath()) ('oxe-richruns-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $scratch | Out-Null
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$answers = [ordered]@{ part = @(); paste = @(); html = @() }
try {
  $b = $x.Workbooks.Add()
  $ws = $b.Worksheets.Item(1)
  # Format Cells' Font tab on part of the text: one row each, A.
  $row = 0
  foreach ($case in $spec.part) {
    $row++
    Fill $ws.Cells.Item($row, 1) $case.src
    $answers.part += [ordered]@{ name = $case.name; cell = "A$row" }
  }
  # Copy and paste: the source in A, what was in the destination in C, the paste there.
  $prow = 0
  foreach ($case in $spec.paste) {
    $prow++
    $r = 100 + $prow * 3
    $src = $ws.Cells.Item($r, 1)
    Fill $src $case.src
    $dst = $ws.Cells.Item($r, 3)
    if ($case.dst) { Fill $dst $case.dst }
    $a = [ordered]@{ name = $case.name; cell = [string]$dst.Address($false, $false) }
    try {
      if ($case.how -eq 'cut') { [void]$src.Cut($dst) }
      elseif ($case.how -eq 'paste') { [void]$src.Copy(); [void]$ws.Activate(); [void]$dst.Select(); [void]$ws.Paste() }
      elseif ($case.how -eq 'otherBook') {
        $b2 = $x.Workbooks.Add(); [void]$src.Copy(); [void]$b2.Worksheets.Item(1).Range('C1').PasteSpecial(-4104)
        $a.other = Runs $b2.Worksheets.Item(1).Range('C1'); $x.CutCopyMode = 0; $b2.Close($false)
      }
      else { [void]$src.Copy(); [void]$dst.PasteSpecial([int]$case.paste, -4142, $false, [bool]$case.transpose) }
    } catch { $a.error = [string]$_.Exception.Message }
    $x.CutCopyMode = 0
    $answers.paste += $a
  }
  # Another program's HTML pasted (as Ctrl+V) at E, one row each.
  $hrow = 0
  foreach ($case in $spec.html) {
    $hrow++
    $r = 300 + $hrow * 3
    $do = New-Object System.Windows.Forms.DataObject
    $do.SetData('HTML Format', (New-Object System.IO.MemoryStream(, (CfHtml $case.html))))
    $do.SetData('UnicodeText', [string]$case.text)
    [System.Windows.Forms.Clipboard]::SetDataObject($do, $true)
    [void]$ws.Activate(); [void]$ws.Cells.Item($r, 5).Select()
    $a = [ordered]@{ name = $case.name; cell = [string]$ws.Cells.Item($r, 5).Address($false, $false) }
    # (The clipboard can be busy a moment after it is set: tried again a few times.)
    for ($try = 0; $try -lt 5; $try++) {
      try { [void]$ws.Paste(); $a.Remove('error'); break } catch { $a.error = [string]$_.Exception.Message; Start-Sleep -Milliseconds 300 }
    }
    $answers.html += $a
  }
  # Every cell read back, then the file saved for its strings' XML.
  foreach ($group in @('part', 'paste', 'html')) {
    foreach ($a in $answers[$group]) {
      $c = $ws.Range($a.cell)
      $got = Runs $c
      foreach ($k in $got.Keys) { $a[$k] = $got[$k] }
      $a.cellFont = [ordered]@{ name = [string]$c.Font.Name; size = $(if ($c.Font.Size -is [System.DBNull]) { 'mixed' } else { [double]$c.Font.Size }) }
    }
  }
  $file = Join-Path $scratch 'rich.xlsx'
  $b.SaveAs($file, 51)
  $b.Close($false)
  $zip = [System.IO.Compression.ZipFile]::OpenRead($file)
  try {
    $read = { param($name) $e = $zip.GetEntry($name); if (-not $e) { return $null }; $s = New-Object System.IO.StreamReader($e.Open()); try { $s.ReadToEnd() } finally { $s.Close() } }
    $sst = [xml](& $read 'xl/sharedStrings.xml')
    $sheet = [xml](& $read 'xl/worksheets/sheet1.xml')
    $ns = New-Object System.Xml.XmlNamespaceManager($sheet.NameTable)
    $ns.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $items = @($sst.sst.si)
    foreach ($group in @('part', 'paste', 'html')) {
      foreach ($a in $answers[$group]) {
        $cnode = $sheet.SelectSingleNode("//m:c[@r='$($a.cell)']", $ns)
        if (-not $cnode) { continue }
        $t = $cnode.GetAttribute('t')
        if ($t -eq 's') { $a.xml = $items[[int]$cnode.SelectSingleNode('m:v', $ns).InnerText].InnerXml -replace ' xmlns="[^"]*"', '' }
        elseif ($t -eq 'inlineStr') { $a.xml = $cnode.SelectSingleNode('m:is', $ns).InnerXml -replace ' xmlns="[^"]*"', '' }
      }
    }
  } finally { $zip.Dispose() }
} finally {
  $x.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  if ($keep) { [System.Windows.Forms.Clipboard]::SetText($keep) }
  Remove-Item -Recurse -Force $scratch -ErrorAction SilentlyContinue
}
[System.IO.File]::WriteAllText((Join-Path (Get-Location) $Out), ($answers | ConvertTo-Json -Depth 10) + "`n", $utf8)
