# Has Excel paste HTML from other programs (cases from clip-html-cases.mjs: [{ name, html, text, read, values }]) and
# says what came of it. For each case the clipboard gets the HTML (as CF_HTML, "HTML Format") and its plain text, as
# a program copying a table puts them there; a new workbook in an invisible instance (macros disabled for it only)
# pastes it at B2 (Worksheet.Paste, as Ctrl+V; with `values`, PasteSpecial as Unicode Text, as Paste Values does) and
# is closed without saving. Then for each cell of `read`: formula, value, text, number format, font, fill, borders,
# alignment, wrap, indent and merged area; with the column widths and row heights. The clipboard's text is put back
# afterwards.
# usage: pwsh -STA tools/excel-compare/excel-clip-paste.ps1 tools/excel-compare/clip-html-cases.json test/clip-html-excel.json
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
$keep = [System.Windows.Forms.Clipboard]::GetText()
$list = Get-Content $Cases -Raw | ConvertFrom-Json
$utf8 = New-Object System.Text.UTF8Encoding($false)
# A colour as Excel's number (BGR) to RRGGBB.
# A font property, or 'mixed' when the cell's text has more than one (rich text: COM gives DBNull).
function Val($v, $as) { if ($v -is [System.DBNull]) { 'mixed' } else { $v -as $as } }
function Hex($n) { $n = [int64]$n; '{0:X2}{1:X2}{2:X2}' -f ($n -band 255), (($n -shr 8) -band 255), (($n -shr 16) -band 255) }
# The CF_HTML form: a header of byte offsets, then the HTML, the fragment between its StartFragment and EndFragment.
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
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$answers = New-Object System.Collections.Generic.List[object]
try {
  foreach ($case in $list) {
    $do = New-Object System.Windows.Forms.DataObject
    $do.SetData('HTML Format', (New-Object System.IO.MemoryStream(, (CfHtml $case.html))))
    $do.SetData('UnicodeText', [string]$case.text)
    [System.Windows.Forms.Clipboard]::SetDataObject($do, $true)
    $b = $x.Workbooks.Add()
    $ws = $b.Worksheets.Item(1)
    [void]$ws.Activate(); [void]$ws.Range('B2').Select()
    $a = [ordered]@{ name = $case.name }
    try {
      if ($case.values) { [void]$ws.PasteSpecial('Unicode Text') } else { [void]$ws.Paste() }
    } catch { $a.error = [string]$_.Exception.Message }
    $cells = [ordered]@{}
    $rng = $ws.Range($case.read)
    foreach ($c in $rng.Cells) {
      $f = $c.Font
      $one = [ordered]@{
        formula = [string]$c.Formula; value = $c.Value2; text = [string]$c.Text; numFmt = [string]$c.NumberFormat
        font = [ordered]@{ name = (Val $f.Name ([string])); size = (Val $f.Size ([double])); bold = (Val $f.Bold ([bool])); italic = (Val $f.Italic ([bool])); underline = (Val $f.Underline ([int])); strike = (Val $f.Strikethrough ([bool])); color = $(if ($f.Color -is [System.DBNull]) { 'mixed' } else { Hex $f.Color }) }
        fill = $(if ([int]$c.Interior.ColorIndex -eq -4142) { $null } else { Hex $c.Interior.Color })
        h = [int]$c.HorizontalAlignment; v = [int]$c.VerticalAlignment; wrap = [bool]$c.WrapText; indent = [int]$c.IndentLevel
        merge = $(if ($c.MergeCells) { [string]$c.MergeArea.Address($false, $false) } else { $null })
      }
      # (Rich text: each character's font, run together where they agree.)
      if (@($one.font.Values | Where-Object { $_ -is [string] -and $_ -eq 'mixed' }).Count) {
        $runs = New-Object System.Collections.Generic.List[object]
        $len = ([string]$c.Value2).Length
        for ($k = 1; $k -le $len; $k++) {
          $cf = $c.Characters($k, 1).Font
          $key = [ordered]@{ name = [string]$cf.Name; size = [double]$cf.Size; bold = [bool]$cf.Bold; italic = [bool]$cf.Italic; underline = [int]$cf.Underline; strike = [bool]$cf.Strikethrough; color = (Hex $cf.Color) }
          $ch = ([string]$c.Value2).Substring($k - 1, 1)
          if ($runs.Count -and (($runs[$runs.Count - 1].font | ConvertTo-Json -Compress) -eq ($key | ConvertTo-Json -Compress))) { $runs[$runs.Count - 1].text += $ch }
          else { $runs.Add([ordered]@{ text = $ch; font = $key }) }
        }
        $one.runs = $runs
      }
      $sides = [ordered]@{}
      foreach ($p in @(@('left', 7), @('top', 8), @('bottom', 9), @('right', 10))) {
        $bd = $c.Borders.Item($p[1])
        $sides[$p[0]] = $(if ([int]$bd.LineStyle -eq -4142) { $null } else { [ordered]@{ style = [int]$bd.LineStyle; weight = [int]$bd.Weight; color = (Hex $bd.Color) } })
      }
      $one.border = $sides
      $cells[[string]$c.Address($false, $false)] = $one
    }
    $a.cells = $cells
    $a.widths = @(foreach ($col in $rng.Columns) { [double]$col.ColumnWidth })
    $a.heights = @(foreach ($row in $rng.Rows) { [double]$row.RowHeight })
    $answers.Add($a)
    $x.CutCopyMode = 0
    $b.Close($false)
  }
} finally {
  $x.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  if ($keep) { [System.Windows.Forms.Clipboard]::SetText($keep) }
}
[System.IO.File]::WriteAllText((Join-Path (Get-Location) $Out), ($answers | ConvertTo-Json -Depth 8) + "`n", $utf8)
