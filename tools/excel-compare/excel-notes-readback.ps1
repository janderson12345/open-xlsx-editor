# Excel reading back what notes-visible-ours.mjs wrote: in an invisible Excel, each workbook opened and every sheet's
# notes (cell, author, text, Comment.Visible) and threads (cell, each comment's author, text and date) read over COM.
# Nothing saved, no keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-notes-readback.ps1 <out.json> <folder notes-visible-ours.mjs saved into>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
$Books = (Resolve-Path $Books).Path
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$result = [ordered]@{}
try {
  foreach ($name in 'shown.xlsx', 'showall.xlsx', 'converted.xlsx') {
    $wb = $x.Workbooks.Open((Join-Path $Books $name), 0, $true)
    $sheets = @()
    foreach ($ws in $wb.Worksheets) {
      $notes = @(); $threads = @()
      foreach ($c in $ws.Comments) { $notes += [ordered]@{ cell = $c.Parent.Address($false, $false); author = $c.Author; text = $c.Text(); visible = $c.Visible } }
      foreach ($t in $ws.CommentsThreaded) {
        $items = @([ordered]@{ author = $t.Author.Name; text = $t.Text(); date = "$($t.Date)" })
        foreach ($r in $t.Replies) { $items += [ordered]@{ author = $r.Author.Name; text = $r.Text(); date = "$($r.Date)" } }
        $threads += [ordered]@{ cell = $t.Parent.Address($false, $false); items = $items }
      }
      $sheets += [ordered]@{ name = $ws.Name; notes = $notes; threads = $threads }
    }
    $result[$name] = $sheets
    $wb.Close($false)
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$result | ConvertTo-Json -Depth 8 | Set-Content -Encoding utf8 $Out
'done'
