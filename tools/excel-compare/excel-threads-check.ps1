# Excel's reading of threaded comments written here (threads.test.js): test/out/threads/threads.xlsx opened read
# only in an invisible Excel; each thread's cell, author, text and replies. Nothing saved, no keys sent.
# usage: pwsh tools/excel-compare/excel-threads-check.ps1 <workbook> <out.json>
param([Parameter(Mandatory)][string]$Book, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$found = @()
try {
  $wb = $x.Workbooks.Open((Resolve-Path $Book).Path, 0, $true)
  $ws = $wb.Worksheets.Item(1)
  foreach ($t in $ws.CommentsThreaded) {
    $replies = @(); foreach ($r in $t.Replies) { $replies += [pscustomobject]@{ author = $r.Author.Name; text = $r.Text() } }
    $found += [pscustomobject]@{ cell = $t.Parent.Address($false, $false); author = $t.Author.Name; text = $t.Text(); replies = $replies }
  }
  $notes = $ws.Comments.Count
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
[pscustomobject]@{ source = 'tools/excel-compare/excel-threads-check.ps1 on test/out/threads/threads.xlsx (threads.test.js)'; threads = $found; legacyComments = $notes } | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 $Out
'done'
