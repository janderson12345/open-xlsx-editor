# Excel's reading of @mentions written here (mentions.test.js): the workbook opened in an invisible Excel, each
# thread's cell, author, text and replies read, then saved as a copy to a folder outside the repository (no change made)
# so the mentions Excel kept can be read from its own threadedComments part. No keys sent, the original not touched.
# usage: pwsh tools/excel-compare/excel-mentions-check.ps1 <workbook> <out.json> <folder for Excel's copy>
param([Parameter(Mandatory)][string]$Book, [Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Copy)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$found = @()
$saved = Join-Path $Copy 'mentions-excel.xlsx'
try {
  # (CorruptLoad 0, normal: a file Excel had to repair would open with [Repaired] in its name or not at all.)
  $wb = $x.Workbooks.Open((Resolve-Path $Book).Path, 0, $false)
  $name = $wb.Name
  $ws = $wb.Worksheets.Item(1)
  foreach ($t in $ws.CommentsThreaded) {
    $replies = @(); foreach ($r in $t.Replies) { $replies += [pscustomobject]@{ author = $r.Author.Name; text = $r.Text() } }
    $found += [pscustomobject]@{ cell = $t.Parent.Address($false, $false); author = $t.Author.Name; text = $t.Text(); replies = $replies }
  }
  if (Test-Path $saved) { Remove-Item $saved }
  $wb.SaveCopyAs($saved)
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$zip = [IO.Compression.ZipFile]::OpenRead($saved)
$mentions = @()
$persons = @()
try {
  foreach ($e in $zip.Entries) {
    $text = (New-Object IO.StreamReader($e.Open())).ReadToEnd()
    if ($e.FullName -like 'xl/threadedComments/*.xml') { foreach ($m in [regex]::Matches($text, '<mention [^>]*/>')) { $mentions += $m.Value } }
    if ($e.FullName -like 'xl/persons/*.xml') { foreach ($m in [regex]::Matches($text, '<person [^>]*/>')) { $persons += $m.Value } }
  }
} finally { $zip.Dispose() }
[pscustomobject]@{ source = 'tools/excel-compare/excel-mentions-check.ps1 on test/out/mentions/mentions.xlsx (mentions.test.js)'; opened = $name; threads = $found; mentionsKept = $mentions; persons = $persons } | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 $Out
'done'
