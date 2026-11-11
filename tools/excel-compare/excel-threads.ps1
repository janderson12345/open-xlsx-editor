# What Excel writes for threaded comments, for threads.test.js: in an invisible Excel, a new workbook; B2 given a
# thread (Range.AddCommentThreaded) and two replies (CommentThreaded.AddReply), D4 a thread with none, then the D4 thread
# deleted and F6 given one after; saved outside the repository and the parts read back: the threadedComments part, the
# persons part, the legacy comments part and VML drawing Excel keeps beside them, the relationships and content types.
# Excel signs each comment with the signed-in person: their name, id and provider are replaced here by a made-up person
# (Avery Chen) before anything is written out. No keys sent, no setting changed.
# usage: pwsh tools/excel-compare/excel-threads.ps1 <out.json> <folder for the saved workbook>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$read = [ordered]@{}
try {
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $t = $ws.Range('B2').AddCommentThreaded('Check these totals')
  [void]$t.AddReply('Done & checked <twice>')
  [void]$t.AddReply("Thanks`nlooks right")
  [void]$ws.Range('D4').AddCommentThreaded('Delete me')
  $ws.Range('D4').CommentThreaded.Delete()
  [void]$ws.Range('F6').AddCommentThreaded('A second thread')
  $author = $t.Author.Name
  $file = Join-Path $Books 'threads.xlsx'
  if (Test-Path $file) { Remove-Item $file }
  $wb.SaveAs($file, 51)
  $wb.Close($false)
  $zip = [IO.Compression.ZipFile]::OpenRead($file)
  foreach ($e in $zip.Entries) {
    if ($e.FullName -match '^(xl/threadedComments/|xl/persons/|xl/comments|xl/drawings/vmlDrawing|xl/worksheets/_rels/|xl/_rels/workbook|\[Content_Types\]|xl/worksheets/sheet1\.xml)') {
      $s = New-Object IO.StreamReader($e.Open()); $read[$e.FullName] = $s.ReadToEnd(); $s.Close()
    }
  }
  $zip.Dispose()
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
# The signed-in person out: their display name, user id and provider id replaced.
$person = $read['xl/persons/person.xml']
$names = @($author)
foreach ($m in [regex]::Matches($person, '(displayName|userId)="([^"]+)"')) { $names += $m.Groups[2].Value }
foreach ($k in @($read.Keys)) {
  $v = $read[$k]
  foreach ($n in ($names | Where-Object { $_ } | Sort-Object Length -Descending)) {
    $v = $v.Replace($n, $(if ($n -match '@') { 'avery.chen@example.com' } elseif ($n -match '^[0-9a-f]{16}$') { '0123456789abcdef' } else { 'Avery Chen' }))
  }
  $v = [regex]::Replace($v, 'providerId="[^"]*"', 'providerId="AD"')
  $read[$k] = $v
}
[pscustomobject]@{ parts = $read } | ConvertTo-Json -Depth 3 | Set-Content -Encoding utf8 $Out
'done'
