# What Excel writes for a shown note, Show All Notes and Convert to Comments, for notes.test.js: in an invisible Excel,
# new workbooks saved outside the repository and their parts read back (the comments part, the VML drawing, the
# threadedComments and persons parts, the relationships and content types), and what Excel reads in them over COM.
#   visible.xlsx  B2's note shown (Comment.Visible = True), D4's left hidden, F6's shown and hidden again.
#   showall.xlsx  B2's and D4's notes (and F6's, shown first) with Application.DisplayCommentIndicator = xlCommentAndIndicator (1, Show All
#                 Notes), each Comment.Visible read after; the setting put back as it was before saving.
#   convert.xlsx  notes on two sheets (B2 and D4 on Sheet1, D4's shown; C3 on Sheet2), a thread on F6 of Sheet1, then
#                 Workbook.ConvertComments() (Review > Notes > Convert to Comments).
#   ours-converted.xlsx  <ours> (a workbook written here, notes by made-up people) opened, ConvertComments(), saved.
# Excel signs notes with Application.UserName and threads with the signed-in person: both are replaced by a made-up
# "Signed In" person before anything is written out. No keys sent; the one setting changed is put back.
# usage: pwsh tools/excel-compare/excel-notes-visible.ps1 <out.json> <folder for the saved workbooks> <ours.xlsx>
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Books, [Parameter(Mandatory)][string]$Ours)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force $Books | Out-Null
$Books = (Resolve-Path $Books).Path
$Ours = (Resolve-Path $Ours).Path
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$was = $x.DisplayCommentIndicator
$result = [ordered]@{}
$names = @($x.UserName)

function Parts($file) {
  $read = [ordered]@{}
  $zip = [IO.Compression.ZipFile]::OpenRead($file)
  foreach ($e in $zip.Entries) {
    if ($e.FullName -match '^(xl/threadedComments/|xl/persons/|xl/comments|xl/drawings/vmlDrawing|xl/worksheets/_rels/|xl/_rels/workbook|\[Content_Types\]|xl/worksheets/sheet\d+\.xml)') {
      $s = New-Object IO.StreamReader($e.Open()); $read[$e.FullName] = $s.ReadToEnd(); $s.Close()
    }
  }
  $zip.Dispose()
  $read
}
# What Excel reads: every sheet's notes (cell, author, text, visible) and threads (cell, each comment's author, text, date).
function Seen($wb) {
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
  $sheets
}
function Save($wb, $name) {
  $file = Join-Path $Books $name
  if (Test-Path $file) { Remove-Item $file }
  $wb.SaveAs($file, 51)
  $file
}

try {
  # Shown and hidden notes.
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  [void]$ws.Range('B2').AddComment('Shown note')
  [void]$ws.Range('D4').AddComment('Hidden note')
  [void]$ws.Range('F6').AddComment('Shown then hidden')
  $ws.Range('B2').Comment.Visible = $true
  $ws.Range('F6').Comment.Visible = $true
  $ws.Range('F6').Comment.Visible = $false
  $f = Save $wb 'visible.xlsx'; $wb.Close($false)
  $parts = Parts $f
  $wb = $x.Workbooks.Open($f)
  $result['visible'] = [ordered]@{ seen = Seen $wb; parts = $parts }
  $wb.Close($false)

  # Show All Notes.
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  [void]$ws.Range('B2').AddComment('One')
  [void]$ws.Range('D4').AddComment('Two')
  [void]$ws.Range('F6').AddComment('Shown before')
  $ws.Range('F6').Comment.Visible = $true
  $x.DisplayCommentIndicator = 1
  $during = Seen $wb
  $f = Save $wb 'showall.xlsx'
  $x.DisplayCommentIndicator = $was
  $after = Seen $wb
  $wb.Close($false)
  $result['showall'] = [ordered]@{ during = $during; afterPutBack = $after; parts = Parts $f }

  # Convert to Comments, on a workbook Excel made.
  $wb = $x.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $s1 = $wb.Worksheets.Item(1); $s2 = $wb.Worksheets.Item(2)
  [void]$s1.Range('B2').AddComment("First line`nSecond line")
  [void]$s1.Range('D4').AddComment('Shown one')
  $s1.Range('D4').Comment.Visible = $true
  [void]$s2.Range('C3').AddComment('On the second sheet')
  [void]$s1.Range('F6').AddCommentThreaded('A thread already')
  $before = Seen $wb
  $s2.Activate()
  $t0 = Get-Date
  $wb.ConvertComments()
  $t1 = Get-Date
  $seen = Seen $wb
  $f = Save $wb 'convert.xlsx'; $wb.Close($false)
  $result['convert'] = [ordered]@{ before = $before; after = $seen; ranFrom = $t0.ToString('yyyy-MM-ddTHH:mm:ss'); ranTo = $t1.ToString('yyyy-MM-ddTHH:mm:ss'); parts = Parts $f }

  # Convert to Comments, on a workbook written here (its notes' authors made-up people, not the one signed in).
  $wb = $x.Workbooks.Open($Ours)
  $before = Seen $wb
  $wb.ConvertComments()
  $seen = Seen $wb
  $f = Save $wb 'ours-converted.xlsx'; $wb.Close($false)
  $result['ours'] = [ordered]@{ before = $before; after = $seen; parts = Parts $f }
} finally {
  try { $x.DisplayCommentIndicator = $was } catch {}
  $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
# The signed-in person and Excel's user name out.
$json = $result | ConvertTo-Json -Depth 12
foreach ($m in [regex]::Matches($json, '<person displayName=\\"([^"\\]+)\\" id=\\"[^"\\]+\\" userId=\\"([^"\\]+)\\" providerId=\\"(?!None\\")')) { $names += $m.Groups[1].Value; $names += $m.Groups[2].Value }
foreach ($n in ($names | Where-Object { $_ -and $_ -notmatch '^(Avery Chen|Sam Rivera|None)$' } | Sort-Object Length -Descending -Unique)) {
  $json = $json.Replace($n, $(if ($n -match '@') { 'signed.in@example.com' } elseif ($n -match '^[0-9a-f]{16}$') { '0123456789abcdef' } else { 'Signed In' }))
}
$json = [regex]::Replace($json, 'providerId=\\"(?!None\\")[^"\\]*\\"', 'providerId=\"AD\"')
$json | Set-Content -Encoding utf8 $Out
'done'
