# Excel opens the workbooks notes-ours.mjs saved and reads each note: its cell, author and text (Comment.Text()),
# and whether any threaded comment is left. Invisible, read-only, closed without saving.
# usage: pwsh tools/excel-compare/excel-notes-check.ps1 <folder> <out.json>
param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$results = @{}
try {
  foreach ($name in 'new.xlsx', 'edit.xlsx') {
    $wb = $x.Workbooks.Open((Join-Path (Resolve-Path $Folder).Path $name), 0, $true)
    $ws = $wb.Worksheets.Item(1)
    $notes = @(); foreach ($c in $ws.Comments) { $notes += , @($c.Parent.Address($false, $false), $c.Author, ($c.Text() -replace "`r", '')) }
    $results[$name] = @{ notes = $notes; threads = $ws.CommentsThreaded.Count }
    $wb.Close($false)
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$results | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 $Out
'read'
