# Which name Excel gives a selection, for namebox.test.js. COM can't read the Name Box (in an invisible Excel its Edit
# control stays empty: WM_GETTEXT read "" with a name selected), so this reads what Excel itself reports for the
# selection, Selection.Name (the Name whose range the selection is exactly; an error when there is none). In an
# invisible Excel, for each case a new workbook of two sheets with the case's names (workbook-wide or a sheet's own,
# hidden or not), print area and table; the case's range selected on its sheet (an active cell moved inside it when the
# case says so) and Selection.Name.Name recorded (a sheet's own name comes as "Sheet1!Name"). No keys sent, no setting
# changed, nothing saved.
# usage: pwsh tools/excel-compare/excel-namebox.ps1 <out.json>
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$B = '=Sheet1!$B$2:$C$4'
# names: n name, r refers to, s the sheet it belongs to (0: the workbook), h hidden. sel: on sheet `on` (1 or 2).
$CASES = @(
  @{ id = 'exact'; names = @(@{ n = 'Rate'; r = $B }); sel = 'B2:C4' },
  @{ id = 'exact, the active cell inside'; names = @(@{ n = 'Rate'; r = $B }); sel = 'B2:C4'; active = 'C3' },
  @{ id = 'part of it'; names = @(@{ n = 'Rate'; r = $B }); sel = 'B2:C3' },
  @{ id = 'more than it'; names = @(@{ n = 'Rate'; r = $B }); sel = 'B2:C5' },
  @{ id = 'its first cell'; names = @(@{ n = 'Rate'; r = $B }); sel = 'B2' },
  @{ id = 'one cell'; names = @(@{ n = 'One'; r = '=Sheet1!$E$5' }); sel = 'E5' },
  @{ id = 'two of the workbook, added Zed first'; names = @(@{ n = 'Zed'; r = $B }, @{ n = 'Alpha'; r = $B }); sel = 'B2:C4' },
  @{ id = 'two of the workbook, capitals'; names = @(@{ n = 'beta'; r = $B }, @{ n = 'Alpha2'; r = $B }); sel = 'B2:C4' },
  @{ id = 'the workbook''s and the sheet''s'; names = @(@{ n = 'Wide'; r = $B }, @{ n = 'Own'; r = $B; s = 1 }); sel = 'B2:C4' },
  @{ id = 'the sheet''s and the workbook''s, the sheet''s first by name'; names = @(@{ n = 'Zwide'; r = $B }, @{ n = 'Aown'; r = $B; s = 1 }); sel = 'B2:C4' },
  @{ id = 'the workbook''s first by name, the sheet''s own'; names = @(@{ n = 'Aw'; r = $B }, @{ n = 'Zown'; r = $B; s = 1 }); sel = 'B2:C4' },
  @{ id = 'the workbook''s first by name, another sheet''s own'; names = @(@{ n = 'Aw'; r = $B }, @{ n = 'Zown'; r = $B; s = 2 }); sel = 'B2:C4' },
  @{ id = 'each sheet''s own, the same name'; names = @(@{ n = 'Mid'; r = $B; s = 2 }, @{ n = 'Mid'; r = $B; s = 1 }); sel = 'B2:C4' },
  @{ id = 'each sheet''s own, the other sheet''s first by name'; names = @(@{ n = 'Aa'; r = $B; s = 2 }, @{ n = 'Bb'; r = $B; s = 1 }); sel = 'B2:C4' },
  @{ id = 'hidden first by name, the sheet''s own'; names = @(@{ n = 'Ahid'; r = $B; h = $true }, @{ n = 'Zown'; r = $B; s = 1 }); sel = 'B2:C4' },
  @{ id = 'the same name, the workbook''s and the sheet''s';names = @(@{ n = 'Dup'; r = $B }, @{ n = 'Dup'; r = $B; s = 1 }); sel = 'B2:C4' },
  @{ id = 'another sheet''s own name over this sheet''s cells'; names = @(@{ n = 'Theirs'; r = $B; s = 2 }); sel = 'B2:C4' },
  @{ id = 'another sheet''s own and the workbook''s'; names = @(@{ n = 'Atheirs'; r = $B; s = 2 }, @{ n = 'Wide'; r = $B }); sel = 'B2:C4' },
  @{ id = 'the sheet''s own, from the other sheet'; names = @(@{ n = 'Own'; r = '=Sheet2!$B$2:$C$4'; s = 1 }); sel = 'B2:C4'; on = 2 },
  @{ id = 'a name of the other sheet''s cells'; names = @(@{ n = 'There'; r = '=Sheet2!$B$2:$C$4' }); sel = 'B2:C4' },
  @{ id = 'hidden'; names = @(@{ n = 'Hid'; r = $B; h = $true }); sel = 'B2:C4' },
  @{ id = 'hidden and visible'; names = @(@{ n = 'Ahid'; r = $B; h = $true }, @{ n = 'Seen'; r = $B }); sel = 'B2:C4' },
  @{ id = 'two areas'; names = @(@{ n = 'Multi'; r = '=Sheet1!$A$1,Sheet1!$C$3:$D$4' }); sel = 'A1,C3:D4' },
  @{ id = 'two areas, the other way round'; names = @(@{ n = 'Multi'; r = '=Sheet1!$A$1,Sheet1!$C$3:$D$4' }); sel = 'C3:D4,A1' },
  @{ id = 'two areas, one of them'; names = @(@{ n = 'Multi'; r = '=Sheet1!$A$1,Sheet1!$C$3:$D$4' }); sel = 'A1' },
  @{ id = 'a whole column'; names = @(@{ n = 'Col'; r = '=Sheet1!$B:$B' }); sel = 'B:B' },
  @{ id = 'whole rows'; names = @(@{ n = 'Rows'; r = '=Sheet1!$3:$5' }); sel = '3:5' },
  @{ id = 'relative'; names = @(@{ n = 'Rel'; r = '=Sheet1!B2' }); sel = 'B2' },
  @{ id = 'a formula'; names = @(@{ n = 'Dyn'; r = '=OFFSET(Sheet1!$A$1,0,0,2,2)' }); sel = 'A1:B2' },
  @{ id = 'a print area'; print = '$B$2:$C$4'; sel = 'B2:C4' },
  @{ id = 'a print area and a name'; print = '$B$2:$C$4'; names = @(@{ n = 'Rate'; r = $B }); sel = 'B2:C4' },
  @{ id = 'a table''s data rows'; table = 'A10:C13'; sel = 'A11:C13' },
  @{ id = 'a whole table'; table = 'A10:C13'; sel = 'A10:C13' },
  @{ id = 'a table''s column'; table = 'A10:C13'; sel = 'B11:B13' },
  @{ id = 'a table and a name over its data rows'; table = 'A10:C13'; names = @(@{ n = 'Body'; r = '=Sheet1!$A$11:$C$13' }); sel = 'A11:C13' }
)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$results = @(); $version = $x.Version
try {
  foreach ($c in $CASES) {
    $wb = $x.Workbooks.Add()
    while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $wb.Worksheets.Item(1).Name = 'Sheet1'; $wb.Worksheets.Item(2).Name = 'Sheet2'
    $wb.Worksheets.Item(1).Activate(); $wb.Worksheets.Item(1).Range('A1').Select()
    foreach ($n in @($c.names | Where-Object { $_ })) {
      $owner = if ($n.s) { $wb.Worksheets.Item($n.s) } else { $wb }
      $made = $owner.Names.Add($n.n, $n.r)
      if ($n.h) { $made.Visible = $false }
    }
    if ($c.print) { $wb.Worksheets.Item(1).PageSetup.PrintArea = $c.print }
    if ($c.table) {
      $ws = $wb.Worksheets.Item(1)
      $r = $ws.Range($c.table)
      $ws.Cells.Item($r.Row, $r.Column).Value2 = 'Fruit'; $ws.Cells.Item($r.Row, $r.Column + 1).Value2 = 'Qty'; $ws.Cells.Item($r.Row, $r.Column + 2).Value2 = 'Price'
      [void]$ws.ListObjects.Add(1, $r, [Type]::Missing, 1)
    }
    $on = if ($c.on) { $c.on } else { 1 }
    $ws = $wb.Worksheets.Item($on); $ws.Activate()
    $ws.Range($c.sel).Select()
    if ($c.active) { $ws.Range($c.active).Activate() }
    $got = $null
    try { $got = $x.Selection.Name.Name } catch { $got = $null }
    $results += [pscustomobject]@{ id = $c.id; names = @($c.names | Where-Object { $_ } | ForEach-Object { [pscustomobject]@{ name = $_.n; refersTo = $_.r; sheet = $(if ($_.s) { $_.s - 1 } else { $null }); hidden = [bool]$_.h } });
      print = $c.print; table = $c.table; on = $on - 1; select = $c.sel; active = $c.active; excel = $got }
    $wb.Close($false)
  }
  # Going to a name (Application.Goto, as the Name Box's list goes): what is selected, the active cell, and the name the
  # selection then has; from Sheet2, for a name of Sheet1's cells.
  $gone = @()
  $wb = $x.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $wb.Worksheets.Item(1).Name = 'Sheet1'; $wb.Worksheets.Item(2).Name = 'Sheet2'
  [void]$wb.Names.Add('Rate', $B)
  [void]$wb.Names.Add('Multi', '=Sheet1!$C$3:$D$4,Sheet1!$A$1')
  foreach ($n in @('Rate', 'Multi')) {
    $wb.Worksheets.Item(2).Activate(); $wb.Worksheets.Item(2).Range('A1').Select()
    $x.Goto($n)
    $gone += [pscustomobject]@{ name = $n; sheet = $x.ActiveSheet.Name; selection = $x.Selection.Address($true, $true); active = $x.ActiveCell.Address($true, $true); excel = $x.Selection.Name.Name }
  }
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
[pscustomobject]@{ excel = $version; cases = $results; goto = $gone } | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8NoBOM $Out
'done'
