# Has Excel open the editor's files with changed M (queries-writeback.mjs) and says what it makes of them: each
# query's Formula as Workbook.Queries reads it, then every connection refreshed (not in the background: a query's
# table with QueryTable.Refresh, the others with WorkbookConnection.Refresh) with any error, the table's range and
# values; then the file saved by Excel, opened again and its Formulas read once more (kept). A file Excel can't open,
# or opens only by repairing it, is an error here.
# An invisible instance, macros disabled for it only; no Excel setting changed, no keys sent.
# usage: pwsh tools/excel-compare/excel-queries-check.ps1 <folder of .xlsx> test/queries-writeback-excel.json
param([Parameter(Mandatory)][string]$Dir, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$Dir = [IO.Path]::GetFullPath($Dir, (Get-Location).Path)
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("oxe-qcheck-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
$formulas = { param($wb) $o = [ordered]@{}; for ($k = 1; $k -le $wb.Queries.Count; $k++) { $q = $wb.Queries.Item($k); $o[$q.Name] = $q.Formula }; $o }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3
$answers = New-Object System.Collections.Generic.List[object]
try {
  foreach ($f in Get-ChildItem $Dir -Filter *.xlsx | Sort-Object Name) {
    $a = [ordered]@{ name = $f.BaseName; errors = @() }
    $wb = $null
    try { $wb = $x.Workbooks.Open($f.FullName, 0, $true) } catch { $a.errors += "open: $($_.Exception.Message)" }
    if ($wb) {
      try {
        $a.formulas = & $formulas $wb
        for ($k = 1; $k -le $wb.Connections.Count; $k++) {
          $c = $wb.Connections.Item($k)
          if ($c.Name -eq 'ThisWorkbookDataModel') { continue }
          try {
            $lo = $null
            foreach ($l in $wb.Worksheets.Item(1).ListObjects) { if ($l.QueryTable -and $l.QueryTable.WorkbookConnection.Name -eq $c.Name) { $lo = $l } }
            if ($lo) { [void]$lo.QueryTable.Refresh($false) }
            else { try { $c.OLEDBConnection.BackgroundQuery = $false } catch {}; $c.Refresh() }
          } catch { $a.errors += "refresh $($c.Name): $($_.Exception.Message)" }
        }
        $lo = $wb.Worksheets.Item(1).ListObjects.Item(1)
        $a.range = $lo.Range.Address($false, $false)
        $a.values = @($lo.Range.Rows | ForEach-Object { , @($_.Cells | ForEach-Object { $_.Value2 }) })
        $saved = Join-Path $tmp "$($f.BaseName).xlsx"
        $wb.SaveAs($saved, 51)
      } finally { $wb.Close($false) }
      try { $wb2 = $x.Workbooks.Open($saved, 0, $true); try { $a.kept = & $formulas $wb2 } finally { $wb2.Close($false) } } catch { $a.errors += "reopen: $($_.Exception.Message)" }
    }
    $answers.Add($a)
  }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
}
$json = ConvertTo-Json -InputObject $answers.ToArray() -Depth 6
[IO.File]::WriteAllText(([IO.Path]::GetFullPath($Out, (Get-Location).Path)), $json + "`n", (New-Object Text.UTF8Encoding($false)))
