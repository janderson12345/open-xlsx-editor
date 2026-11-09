# Has Excel print each .xlsx in a folder to PDF: opened read-only and invisibly, links not updated, macros disabled for
# this instance only, the sheet it opens on exported, closed without saving. No Excel setting is changed and no keys
# are sent. Workbooks already printed are skipped, so a rerun picks up where it stopped. Each one gets a time limit.
# usage: pwsh tools/excel-compare/excel-print.ps1 <xlsxDir> <pdfDir> [secondsPerWorkbook]
param([Parameter(Mandatory)][string]$In, [Parameter(Mandatory)][string]$Out, [int]$Seconds = 120)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $Out | Out-Null
$In = (Resolve-Path $In).Path; $Out = (Resolve-Path $Out).Path
$files = Get-ChildItem $In -File | Where-Object { $_.Extension -match '^\.xls[xm]$' -and -not (Test-Path (Join-Path $Out ($_.BaseName + '.pdf'))) }
foreach ($f in $files) {
  $pdf = Join-Path $Out ($f.BaseName + '.pdf')
  # One Excel per workbook, in its own job, so a workbook that hangs Excel can be abandoned after the time limit.
  $job = Start-Job -ArgumentList $f.FullName, $pdf -ScriptBlock {
    param($book, $pdf)
    $x = New-Object -ComObject Excel.Application
    $x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
    $x.AutomationSecurity = 3 # msoAutomationSecurityForceDisable: this instance runs no macros
    try {
      # Open(Filename, UpdateLinks: 0 = don't, ReadOnly: true)
      $wb = $x.Workbooks.Open($book, 0, $true)
      # The sheet the workbook opens on, as the editor shows it: xlTypePDF = 0.
      $wb.ActiveSheet.ExportAsFixedFormat(0, $pdf)
      $wb.Close($false)
    } finally {
      $x.Quit()
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
      [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
  }
  if (Wait-Job $job -Timeout $Seconds) { Receive-Job $job -ErrorAction SilentlyContinue | Out-Null; if ($job.State -eq 'Failed') { "FAILED $($f.Name): $($job.ChildJobs[0].JobStateInfo.Reason.Message)" } }
  else { "TIMEOUT $($f.Name)"; Stop-Job $job }
  Remove-Job $job -Force
}
"printed $((Get-ChildItem $Out -Filter *.pdf).Count)"
