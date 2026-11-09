# One full comparison with Excel: build the demo page, print every workbook through it, extract the cell text of both
# sets of PDFs, and write the report. Excel's PDFs are made once (excel-print.ps1) and kept; ours go to a new folder per
# run, so two runs can be compared (diff.py).
# usage: pwsh tools/excel-compare/run.ps1 -Books <folder of .xlsx> -Excel <folder of Excel's PDFs> -Out <work folder> -Run <name> [-Node <node.exe>]
param(
  [Parameter(Mandatory)][string]$Books,
  [Parameter(Mandatory)][string]$Excel,
  [Parameter(Mandatory)][string]$Out,
  [Parameter(Mandatory)][string]$Run,
  # Node 20 or later (Playwright needs it).
  [string]$Node = 'node'
)
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path "$PSScriptRoot\..\..").Path
New-Item -ItemType Directory -Force $Out | Out-Null
$Out = (Resolve-Path $Out).Path
$Books = (Resolve-Path $Books).Path
$Excel = (Resolve-Path $Excel).Path
# Each step has a time limit: a step that hangs stops the run instead of waiting forever.
function Step($exe, $argList, $minutes, $log) {
  $p = Start-Process -FilePath $exe -ArgumentList $argList -NoNewWindow -PassThru -RedirectStandardOutput "$log.out" -RedirectStandardError "$log.err" -WorkingDirectory $repo
  if (-not $p.WaitForExit($minutes * 60000)) { $p.Kill($true); throw "timeout: $exe $argList" }
  if ($p.ExitCode -ne 0) { Get-Content "$log.err" -Tail 20; throw "failed ($($p.ExitCode)): $exe $argList" }
}
Step $Node @('build.mjs') 3 "$Out\build"
if (-not (Test-Path "$Out\cells-excel")) { Step $Node @('tools\excel-compare\cells.cjs', "`"$Excel`"", "`"$Out\cells-excel`"") 20 "$Out\cells-excel" }
Step $Node @('tools\excel-compare\print-ours.mjs', "`"$Books`"", "`"$Out\pdf-$Run`"") 60 "$Out\print-$Run"
Step $Node @('tools\excel-compare\cells.cjs', "`"$Out\pdf-$Run`"", "`"$Out\cells-$Run`"") 20 "$Out\cells-$Run"
Step 'python' @('tools\excel-compare\compare.py', "`"$Out\cells-excel`"", "`"$Out\cells-$Run`"") 5 "$Out\compare-$Run"
Copy-Item "$Out\compare-$Run.out" "$Out\report-$Run.txt"
Get-Content "$Out\report-$Run.txt" -Tail 1
