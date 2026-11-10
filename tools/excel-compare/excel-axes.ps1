# Has Excel scale a chart's value axis by itself for sets of numbers (cases: [{ name, values: [numbers], type: "column" |
# "line" | "bar" }]): the axis minimum, maximum and major unit Excel chose (Axes(xlValue).MinimumScale, MaximumScale,
# MajorUnit), for charts drawn from the numbers in a new workbook. An invisible instance, macros disabled for it only,
# closed without saving.
# usage: pwsh tools/excel-compare/excel-axes.ps1 tools/excel-compare/axes-cases.json test/axes-excel.json
param([Parameter(Mandatory)][string]$Cases, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$list = Get-Content $Cases -Raw | ConvertFrom-Json
$types = @{ column = 51; line = 4; bar = 57 }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$x.AutomationSecurity = 3 # msoAutomationSecurityForceDisable: this instance runs no macros
$answers = New-Object System.Collections.Generic.List[object]
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  foreach ($case in $list) {
    $ws.Cells.Clear()
    $vals = @($case.values)
    for ($i = 0; $i -lt $vals.Count; $i++) { $ws.Cells.Item($i + 1, 1).Value2 = [double]$vals[$i] }
    $co = $ws.ChartObjects().Add(100, 20, 360, 220)
    $co.Chart.ChartType = $types[$(if ($case.type) { [string]$case.type } else { 'column' })]
    $co.Chart.SetSourceData($ws.Range("A1:A$($vals.Count)"))
    $ax = $co.Chart.Axes(2)
    $answers.Add([ordered]@{ name = [string]$case.name; min = [double]$ax.MinimumScale; max = [double]$ax.MaximumScale; major = [double]$ax.MajorUnit })
    $co.Delete()
  }
  $wb.Close($false)
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
[ordered]@{ excel = 'Excel.Application: a chart over the numbers, Axes(xlValue) MinimumScale, MaximumScale, MajorUnit'; cases = $answers } | ConvertTo-Json -Depth 4 | Set-Content -Encoding utf8 $Out
"$($answers.Count) cases"
