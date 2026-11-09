# Comparing the layout with Excel

Measures how closely the editor places cell text compared with Excel: both print the same workbooks to PDF, the text
of the two PDFs is matched in order, and a cell counts when it's on the same page and its text is within 1px of
Excel's across (x) and down (baseline y). See `docs/LAYOUT.md` for what the numbers mean and where they stand.

Windows, with Excel and Microsoft Edge installed (Edge's fonts and font rendering are part of what's measured).
Node 20 for the repository's Playwright, Python 3, PowerShell 7.

## Once: Excel's PDFs

```powershell
pwsh tools/excel-compare/excel-print.ps1 <folder of .xlsx> <folder for Excel's PDFs>
```

Excel opens each workbook read-only and invisibly (links not updated, macros disabled for that instance), exports the
sheet it opens on as PDF and closes it without saving. No Excel setting is changed and no keys are sent, so it can run
while you work. Workbooks already printed are skipped. Keep the PDFs: Excel's output only changes when Excel does.

## Each change: a run

```powershell
pwsh tools/excel-compare/run.ps1 -Books <folder of .xlsx> -Excel <folder of Excel's PDFs> -Out <work folder> -Run 2 -Node <node 20>
```

It builds the demo, prints every workbook through the editor's print layout (`print-ours.mjs`, four at a time, Edge,
standards mode asserted), extracts the cell text of both sets of PDFs (`cells.cjs`, Excel's once) and writes
`<work folder>/report-2.txt`: one line per workbook, then the totals.

```
basic        pages   2/1   samePage 47/47/49 x1px 2 y1px 5 both 1 meanAbs x 14.0 y 3.1
```

`pages` is Excel's count / ours; `samePage` is cells on the same page / cells matched / Excel's cells; `x1px`, `y1px`
and `both` count cells within 1px. Right-aligned text is compared at its right end, left-aligned at its start.

It also extracts the filled shapes of both (`shapes.cjs`: cell fills, and borders, which Excel draws as filled
rectangles) and writes `report-shapes-2.txt`: per workbook, Excel's shapes matched by one of ours of the same colour
with every edge within 1px, those whose colour is there but not the place, and those missing.

Compare two runs, workbook by workbook:

```powershell
python tools/excel-compare/diff.py <work folder>/report-1.txt <work folder>/report-2.txt
```

A layout change is kept when the total doesn't get worse and nothing that got worse is left unexplained. Run twice
with no change first to learn the noise floor.

## The workbooks

Real workbooks are kept out of the repository (`..\samples`), and referred to by label (`sample X014`), never by
name. The generated fixtures in `test/fixtures` can be used freely.
