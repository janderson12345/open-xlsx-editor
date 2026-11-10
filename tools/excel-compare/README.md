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

## Formula results

```sh
pwsh tools/excel-compare/excel-values.ps1 <book.xlsx> <out.json>
```

Excel opens the workbook read-only and invisibly (links not updated, macros off for this instance), works out every
formula (`CalculateFull`) and the script writes each formula cell's sheet, address, formula and value to JSON; closed
without saving. `test/formulas-excel.json` came from `test/fixtures/formulas.xlsx` this way (docs/CALC.md).

```sh
node tools/excel-compare/lookup-cases.mjs <cases.json>
pwsh tools/excel-compare/excel-lookups.ps1 <cases.json> test/lookup-excel.json
```

Approximate lookups over 400 unsorted lists (the same lists each time, from a fixed seed): Excel answers MATCH 1,
MATCH -1 and VLOOKUP TRUE for each in a new workbook in an invisible instance, closed without saving. On a list that
isn't sorted, Excel's binary search gives its own answer, which `test/lookups.test.js` expects of the engine.

## Editing steps

```sh
node tools/excel-compare/shift-cases.mjs > tools/excel-compare/shift-cases.json
pwsh tools/excel-compare/excel-steps.ps1 tools/excel-compare/shift-cases.json test/shift-excel.json
```

The same for `paste-cases.mjs` (Paste Special, Paste Link) and `fill-cases.mjs` (Fill, AutoFill). Each case is a small sheet built in a new workbook (cells, array formulas, formats, merged ranges, widths, names), one
or more steps done there by Excel (inserting and deleting cells, Paste Special, Paste Link, Fill, the fill handle), and what came of it
written out: each cell's formula, value, number format, bold, fill and bottom border, the merged ranges, the column
widths and the names, or the refusal. An invisible instance, closed without saving.

```sh
STEPS_SAVE=<folder> node --test test/shift.test.js   # or test/paste.test.js
pwsh tools/excel-compare/excel-readback.ps1 <folder>/saved.json <out.json>
```

The same cases done by the editor and saved, then opened in Excel and read the same way, to compare with what Excel
made itself.

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
