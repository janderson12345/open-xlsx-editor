# Comparing the layout with Excel

Measures how closely the editor places cell text compared with Excel: both print the same workbooks to PDF, the text
of the two PDFs is matched in order, and a cell counts when it's on the same page and its text is within 1px of
Excel's across (x) and down (baseline y). See `docs/LAYOUT.md` for what the numbers mean and where they stand.

Windows, with Excel and Microsoft Edge installed (Edge's fonts and font rendering are part of what's measured).
Node 20 for the repository's Playwright, Python 3, PowerShell 7. With `OXE_FONTS=1` the demo serves open-metric-fonts'
Morrow Sans from `demo/fonts/omf/morrow-sans` (the README's Fonts section), as a page would for Aptos Narrow.

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

```sh
node tools/excel-compare/emoji-cases.mjs <formulas.json>
pwsh tools/excel-compare/excel-formulas.ps1 <formulas.json> <units.json> -Book test/fixtures/basic.xlsx
pwsh tools/excel-compare/excel-formulas.ps1 <formulas.json> <points.json> -Book <basic.xlsx with setVersion 2>
node tools/excel-compare/formula-results.mjs test/emoji-excel.json units=<units.json> points=<points.json>
```

`excel-formulas.ps1` has Excel work out a list of formulas: each is written into a cell of a sheet added to a new
workbook (or to `-Book`, opened read-only), its Value2 read, and the workbook closed without saving. A text result
keeps its UTF-16 units, so half an emoji survives. `formula-results.mjs` puts runs over the same formulas into one
test file. The emoji cases (docs/CALC.md) are worked out once in basic.xlsx and once in a copy whose workbook.xml asks
for function compatibility version 2 (as `test/emoji.test.js` builds it).

## Editing steps

```sh
node tools/excel-compare/shift-cases.mjs > tools/excel-compare/shift-cases.json
pwsh tools/excel-compare/excel-steps.ps1 tools/excel-compare/shift-cases.json test/shift-excel.json
```

The same for `paste-cases.mjs` (Paste Special, Paste Link), `fill-cases.mjs` (Fill, AutoFill), `clear-cases.mjs` (Clear) and `sort-cases.mjs` (Sort: Worksheet.Sort, and Range.Sort where headers are guessed). Charts' automatic axis has its own, `excel-axes.ps1` on `axes-cases.mjs`: a chart over each set of numbers, its value axis read. Data validation has its own, `excel-dv.ps1` on `dv-cases.mjs` (and `excel-dv-x14.ps1 <out.json> [-Open <folder>]` for the rules Excel writes in its 2010 form, those reading another sheet, and for Excel reading back the editor's): each cell judged with Range.Validation.Value, and the rule's XML kept. Conditional formatting has its own, `excel-cf.ps1` on `cf-cases.mjs`: the rules added, each cell's look read with Range.DisplayFormat, and the rules' XML kept. Filter has its own, `excel-filter.ps1` on `filter-cases.mjs`: the filter applied, then the workbook saved by Excel so its autoFilter XML and hidden rows can be compared. AutoSum has its own script, `excel-autosum.ps1` on `autosum-cases.mjs`: the ribbon's command needs a visible Excel, so it attaches to the one open (Windows PowerShell) and works in a workbook of its own; only selections of several cells, since on one cell Excel waits for Enter and answers nothing over COM. Each case is a small sheet built in a new workbook (cells, array formulas, formats, merged ranges, widths, names), one
or more steps done there by Excel (inserting and deleting cells, Paste Special, Paste Link, Fill, the fill handle), and what came of it
written out: each cell's formula, value, number format, bold, fill and bottom border, the merged ranges, the column
widths and the names, or the refusal. An invisible instance, closed without saving.

```sh
STEPS_SAVE=<folder> node --test test/shift.test.js   # or test/paste.test.js
pwsh tools/excel-compare/excel-readback.ps1 <folder>/saved.json <out.json>
```

The same cases done by the editor and saved, then opened in Excel and read the same way, to compare with what Excel
made itself.

Notes, pictures, charts and shapes when rows and columns move: `anchor-cases.mjs [--no-table]` makes the cases,
`excel-anchors.ps1 <cases> <out.json>` has Excel do each one and read where everything went (nothing saved), and
`anchors-ours.mjs <folder> [--table]` with `excel-anchors-check.ps1 <folder> <out.json>` has Excel read back the
editor's results (docs/ROWCOL.md).

PivotTables when rows and columns move: `excel-pivot-lines.ps1 <pivot-cases.json> <out.json>` (nothing saved), and
`pivot-lines-ours.mjs <folder>` with `excel-pivot-lines-check.ps1 <folder> <out.json>` for Excel reading back the
editor's results.

Sparklines, data validation rules, notes and Circle Invalid Data's ovals when rows, columns and cells move or are
sorted, and slicers and timelines when rows and columns move (and conditional formats of both forms, the same way:
`cfmove-cases.mjs`, `excel-cfmoves.ps1 <cases> <out.json> [-Books test/out/cfmoves]`, `refmove-results.mjs cfmoves`):

```sh
node tools/excel-compare/refmove-cases.mjs > tools/excel-compare/refmove-cases.json
pwsh tools/excel-compare/excel-refmoves.ps1 tools/excel-compare/refmove-cases.json <out.json>
node tools/excel-compare/refmove-results.mjs refmoves tools/excel-compare/refmove-cases.json <out.json> > test/refmoves-excel.json
node tools/excel-compare/slicer-cases.mjs > tools/excel-compare/slicer-cases.json
pwsh tools/excel-compare/excel-slicers.ps1 tools/excel-compare/slicer-cases.json <out.json> [<folder for Excel's books>]
node tools/excel-compare/refmove-results.mjs slicers tools/excel-compare/slicer-cases.json <out.json> > test/slicers-excel.json
```

Each case's workbook is saved to %TEMP%, its parts read and the file deleted (Excel's books are kept only when a folder
is given: they carry Excel's user name, so keep them outside the repository). After `npm test` has saved the editor's
results in `test/out/refmoves` and `test/out/slicers`, `excel-refmoves.ps1 ... -Books test/out/refmoves` and
`excel-slicers.ps1 ... -Open test/out/slicers` read those back the same way, `repaired` saying whether Excel had to
repair one (a repair log appeared in %TEMP%).

PivotTables have their own: `pivot-cases.mjs` makes the cases, `excel-pivot.ps1 <cases> <out.json> <books folder>` has Excel build,
change and refresh each, `pivot-parts.mjs` takes the three pivot parts into `test/pivot-excel.json` (without the user's name), and
`pivot-ours.mjs <folder>` with `excel-pivot-check.ps1 <folder>` has Excel open and refresh the editor's refreshes (docs/PIVOT.md);
`pivotnew-ours.mjs <folder>` with the same check, the editor's new PivotTables; `excel-pivot-move.ps1 <cases> <out.json>
<books folder>` has Excel move a pivot's fields, and `pivotmove-ours.mjs <folder>` the editor's moves for the check.

Defined names: `excel-names.ps1 <out.json> <books folder>` adds candidate names with Names.Add and saves names added out
of order, for `test/names-excel.json` (docs/NAMES.md). `excel-namebox.ps1 <out.json>` selects ranges under
names of each kind and records the selection's name Excel reports (Selection.Name), and where Application.Goto of a
name puts the selection, for `test/namebox-excel.json` (docs/NAMES.md, "The Name Box").

Dynamic arrays: `dynamic-cases.mjs` makes the cases, `excel-dynamic.ps1 <cases.json> <out.json> <books folder>` has Excel
365 enter each with Formula2 and save it, and `dynamic-ours.mjs <folder>` with `excel-dynamic-check.ps1 <folder>
<out.json>` has Excel read the editor's files the same way (docs/DYNAMIC.md).

Go To Special: `special-cases.mjs` makes the cases, `excel-special.ps1 <cases.json> test/fixtures/special.xlsx
test/special-excel.json` has Excel find each (the selection made, the choice asked as the box asks it, the result
selected and read; nothing saved) (docs/GOTOSPECIAL.md).

Sparklines: `excel-sparklines.ps1 <out.json> <books folder>` builds sparkline groups with each option in an invisible
Excel, saves them and keeps the saved extension and COM's view of each (`test/sparklines-xml-excel.json`), and has
Excel print them and sheets of other cell sizes and line weights to PDF. Excel prints each sparkline as a picture:
`node sparkimages.cjs <file.pdf> <outDir>` takes the pictures out of a PDF, and `python spark-features.py <outDir>
<out.json>` measures them (bars, coloured blobs, the line's trace). Run on Excel's PDF of
`test/fixtures/sparklines.xlsx` (`excel-print.ps1`), that gives `test/sparklines-excel.json` (docs/SPARKLINES.md).
`excel-sparkedit.ps1 <out.json> <books folder>` has Excel make sparklines from pairs of location and data and run the
Sparkline tab's edits through COM, each saved (`test/sparkedit-excel.json`); `excel-sparkedit-check.ps1 <folder>
<out.json>` has Excel read back the editor's files (`test/out/sparklines`, from `npm test` and `test/sparkline-tab.mjs`)
into `test/sparkedit-readback-excel.json`.
`excel-sparkstyles.ps1 <out.json> <books folder>` is the one script that shows an Excel window: a style can only be
applied from the ribbon's gallery, so it makes its own Excel visible and invokes each of the 36 gallery items through
UI Automation (no keys or mouse), saving a copy after each (`test/sparkstyles-excel.json`).

Tables: `excel-tablestyles.ps1 <out.json>` has Excel duplicate each of its 60 built-in table styles and keeps the dxfs it
saves for them (`test/tablestyles-excel.json`; `node tablestyles-data.mjs` turns them into `src/engine/tablestyles.js`).
`excel-tablelook.ps1 <cases.json> <out.json> [-AllStyles]` reads every cell's DisplayFormat of a table in a style with
the case's options (`tablelook-cases.json`, `test/tablelook-*-excel.json`); `excel-tints.ps1` the colour of 22 colours
at ten tints (`test/tints-excel.json`); `excel-tables.ps1 table-cases.json <out.json>` makes and changes tables and keeps
what Excel saves (`test/tablemake-excel.json`), and `excel-tables-check.ps1 test/out/tablemake <out.json>` has Excel read back the
editor's files of the same cases (`test/tablemake-readback-excel.json`); `excel-cellstyles.ps1 <out.json>` the built-in cell styles and what
applying each one saves (`test/cellstyles-excel.json`); `excel-newstyle.ps1 <out.json>` what New Cell Style (Styles.Add by
example) and a style's Modify save, the cells using it with them (`test/newstyle-excel.json`). The table drawing is compared with Excel's PDF of
`test/fixtures/tables.xlsx` by `compare-cover.py <excel shapes> <our shapes>`, which compares the colour showing at
points inside Excel's filled shapes (docs/TABLES.md).

PDF export (src/pdf.js) is measured the same way: `OXE_EXPORT=1 node tools/excel-compare/print-ours.mjs <books> <out>` makes
the PDFs through `exportPdf` instead of the browser's print (docs/PDF.md).

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

## Save fidelity

```sh
pwsh tools/excel-compare/excel-make-books.ps1 <folder>
node tools/excel-compare/fidelity.mjs <folder> <out>
pwsh tools/excel-compare/excel-fidelity.ps1 <folder> <out>
```

Workbooks made by Excel (or any folder of them) opened, edited and saved by the editor; the parts kept byte for byte
counted, then Excel compares every cell and the count of each feature in the original and the saved copy
(docs/FIDELITY.md).

Rich text typed into: `excel-richtype.ps1 test/richtype-excel.json` starts its own hidden Excel, gives a cell runs
(`Characters(start, length).Font`), adds or replaces text with `Characters(start, length).Insert`, and reads each
character's font back: which run's font the new text takes (docs/FORMATS.md, Rich text). Nothing saved.

Rich text formatted, copied and pasted: `node tools/excel-compare/richruns-cases.mjs`, then `pwsh -STA
tools/excel-compare/excel-richruns.ps1 tools/excel-compare/richruns-cases.json test/richruns-excel.json` starts its own
hidden Excel, sets part of a cell's text's font as Format Cells' Font tab does, copies a cell of rich text with Paste,
each Paste Special choice, a cut and into another workbook, pastes made-up HTML with mixed formatting in one cell (put
on the clipboard as CF_HTML, its text put back afterwards), and reads each character's font back with the strings'
XML of the workbook it saved in a temporary folder, removed after (docs/FORMATS.md, docs/CLIPBOARD.md).

## The workbooks

Real workbooks are kept out of the repository (`..\samples`), and referred to by label (`sample X014`), never by
name. The generated fixtures in `test/fixtures` can be used freely.
