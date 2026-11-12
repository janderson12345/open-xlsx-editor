# Cell formats

How the Home tab's commands change cells (`src/ooxml/formatops.js`) and how the change is written to the styles part
(`src/ooxml/stylesedit.js`). Each rule was measured by running Excel's own ribbon commands over COM
(`CommandBars.ExecuteMso` on a new workbook in a hidden instance; no keys sent, nothing saved but copies outside the
repository) and reading back what Excel showed and what it saved. The workbooks written here were then opened in Excel,
which reads back every font, fill, alignment, border, number format and cell style as written. `npm test`
(`test/formats.test.js`) repeats each step and expects Excel's result.

## How a change is stored

- A cell's format is an entry of `cellXfs`; the change makes a copy of the cell's entry with the change in it, and of
  the font, fill or border it points to. An equal entry already there is used instead, so formatting many cells with
  the same look adds one entry; a new one goes at the end and no index a cell uses ever moves.
- A formatted empty cell is kept as a `c` element with only its `s`, as Excel keeps it.
- New lines are automatic black (`indexed="64"`), as Excel writes the menu's borders.
- A workbook without a styles part gets one the first time it is formatted (Calibri 11, the two fills Excel always
  keeps, the Normal style).
- Rows without a fixed height grow when their font does: Excel fits them when it opens the file (20pt text: 26.25pt).

## Excel's rules

| Rule | Seen |
|---|---|
| Bold, Italic, Underline, Strikethrough and Wrap Text follow the active cell: on for every selected cell unless the active cell has it, then off for all. | A1 plain, B1 bold: both bold. A1 bold: both plain, whichever else is bold. |
| Underline is single. | `Font.Underline` 2. |
| Grow and Shrink Font set every selected cell to one step from the active cell's size, through 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72, and no further. A size between steps goes to the next one. | From 11: 12, 14 ... 72, 72. Down: 10, 9, 8, 8. 13 grows to 14. 11 and 20 both become 12. |
| Align Left, Center, Right, Top, Middle and Bottom set the alignment; pressing one again leaves it. | Center twice: still centred (through ExecuteMso; to be checked by hand on the ribbon). |
| Increase Indent adds a level and makes the cell left-aligned, unless it is right-aligned. Decrease Indent takes one away and leaves the alignment. | General or centred: left, 1. Right: right, 1. Up twice then down: 1, left. |
| A border goes on the cells along the selection's edge (Bottom: its bottom row). Inside lines change only with All Borders and No Border. | B2:C3 Outside: B2 left and top, C3 bottom and right. |
| A border drawn on the selection's edge takes away the neighbour's facing side, unless it is the same line. No Border takes them away too. | Top on B2: B1's medium bottom goes, a thin one stays. No Border on B2: B1's bottom and C2's left go. |
| Thick Outside Borders and Thick Bottom Border are stored as `medium`; Bottom Double Border as `double`. | Excel's saved copy. |
| Percent Style, Comma Style and the $ button (Accounting Number Format) apply Excel's built-in cell styles Percent (`builtinId` 5, format 9), Comma (3, format 43) and Currency (4, format 44). The style is added to `cellStyleXfs` and `cellStyles` when the workbook hasn't got it, and the codes of formats 43 and 44 are written out. | `Style.Name` Percent, Comma, Currency; Excel's saved `styles.xml`. |
| Increase and Decrease Decimal change the active cell's format and give it to every selected cell. From General they start from the decimals the cell shows; each section of a format gets one place more or fewer; a date, a time, text or a General cell that isn't a number doesn't change. | 1.5: `0.00`, `0.000`. 2.123456789 (shown 2.123457) down: `0.00000`. `0.00%` down: `0.0%`, `0%`. `#,##0.00 "kg";[Red]-#,##0.00` up: both sections. `m/d/yyyy` unchanged. |
| Font Color and Fill Color start as red (`FFFF0000`) and yellow (`FFFFFF00`, solid, background `indexed="64"`). | `Font.Color` 255, `Interior.Color` 65535. |

## Merging (`src/ooxml/mergeops.js`)

| Rule | Seen |
|---|---|
| Merging keeps one value, the first in reading order, and puts it in the top-left cell; the others are emptied. Excel asks first: "Merging cells only keeps the upper-left value and discards other values." | A1, B1, A2 filled in, A1:B2 merged: `a1` in A1, the rest empty. A1 and B1 empty: A2's `a2` moves to A1. |
| Every cell of the area gets the format of the cell whose value is kept (Excel writes them all with it). | A1 filled yellow, B2 bold: all four filled, none bold. A2's value kept: all four plain. |
| Merge & Center also centres; Merge Cells and Merge Across leave the alignment. Merge Across merges each row of the selection on its own. | A1:B2 across: A1:B1 and A2:B2, A2 keeping `a2`. |
| A selection that overlaps a merged range is merged with it into one. | A1:B2 merged, then B2:C3: A1:C3. |
| Merge & Center on a selection with a merged range in it unmerges, and sets the alignment back to General. Unmerge Cells leaves the alignment. | A1 after Merge & Center, pressed again: not merged, General. Unmerged: still centred. |

## Clear (`src/ooxml/clearops.js`)

Measured over COM: `Range.Clear`, `ClearContents` and `ClearFormats` on a small table with a formula, formats, a
border and two merged ranges, six cases (`tools/excel-compare/excel-steps.ps1` on `clear-cases.json`, from
`clear-cases.mjs`, giving `test/clear-excel.json`). `test/clear.test.js` expects Excel's cells and merged ranges
in each; cleared here and opened in Excel, the three that clear read back as Excel made them.

| Rule | Seen |
|---|---|
| Clear All empties the cells and gives them the Normal format; merged ranges inside are unmerged. | B2:E8: values, bold, fill and border gone, D8:E8 unmerged. |
| Clear Contents (and Delete) empties them; formats and merged ranges stay. | B2 keeps bold and 0.00. |
| Clear Formats gives them the Normal format and unmerges; the values stay. | B3 `=B2*2` kept, plain. |
| A selection over part of a merged range is refused, whichever Clear: "We can't do that to a merged cell." | B9:C10 with A9:B9 merged. |

A click on a merged cell selects the whole merged range in Excel; here a selection of one of its cells clears the
whole range in the same way.

### Comments and Notes, Hyperlinks (`src/ooxml/clearparts.js`)

Measured over COM: `Range.ClearComments`, `ClearHyperlinks`, `Hyperlinks.Delete` (Remove Hyperlinks), `Clear` and
`ClearContents` on a sheet with a hyperlink to an address in A1 (bold), a note in B1, a threaded comment in C1 and one
hyperlink over D1:E2, and the files Excel then saved. `test/clearparts.test.js` expects the same; cleared here and
opened in Excel, its hyperlinks, notes and comments are those.

| Rule | Seen |
|---|---|
| Clear Comments and Notes takes the notes and threaded comments (with their replies) of the cells, nothing else. | B1 alone: its note goes, C1's thread stays. A1:E2: the comments, drawing and threaded comments parts go. |
| Clear Hyperlinks takes the hyperlinks meeting the cells, each whole, and leaves the formats (the Hyperlink style stays). | D1 alone: D1:E2's hyperlink goes; A1's stays. |
| Remove Hyperlinks does that and takes the formats of the hyperlinks' cells away, theirs only. | A1 loses its bold, E2 its style; a bold F1 beside them stays bold. |
| Clear All takes them all with the cells; Clear Contents takes the hyperlinks (their formats stay) and leaves the notes; Clear Formats leaves both. | |

Excel drops a comments part it leaves empty; here it stays, empty, as after Delete Note.

## Format Cells (`src/editor/formatbox.js`, `formatops.js` formatCells)

Ctrl+1 opens Excel's box on the active cell's format; only what is changed in it is given to the selection, each other
setting staying as each cell has it. How Excel saves each setting was measured over COM
(`tools/excel-compare/excel-formatcells.ps1`: one cell per setting on a new workbook, saved, its `xf`, font, fill and
border read back; `test/formatcells-excel.json`). `test/formatcells.test.js` makes all 43 here and expects Excel's text.

| Rule | Seen |
|---|---|
| Rotation is `textRotation`: 1 to 90 degrees up as themselves, -1 to -90 as 91 to 180, stacked letters 255. | 45: 45; -90: 180; vertical text: 255. |
| An indented or Fill-aligned cell isn't rotated: Excel drops the rotation, whichever was set first. Centred text keeps it. | Indent 2 then 30 degrees, or 30 then indent: `horizontal="left" indent="2"`, no rotation. |
| Wrap text and shrink to fit can both be on in the file (the box offers one or the other). | `wrapText="1" shrinkToFit="1"`. |
| The alignment's attributes go in Excel's order: horizontal, vertical, textRotation, wrapText, indent, shrinkToFit, readingOrder. | `horizontal="left" indent="2" shrinkToFit="1"`. |
| A gradient is `gradientFill` with its stops; degree 0 is left out; a path gradient has `type="path"` and its rectangle. | `<gradientFill degree="90"><stop position="0">...` |
| A pattern's own colour is `fgColor`, the cell's colour `bgColor`; a shade without colours has none. | Reverse diagonal stripe, red on yellow: `fgColor` FF0000, `bgColor` FFFF00; 50% gray: `mediumGray` alone. |
| Diagonals are `diagonalUp` and `diagonalDown` on the border (in that order) and one `diagonal` line for both. | `<border diagonalUp="1" diagonalDown="1">...<diagonal style="thin">` |
| Superscript and subscript are the font's `vertAlign`; the accounting underlines `singleAccounting`, `doubleAccounting`. | |
| Protection: `locked="0"` when unlocked, `hidden="1"` when hidden, nothing when locked and shown. | |
| Over a range, the outline goes on its edge and the inside lines on both facing sides of each pair of cells. | D2:F4: E3 dashed all round, D2 thin left and top. |

Rows grow for turned text as Excel's AutoFit grows them (Calibri 11, measured with `Rows.AutoFit`): stacked letters a
line each and 1.5pt ('ab' 31.5, 'Stack' 76.5); text at an angle by its width up the angle and a line across it, within
2pt of Excel's 36 to 81pt for 15 to 90 degrees.

Not settled: a range's border over COM also took a matching thin line off the neighbour beyond its edge (G3's left
beside F3's new right), where the ribbon's Borders measured earlier keeps a matching line; this follows the ribbon.
The degrees the box's own Fill Effects variants give, and the font sizes Excel steps through when shrinking, are not
yet measured.

## A row's or a column's own format

Measured over COM (`Columns(2).Font.Bold`, `Range("D:E").Interior.Color`, `Rows(5).Font.Italic`, `Rows(7).Font.Size`,
and the sheet Excel saved; test/linestyles.test.js): a formatted column is `<col min max style="X">` (with the
default width), a formatted row `<row s="X" customFormat="1">`; the cells already there get the change, and where a
formatted row crosses a formatted column Excel makes an empty cell with both (B5: bold and italic). A row's font
made larger grows it (ht="18.75" for 14pt).

Here (`src/ooxml/sheet.js` lineStyle): an empty cell has its row's own format, else its column's; its fill and
borders are drawn from it, a cell typed there takes it (`src/ooxml/sheetedit.js`), and Format Cells > Protection on
a column counts for a protected sheet. Formatting whole rows or columns here still writes the used range's cells
only (Not yet, below).

## Rich text

A cell's text in runs of several fonts (`src/ooxml/strings.js` stringItemRuns, drawn by `src/editor/cells.js`): each
run in its own font over the cell's, bold, italic, underline, strikethrough, size, colour, font and super- or
subscript. Measured over COM (`Characters(start, length).Font` set, the sharedStrings Excel saved,
test/richtext.test.js): each run an `<r>` with an `<rPr>` giving every property (the font's name as `rFont`), the
first run without one when it is the cell's own font; the cell's format takes the first run's font (a struck first
word made the cell's font struck). Shown as Excel shows them; Show Formulas, and a number format that changes the
text, show it plain.

Not yet: editing a cell's runs (typing replaces them with plain text, as a new value), copying them, and a row's
height following a larger run (Excel's file says the height, so a file from Excel shows right; a row grown here
doesn't).

## Format Painter

Format Painter copies the selection (outlined, as a copy is, the system clipboard left alone) and pastes its formats
on the next cells selected, through Paste Special > Formats (measured for `test/paste-excel.json`). Pressed twice
quickly it stays on; Esc, or pressing it again, ends it.

## Not yet

Cell styles other than the three above, and formatting whole rows or columns (a selection past the used range is cut to it: Excel
gives the row or column itself the format).
