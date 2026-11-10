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

## Not yet

Text orientation, shrink to fit, the Format Cells dialog, cell styles other than the three above,
pattern and gradient fills, and formatting whole rows or columns (a selection past the used range is cut to it: Excel
gives the row or column itself the format).
