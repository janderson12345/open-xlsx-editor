# Fill and the fill handle

How cells are filled (`src/ooxml/fillops.js`), and what Excel was seen to do. Measured over COM: 42 cases of
`Range.FillDown`, `FillRight`, `FillUp`, `FillLeft` and `Range.AutoFill` (the fill handle) on a new workbook in a
hidden instance (`tools/excel-compare/excel-steps.ps1` on `fill-cases.json`, from `fill-cases.mjs`, giving
`test/fill-excel.json`). `npm test` (`test/fill.test.js`) fills the same cells and expects Excel's in every cell
(formula and value, number format, bold, fill): 42 of 42. Filled by the editor, saved and opened in Excel
(`excel-readback.ps1`), 42 of 42 read back as Excel made them.

## In the editor

- The fill handle: the small square on the selection's bottom-right corner (a crosshair over it). Dragged down, up,
  right or left (the way the pointer is furthest out of the selection), the cells it will fill are outlined; let go,
  they are filled and selected. Ctrl held while letting go fills the other way: a single cell counts on, more cells
  are copied as they are (Excel's Ctrl-drag, from its documentation; COM doesn't reach the drag).
- A double click on the handle fills down to the last row of the region round the selection (as Ctrl+* finds it:
  grown while a filled cell touches it, side or corner), stopping before a cell already filled in the selection's own
  columns. Measured in Excel's own window (the handle double-clicked, the cells set up and read over COM;
  `test/fill-dblclick-excel.json`, 13 of 13): three values on the left and eight on the right fill eight rows, ten and
  four fill ten; a gap in the column beside stops it; data touching only at a corner further on, or two columns over
  through a filled cell, takes it on; a cell filled below in its own column stops it a row above; with an empty column
  between, nothing.
- After a fill by the handle, Excel's Auto Fill Options button at the filled cells' corner: Copy Cells, Fill Series,
  Fill Formatting Only, Fill Without Formatting, and for dates Fill Days, Fill Weekdays, Fill Months, Fill Years. A
  choice fills the cells again that way, one more step of the undo history; the button goes when the selection moves.
- Custom lists (`fillLists`, as Excel's Edit Custom Lists) go round as the day and month names do.
- Ctrl+D (Fill Down) and Ctrl+R (Fill Right), and Home > Editing > Fill: Down, Right, Up, Left.
- Each is one step of the undo history. Not on a read-only workbook; `features: { editing: false }` turns them off.

## Excel's rules

| Rule | Seen |
|---|---|
| Fill Down, Right, Up, Left copy the selection's first row (column; last for Up and Left) over the rest: values, formulas moved as a copy moves them, formats. One row selected (one column), Ctrl+D copies the row above (Ctrl+R the column to the left). | A1 `=B1+1` bold, A1:A3 down: `=B2+1`, `=B3+1`, bold. A2:B2 down under 7, `=A1+1`: 7, `=A2+1`. |
| The handle on one number copies it; with Fill Series (or Ctrl), it counts on by one. | 1: 1, 1, 1. Series: 2, 3, 4. |
| Two numbers go on by their difference; three or more by the straight line through them (least squares), shown to 15 digits. | 1, 3: 5, 7, 9. 5, 3: 1, −1. 0.1, 0.2: 0.3 ... 0.6. 2, 4, 7: 9.33333333333333, 11.8333333333333 ... |
| Text ending (or starting) with a number counts on, by the step of two such cells; a number in the middle is copied with the text. | `Item 1`: `Item 2`. `Item 1`, `Item 3`: `Item 5`. `1 apple`: `2 apple`. `a1b`: `a1b`. |
| Day and month names, short and long, go round their lists in the step given and in the same capitals; quarters go round four. | `Mon`: `Tue` ... `Sun`, `Mon`. `January`, `March`: `May`. `JAN`: `FEB`. `Q1`, `Qtr1`: up to 4, then 1. |
| A date goes on by a day, a time by an hour. Two dates a whole number of months apart (the day kept, or the month's last day) go on by months; otherwise by their days. | 1/30/2024: 1/31, 2/1. 1/1 and 1/8: weekly. 1/31 and 2/29/2024: 3/31, 4/30, 5/31. 9:00: 10:00. |
| Formulas are moved as a copy moves them. | `=B1*2`: `=B2*2`. `=$B$1+B1`: `=$B$1+B2`. |
| Cells of different kinds repeat as a block, each run of like cells going on in its own way; plain text is copied. | `a`, `1`: `a`, `2`, `a`, `3`. `x`, `y`: `x`, `y`, `x`. Across, `Jan`, `10`: `Feb`, `11`, `Mar`, `12`. |
| Each column (row) of the selection is filled on its own; filling upwards or leftwards runs the series backwards. | `1`, `3` beside `Item 7`, `Item 8`: 5, 7, 9 beside `Item 9` ... 5, 6 filled up: 1 to 4. |
| The copied cells' formats repeat over the new cells. Copy Cells repeats the values too; Formatting Only the formats; Without Formatting the series alone. | 1 bold, 2: bold every other cell. |

## Fill > Justify (measured)

Home > Editing > Fill > Justify rewraps the text of the selection's first column to the selection's width
(`ooxml/justifyops.js`, `editor/textwidth.js`). Measured with Excel's Range.Justify in a hidden Excel: 15 cases of
how it treats cells, then 36 paragraphs (Aptos Narrow, Calibri and Arial at 64, 89, 145 and 215px) for where lines
break (`test/justify-excel.json`).

| Excel's rule | Excel's example |
|---|---|
| Each run of filled cells is a paragraph: their text joined with a space, spaces at its start dropped, broken at spaces into lines as wide as the selection allows. A word wider than that has a line of its own. | `one`, `two`, `three`, `four` in a 64px column: `one two`, `three four`. |
| Spaces beyond the one a break takes stay at the end of the line. | `  lead and   inner  `, `next`: `lead and  `, `inner  `, `next`. |
| The blank rows after each paragraph stay as many; paragraphs below move with the lines above. Rows no longer needed, to the selection's end, are cleared. | `one`, blank, blank, `two`, `three`: `one`, blank, blank, `two three`. |
| The selection's width is all its columns'. | A:C, 192px: 4 lines where A alone gave 16. |
| Only the first column's values move; the other columns and every cell's format stay where they were. | A1 bold, A2 italic, `one`...`four` joined: A1 bold `one two`, A2 italic `three`. |
| Text needing more rows than selected goes on below, over the cells there, numbers too; Excel asks first, "Text will extend below range." | A1:A2 of a sentence: 9 rows, `below1` in A4 and 7 in A6 written over. |
| A number, TRUE or FALSE, an error or a formula in the first column: "Cannot justify cells containing numbers or formulas.", and nothing changes. A number typed as text (`'123`) is text. | `alpha`, 42, `beta`: refused. |

Where lines break depends on the font, and Excel measures text its own way, not as a browser draws it: each
character's advance at the font's size in whole pixels (15 for 11pt), rounded to a whole pixel, without kerning, a
line kept while at most the width less 4px. Fitted on the 36 paragraphs, this broke 32 exactly as Excel did (Aptos
Narrow 12 of 12, measured in its twin Morrow Sans Narrow, Calibri 8 of 12, Arial 12 of 12); the browser's own widths
times a factor did no better than 28. Excel's advances are its fonts' hinted ones, which this follows only roughly
(Aptos Narrow's digits are 7px at 15, not 8): a long paragraph can break a word differently. The browser check runs
the Calibri and Arial paragraphs through the Fill menu: 20 of 24, as fitted.

## Fill Days, Weekdays, Months and Years (Auto Fill Options for dates)

Measured over COM as the others (AutoFill's xlFillDays to xlFillYears: `fill2-cases.mjs`, `test/fill2-excel.json`,
15 of 15 cells as Excel's, `src/ooxml/filldates.js`):

| Rule | Seen |
|---|---|
| One date goes on by one unit a row: a day, the next weekday (Saturday and Sunday left out), a month or a year, the day kept or the month's last. | Thu 11/12/2026, weekdays: Fri, Mon, Tue... A Saturday: Mon. 1/31/2024, months: 2/29, 3/31, 4/30. 2/29/2024, years: 2/28/2025. |
| Dates the same number of units apart go on by it. | 3/1/2020 and 3/1/2022, years: 2024, 2026. 1/31 and 2/29, days: 29 days on. Mon, Tue, weekdays: Wed, Thu, Fri, Mon. |
| Dates whose steps differ, or that don't move in the unit, are copied as a block; two such dates instead go on by one month (year; a week for weekdays) each round. | 1/1, 1/2, 1/4, days: copied. 1/15, 1/20, 1/31, months: copied. 1/1 and 1/8, months: 2/1, 2/8, 3/1. Thu and Sat, weekdays: Thu, Fri a week on. |
| Fill Days and Weekdays drop the time of day. | 1/30/2024 9:30: 1/31/2024 0:00. |

Upwards and leftwards these choices take Excel's guess, as before (not measured).

## Not yet

Fill > Across Worksheets (sheets can't be grouped yet), a drag back into the selection (Excel clears those cells), and
custom lists measured in Excel (adding one over COM would change the user's own Excel settings; they go round as the
built-in lists, whose rules were measured).
Justify's 255-character limit on a line is not checked.
