# Copy, cut and paste

How cells are copied, cut and pasted (`src/ooxml/clipops.js`, references moved by `moveRefs` in
`src/calc/tokens.js`), and what Excel was seen to do. Each rule was measured by doing the same thing in Excel over COM
(`Range.Copy` or `Range.Cut`, then `Worksheet.Paste`, on a new workbook in a hidden instance). `npm test`
(`test/clipboard.test.js`) repeats each step and expects Excel's result; a workbook pasted into here opens in Excel
with the same formulas and values.

## In the editor

- Ctrl+C, Ctrl+X and Ctrl+V on the sheet, the browser's own Copy, Cut and Paste, and the Home tab's Clipboard group.
- What was copied or cut is outlined (a moving dashed line, still when the system asks for less motion) until it is
  pasted (a cut) or let go: Esc, typing in a cell, or a change to the sheets.
- The system clipboard gets the cells' text as shown, tabs between cells and a line for each row, as Excel puts it,
  so it pastes into other programs. Pasting it back here pastes the cells themselves (formulas, formats, merged
  ranges); any other text is read cell by cell as if typed there (`=` starts a formula, `50%` is 0.5).
- A paste is one step of the undo history, whatever sheets it changed.
- Browsers send no copy or paste event to a sheet with no text selected, so the keys take the cells themselves; the
  system clipboard is then read and written through the async clipboard, where the page may use it.

## Excel's rules

| Rule | Seen |
|---|---|
| A copied formula's relative references move with it, its $ ones stay. A reference that would go off the sheet becomes `#REF!`. | B2 `=A2*2` to D5: `=C5*2`. To A1: `=#REF!*2`. B3 `=$A$1+A3` to another sheet's C3: `=$A$1+B3`. |
| A copy brings its cells' formats and merged ranges, not the column widths. | D5 bold, as B2. A1:B1 merged, copied to D1: D1:E1 merged. Column A 20 wide, E stays 8.43. |
| Pasted over a selection that is a whole number of copies in both directions, the copy repeats over it; otherwise it is pasted once from the selection's top-left cell. | B2:B3 into D5:D8: twice. Into D5:D7: once, D7 untouched. B2 into D5:E6: four times. |
| A cut moves its cells, formats and merged ranges; their places are left empty and unformatted. | A2:B2 cut to C5: D5 `=C5*2` and bold, B2 plain. |
| After a cut, every formula in the workbook that reads the moved cells follows them, on any sheet, when the whole reference was inside what was cut; a range that only partly was stays as it was. | A1:A2 cut to C1: `=A1+A2` is `=C1+C2`, `=SUM(A1:A2)` is `=SUM(C1:C2)`, `=SUM(A1:A3)` stays; another sheet's `=Sheet1!A1` is `=Sheet1!C1`. |
| A reference to a single cell that a cut was pasted over becomes `#REF!`. | A1 cut to A3: `=$A$1+A3` is `=$A$3+#REF!`. |

## Not yet

Paste Special (values only, formats only, transpose...), the Format Painter, pasting the formats of another
program's cells (only its text is read), defined names following a cut, and a range partly over the cells a cut was
pasted onto (Excel shrinks it: `=SUM(A1:A3)` became `=SUM(A2:A3)` when A1 was moved onto A3).
