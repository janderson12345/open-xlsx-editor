# Copy, cut, paste and Paste Special

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

## Paste Special and Paste Link

Measured over COM: 30 cases of `Range.Copy` then `PasteSpecial` (what to paste, the operation, Skip blanks,
Transpose) or `Worksheet.Paste` with Link, on two small sheets: numbers, text, a blank, relative, $ and outside
formulas, a percentage, a filled, bold, bordered cell and a wide column copied onto cells with their own values,
formulas, text, blanks and formats (`tools/excel-compare/excel-steps.ps1` on `paste-cases.json`, from
`paste-cases.mjs`, giving `test/paste-excel.json`). `test/paste.test.js` does each and expects Excel's cells
(formula, value, number format, bold, fill, border), merged ranges and widths: 30 of 30. Saved by the editor and
opened in Excel (`excel-readback.ps1`), 30 of 30 read back as Excel made them.

- Ctrl+Alt+V, Home > Paste ▾ > Paste Special..., or a cell's right-click menu: Excel's box (Paste, Operation, Skip
  blanks, Transpose, Paste Link). Paste ▾ also has Formulas, Values, Formatting, Transpose and Paste Link at once.
- For cells copied in the editor (from this workbook or another open here; formats only within one workbook). Not
  after a cut, as in Excel; not for text from another program.
- One step of the undo history; the copy stays to be pasted again.

| Rule | Seen |
|---|---|
| Formulas: formulas (moved as a paste moves them) and values, the destination's formats kept. Values: what each cell shows, formulas worked out. Formats: the formats alone. All except borders: everything, the destination's borders kept. | A2 `=A1*2` to E2: `=E1*2` (Formulas); 2 (Values). F1 filled stays filled under Formulas and Values. |
| Formulas and number formats, Values and number formats: the number format alone joins the destination's other formats. | E3 bold, A3 0% pasted: bold and 0%. |
| Column widths: the widths alone. | Column A 20 wide: E 20 wide, nothing else changed. |
| An empty cell of the copy empties its destination, and its format goes with All; Skip blanks leaves the destination as it was, format too. | C2 blank onto G2 50 (0.0): G2 empty; with Skip blanks, 50 in 0.0. |
| Add, Subtract, Multiply, Divide: the destination's number and the copy's worked out; a blank counts as 0; text on either side is left as it was; dividing by 0 gives `#DIV/0!`. | 10 + 1: 11. 50 × blank: 0. 50 ÷ blank: `#DIV/0!`. `x` + `t`: `x`. `y` + 3: `y`. |
| With a formula on either side, the result is a formula: the destination's in brackets, the copy's in brackets, a number as it is. | `=E1+1` + 3: `=(E1+1)+3`. 40 − `=E1*2`: `=40-(E1*2)`. Blank + `=H9*2`: `=0+(H9*2)`. `=TRUE` + `=F1+H9`: `=(TRUE)+(F1+H9)`. |
| An operation with Formats changes nothing but formats. | Formats, Add: the formats only. |
| Transpose: rows become columns. A reference to a cell of the copy, $ or not, follows that cell to where it went; one outside keeps its $ parts and its relative parts turn with the copy. | A1:C3 to E1: B2 `=$A$1+B1` is F2 `=$E$1+E2`; B3 `=SUM(A1:B1)` is G2 `=SUM(E1:E2)`; B2 `=B1+D9` is F2 `=E2+M4`; `=$D$9+1` stays. |
| A selection that is a whole number of copies takes the copy that many times, as Paste does. | A1:C3 values onto E10:G15: twice. |
| Paste Link: a formula for each cell reading the copied cell as far from the copy's top-left, relative; for one cell, absolute. Over a whole number of copies, the links carry on across the selection rather than repeating. Cells in the General format take the number format of the copy's first cell. | A1:C3 to E1: `=A1` ... `=C3`, each 0.00 as A1. A1 to E5: `=$A$1`. B1:C1 onto E5:H6: `=B1` to `=E2`. |

Not measured: an operation with a TRUE or FALSE constant (left as it is here, as text is), Transpose of a reference
whose $ is on one part only (each part turned as above), column widths onto columns that had their own width (they
keep it here), and merged ranges with Transpose (not pasted).

## Notes pasted

A copy carries its cells' notes (with their authors). Measured over COM (B2:C3 with a note on B2 copied; Paste,
PasteSpecial xlPasteAll, xlPasteComments and xlPasteValues onto cells, some with notes of their own): Paste, All and
Comments and Notes leave each pasted cell with the copy's note there, or none (a note there before is taken away, so
F6's went when C3, which had none, was pasted over it); Values brings none. Here (`clipops.js` pasteNotes): Ctrl+V,
Paste Special's All, All except borders and Comments and Notes do the same, over a repeated paste too,
as one step of Undo with the cells (test/pastenotes.test.js, test/editing-formats.mjs).

Comment threads too: measured over COM (B2 a thread with a reply, C2 a resolved one, B2:C2 pasted with Paste,
PasteSpecial xlPasteComments, and cut), Excel writes each pasted thread with its comments' text, authors and times,
`done="1"` kept, under new ids; a cut leaves none behind. The same here, with threadedit.js's own writing.

## Validation pasted

A copy carries the validation rules over its cells. Measured over COM (C2:C3 a custom rule =C2>B2, B2:B3 a list
=$A$1:$A$3, B2:C3 copied; Paste onto H5, PasteSpecial xlPasteValidation onto L5 and onto O5:R8, the file saved):
on the sheet copied from, the pasted cells join the rule itself, its formula as it was (written for its first cell,
so I5 reads I5>H5), and ranges that carry on one another are joined (`C2:C3 I5:I6 M5:M6 P5:P8 R5:R8`); a rule of
the pasted-over cells loses them. Here (`clipops.js` pasteValidation, `dvedit.js` addToRule): Ctrl+V, Paste Special's
All, All except borders and Validation do the same, Excel's file to the attribute (test/pastenotes.test.js); onto
another sheet, a rule of its own, its formula moved as a copied formula is (*measured*: B2:C3 to another sheet's H5,
`I5>H5` over I5:I6; a list `=$A$1:$A$3` and `$E$1` stay as written, no sheet named, so they read the new sheet's
cells; a rule there over H5:K8 keeps the cells not pasted over, J5:K8 H7:I8).

With Transpose (*measured*: A1:A3 or A1:B2 with a whole-number, custom or list rule, PasteSpecial All with Transpose at
C1 or D1, the file saved): the same over the turned cells. On the sheet copied from they join the rule (`A1:A3
C1:E1`, its formula as it was; A1:B2's two columns to rows 1 and 2, `D1:E1` and `D2:E2`), the pasted-over cells leave
their rules (C1:F1's list keeps F1); onto another sheet a rule of its own, moved by the corner only (=A1>B1 over A1:A3
turned to C1: =C1>D1 over C1:E1). Six cases, all the same here.

## A cut's notes and validation

Measured over COM (B2:C3 cut to E5, B2 with a note, C2:C4 a custom rule =C2>B2, F6 a note and E5 a rule of their
own): the notes move with their cells (B2's to E5, F6's gone); each rule's cut cells go with them, after its other
cells, and the formula, written for the rule's first cell, is rewritten for the new one (`C4 F5:F6`, =C4>B4); E5's
rule, pasted over, goes. Here (`clipops.js` moveCutParts, `dvedit.js` moveValidation) a cut does the same, Excel's
file to the attribute (test/pastenotes.test.js); to another sheet, the rules are taken off the cut cells and pasted
there as a copy's are (*measured*: the same rules there as a copy's, the cut cells' rules gone).

Cells taken out of a rule any other way (Data Validation's Clear All, a paste over them) rebase its formulas too when
its first cell changes, as Excel's cut did.

## Conditional formats (measured)

Conditional formats come with their cells (`ooxml/clipcf.js`). Measured over COM in a hidden Excel: a copy and
PasteSpecial (or a cut) in each case, then the conditionalFormatting in the file Excel saved; 22 cases, all the same
here (`test/pastecf.test.js`).

| Excel's rule | Excel's example |
|---|---|
| Paste, All, Formats, All except borders, All using Source theme and All merging conditional formats bring them, with Skip blanks or an operation too. Values, Formulas, the number-format pairs and Validation don't. | Values at C1: no rule there. |
| The pasted cells first leave every rule over them; All merging conditional formats leaves those rules as they were. | A1:A3 pasted at E1 over E1:E4's rule: E4 keeps it; merging: E1:E4 keeps it. |
| A rule that loses its first cell is written for its new one. | =E1>F1 over E1:E4, E1:E2 pasted over: =E3>F3 over E3:E4. |
| Each conditionalFormatting the copy meets comes as a new one over the pasted cells: all its rules, their formats copied, their formulas moved as a copied formula moves, first in priority (the sheet's others that many later). | "> 5" and =A1>B1 over A1:A3, pasted at C1: C1:C3, =C1>D1, priorities 1 and 2; A1:A3's 3 and 4. |
| On the sheet copied from, a rule over one row or one column, pasted along that line against or over it, takes the cells in instead, its formulas written for its new first cell. Otherwise a copy of its own. | A1:A3 at A4: A1:A6; A3:A4 at A1 of A3:A5: A1:A5 (=A1>B1); A1:C1 at D1: A1:F1. At B1, A5, B4: copies. A 2 by 2 rule: always a copy. |
| With Transpose, the turned cells, the formulas moved by the corner only. | A1:A3 at C1: C1:E1, =C1>D1. |
| To another sheet: a copy there; the sheet copied from as it was. | Sheet2!C1:C3, =C1>D1. |
| A cut moves the rules over the cells cut, their formulas as a cut formula's. | =A1>B1 over A1:A3, cut to C1: C1:C3, =C1>B1. |

Paste Special's box now has Excel's twelve choices, All using Source theme and All merging conditional formats
among them; within one workbook the first pastes as All does (its theme is the same).

Not yet: rules with Excel 2010's settings in the extLst (data bars, its icon sets) aren't brought; a cut of part of
a rule leaves the rule where it was (not measured).

## Not yet

Paste
Link to another workbook; Paste Special of text or HTML from another program; the Format Painter, pasting the formats of another
program's cells (only its text is read), defined names following a cut, and a range partly over the cells a cut was
pasted onto (Excel shrinks it: `=SUM(A1:A3)` became `=SUM(A2:A3)` when A1 was moved onto A3).
