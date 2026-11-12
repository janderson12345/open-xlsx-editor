# Protect Sheet

Review > Protect Sheet and Unprotect Sheet (`src/ooxml/protection.js`, `src/editor/protectbox.js`), the changes a
protected sheet refuses (`allowed` in `src/editor/commands.js`), and what Excel was seen to do. Measured over COM:
`Worksheet.Protect` with each of its options and a password, `Unprotect`, and the files Excel saved
(test/protection.test.js). A file protected here opens in Excel protected the same way, its password working there.

## In the editor

- Review > Protect Sheet opens Excel's box: Protect worksheet and contents of locked cells, Password to unprotect
  sheet, and Allow all users of this worksheet to: Select locked cells, Select unlocked cells, Format cells, Format
  columns, Format rows, Insert columns, Insert rows, Insert hyperlinks, Delete columns, Delete rows, Sort, Use
  AutoFilter, Use PivotTable & PivotChart, Edit objects, Edit scenarios (the first two ticked, as in Excel). With a
  password, Confirm Password asks for it again ("Confirmation password is not identical." when it isn't).
- The button then reads Unprotect Sheet; with a password it asks for it, and a wrong one gets Excel's words.
- On a protected sheet, a change to a locked cell (typing, paste, clear, fill, Paste Special, the agent's) is refused
  with Excel's message: "The cell or chart you're trying to change is on a protected sheet. To make a change,
  unprotect the sheet. You might be requested to enter a password." Unlocked cells (Format Cells > Protection) change.
- Formatting cells, column widths and row heights (hiding too), inserting and deleting rows and columns (and cells),
  sorting and filtering are refused unless the box allowed them; turning the filter on or off is refused while the
  sheet is protected. Sorting also needs the cells unlocked, as in Excel.
- Protecting and unprotecting aren't steps of Undo.

## Excel's rules

| Rule | Seen |
|---|---|
| Protected without a password: `sheet`, `objects` and `scenarios`, after `sheetData`. | `<sheetProtection sheet="1" objects="1" scenarios="1"/>` |
| Each thing allowed is its attribute "0", in the schema's order. | `formatCells="0" formatColumns="0" formatRows="0" insertColumns="0" insertRows="0" insertHyperlinks="0" deleteColumns="0" deleteRows="0" sort="0" autoFilter="0" pivotTables="0"` |
| The selection the other way round: "1" where it is not allowed. | Unlocked cells only: `selectLockedCells="1"`; none: `selectLockedCells="1" selectUnlockedCells="1"` |
| Edit objects and Edit scenarios allowed: their attribute left out. | `<sheetProtection sheet="1" scenarios="1"/>` |
| A password: SHA-512 of a 16-byte salt and the password (UTF-16LE), hashed again 100,000 times with the round's number (4 bytes, little-endian) after it; base64. | Test-pass 1 with Excel's salt `elg2Pkx22aHeJJEnE+FfVA==`: Excel's own hash, to the byte. |
| Unprotecting removes the element. | No `sheetProtection`. |
| A locked cell changed: refused. | "The cell or chart you're trying to change is on a protected sheet. ..." |

Files from before Excel 2013 keep a 16-bit hash in `password`; it is read (test: "test" is CBEB), not written.

## Not yet

- Selecting: Select locked cells and Select unlocked cells are written and kept but not enforced (every cell can be
  selected). Insert hyperlinks, PivotTables, objects (notes, charts, pictures) and scenarios aren't enforced either.
- Excel refuses as soon as a locked cell is typed into; here the entry is refused when it is entered.
- Protect Workbook (structure: sheets added, deleted, renamed, moved), Allow Edit Ranges (`protectedRanges`), and a
  cell's Hidden (its formula not shown in the formula bar while the sheet is protected).
