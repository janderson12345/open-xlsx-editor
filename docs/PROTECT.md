# Protect Sheet and Protect Workbook

Review > Protect Sheet and Unprotect Sheet (`src/ooxml/protection.js`, `src/editor/protectbox.js`), the changes a
protected sheet refuses (`allowed` in `src/editor/commands.js`), and what Excel was seen to do. Measured over COM:
`Worksheet.Protect` with each of its options and a password, `Unprotect`, and the files Excel saved
(test/protection.test.js; in the browser, test/editing-protect.mjs). A file protected here opens in Excel protected the same way, its password working there.

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
- A cell whose format has Hidden (Format Cells > Protection) shows nothing in the formula bar while the sheet is
  protected, a formula or a constant. With Show Formulas its formula shows as nothing (a hidden constant still shows
  itself), and FORMULATEXT of it gives #N/A while ISFORMULA still gives TRUE. Measured over COM: `Range.Formula` gives
  "" for both, `Range.Text` with `DisplayFormulas` gives "" for the formula and the text for the constant, FORMULATEXT
  #N/A after the next calculation (Excel writes it `ca="1"`, so it is volatile here). Unprotected, all shows again.
- The rest of the sheet, as Excel was seen to refuse it over COM on a protected sheet:

  | Change | Refused unless | Here |
  |---|---|---|
  | New note | Edit objects allowed (locked cell or not) | Refused before its box opens |
  | A note edited or deleted | Never refused | Goes through |
  | New chart | Always refused, Edit objects allowed or not | Refused |
  | A chart moved, sized, changed or deleted | Edit objects allowed | Refused otherwise |
  | Conditional formatting | Format cells allowed | Refused otherwise (new rules, Clear Rules, the Rules Manager) |
  | Data Validation, Ignore Error | Always refused | Refused |
  | Hyperlink removed (Clear or Remove Hyperlinks) | Its cell unlocked | Refused on a locked cell |
  | A PivotTable refreshed | Always refused (its cells locked), Use PivotTable allowed or not | Refused |
  | A PivotTable's fields moved | Use PivotTable & PivotChart allowed | Refused otherwise |
  | Advanced Filter in place | Never refused | Goes through |

  A hyperlink inserted (Insert > Link) needs Insert hyperlinks allowed and its cell unlocked, as Excel's did.
  Clear All of locked cells with notes or hyperlinks is refused as typing there is.
- Typing a date, a percentage or a HYPERLINK formula into an unlocked cell of a protected sheet that doesn't allow
  Format cells enters it without the format Excel would give it (refused whole before); what Excel does there wasn't
  measured.
- Protecting and unprotecting aren't steps of Undo.

## Protect Workbook

Review > Protect Workbook opens Excel's Protect Structure and Windows box (Structure ticked; Windows can't be, as in
Excel 365; Password (optional)), and Confirm Password with a password. The button stays pressed while the workbook is
protected; pressed again, it unprotects it (asking for the password if there is one).

While the structure is protected, the sheets can't be inserted, copied, deleted, renamed, moved, hidden, shown or
coloured, and a PivotTable can't be put on a new sheet: refused with "Workbook is protected and cannot be changed."
(Excel greys those commands out; its words for the refusal couldn't be read over COM, which gives a general error.)
Cells can still be changed, as in Excel.

Measured over COM (`Workbook.Protect`, `Unprotect`, and the files Excel saved): `<workbookProtection
lockStructure="1"/>` after `workbookPr`; with a password, `workbookAlgorithmName="SHA-512"`, `workbookHashValue`,
`workbookSaltValue` and `workbookSpinCount="100000"` before it, hashed as a sheet's password is (Excel's own hash
reproduced to the byte). Excel refused adding, renaming, deleting, moving, hiding and colouring a sheet. A workbook
protected here opens in Excel protected: a new sheet refused, a wrong password refused, the right one unprotecting it.
Without a password any password unprotects it, in Excel and here.

## Allow Edit Ranges

`src/ooxml/editranges.js`: ranges of a sheet whose cells can be changed while it is protected, though they are locked.
Measured over COM (`Protection.AllowEditRanges.Add` with and without a password, cells changed on the protected
sheet, and the file Excel saved):

- `<protectedRanges>` right after `<sheetProtection>`, a `<protectedRange sqref="B2:C3" name="Open"/>` for each,
  several areas apart by spaces (`sqref="G1:G2 H4"`); a password as a sheet's (`algorithmName`, `hashValue`,
  `saltValue`, `spinCount` before `sqref`), Excel's own hash reproduced to the byte.
- A range without a password: its cells change at once. A change reaching past it (B2:D2 over B2:C3) is refused, in
  the protected sheet's words.
- A range with a password: refused until it is unlocked with its password; a wrong one is refused. Opened again,
  it is locked again (Excel doesn't keep the unlocking, and neither does this editor).
- Adding a range while the sheet is protected: refused by Excel.
- A file with ranges written here opens in Excel with them: the open range changes, the password range is refused
  until its password unlocks it.

In the editor, Review > Allow Edit Ranges opens Excel's Allow Users to Edit Ranges box: "Ranges unlocked by a
password when sheet is protected:", each range's Title and Refers to cells (=$B$2:$C$3,$H$4), and New..., Modify...,
Delete, which can't be pressed while the sheet is protected (Excel greys them out). New Range offers Range1 (the first
RangeN not taken) and the selection's cells; Range password, with Confirm Password; Modify Range keeps the password
when it is left blank. Changes there aren't steps of Undo, as protecting isn't. On the protected sheet, typing or
pasting into a range with a password asks for it in Excel's Unlock Range box ("A cell you are trying to change is
password protected." / "Enter the password to change this cell:"), and the change is made once it is right; a wrong
one gets Excel's words. test/editranges.test.js; in the browser, test/editing-protect.mjs.

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

## Selecting on a protected sheet

Select locked cells and Select unlocked cells (Protect Sheet's first two choices) are kept as Excel keeps them, and the
keys and clicks go as Excel's do, measured in its own window (keys pressed and cells clicked there, the active cell read
over COM after each; five cells unlocked: B2, C2, B3, C3, E5; `test/protsel.test.js`):

| Allowed | Arrows | Tab, Shift+Tab | Enter, Shift+Enter | A click |
|---|---|---|---|---|
| Both (Excel's default) | as usual | the next (last) unlocked cell in reading order, wrapping round: E5 to B2 | as usual | as usual |
| Unlocked cells only | Right and Left as Tab; Down and Up the next (last) unlocked cell down the columns, wrapping: B2, B3, C2 | as above | down the columns: E5 to B2 | on a locked cell, nothing |
| Neither | nothing | as above: B2 to C2 | down the columns: C2 to C3 | nothing |

## Not yet

- (Selecting is enforced as Excel does it: below.) Scenarios aren't enforced (there are none here).
- Comments (threads) on a protected sheet aren't stopped; Excel's rule for them wasn't measured. Nor is a new
  PivotTable put on a protected sheet's cells.
- Excel refuses as soon as a locked cell is typed into; here the entry is refused when it is entered.
- The Allow Users to Edit Ranges box's Permissions..., "Paste permissions information into a new workbook", Protect
  Sheet... and Apply; its words were taken from Excel's box, not measured over COM (it can't read them), and so were
  Unlock Range's and the New Range box's refusals (a title missing or taken, cells not of this sheet).
- F2 on a hidden unlocked cell starts from nothing, as the formula bar shows; what Excel's own edit starts from wasn't
  measurable over COM.
