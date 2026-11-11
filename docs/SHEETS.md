# Sheets

How sheets are added, renamed, moved, hidden, coloured and deleted (`src/ooxml/sheetops.js`), and what Excel was
seen to do. Each rule was measured by doing the same thing to the test workbooks in Excel over COM (opened read-only
and invisibly; copies saved outside the repository). Then the workbooks written here were opened in Excel:
- **Recalculation.** Excel's results match ours: 399 of 399 after a rename, 401 of 401 after a delete.
- **What Excel reads back.** Excel reads back the same sheets, hidden states, tab colours and names.

`npm test` (`test/sheets.test.js`) repeats each step and expects Excel's result.

## Excel's rules

| Rule | Seen |
|---|---|
| New sheets are named Sheet1, Sheet2..., from a count that starts at 1 when the workbook is opened. A name already taken is passed over, and a sheet renamed doesn't give its number back. | basic.xlsx: Sheet1; renamed X, the next is Sheet2. With Data renamed Sheet3: Sheet1, Sheet2, Sheet4. |
| A new sheet goes before the sheet shown (Insert, Shift+F11), or after it (the + button). It gets the next sheetId and an empty part: `dimension A1`, a sheet view, the default row height, `<sheetData/>`, the default margins. | Excel's saved copy of basic.xlsx with three sheets added. |
| A sheet name can't be blank, longer than 31 characters, contain `: \ / ? * [ ]`, or start or end with an apostrophe. | "You typed an invalid name for a sheet or chart. Make sure that: ..." |
| A space, an apostrophe inside, and any other character are accepted. | `' '`, `a'b`, `Sheet 9` |
| Names are compared without regard to case. | "That name is already taken. Try a different one." for `notes & MORE`, `data`. |
| "History" is reserved. | "History is a reserved name." |
| A rename rewrites every formula and defined name that names the sheet. | `=Data!B2:B9` becomes `='Sales 2024'!B2:B9`; `Items` becomes `='Sales 2024'!$A$2:$A$9`. |
| A sheet name is quoted in a formula unless it reads as a name. Names that look like a cell (up to XFD1048576), like R1C1, R or C, TRUE or FALSE, or that start with a digit are quoted. Letters of any script are names. | `'A1'!`, `'XFD1'!` but `XFE1!` and `A1048577!`; `'R1C1'!`, `'R'!`, `'TRUE'!`, `'2024'!`; `Sales.Q1!`, `Données!`, `_x!`; `'It''s'!`. |
| A delete turns each reference to the sheet into `#REF!` in the file (the whole reference: Excel shows `=#REF!A1`, but saves `#REF!`). Its results are #REF!. | `='My Data'!A1*2` saved as `#REF!*2`; `=SUM('My Data'!A1:B2,Data!B2)` as `SUM(#REF!,Data!B2)`. |
| A delete also takes away the names that belong to the sheet. A formula using one shows #NAME?. A workbook name that read the sheet becomes `#REF!`. | `=Mine` is #NAME?; `LocalOnMy` saved as `#REF!`. |
| Names that belong to a sheet follow it when sheets move (`localSheetId`). | Data moved first: `Here` saved with `localSheetId="0"`. |
| The last sheet shown can't be hidden or deleted. | Excel refuses over COM. Its dialog says "A workbook must contain at least one visible worksheet." |
| Tab colours resolve like cell colours. | Theme 5, tint 0.4 is #F1A983; theme 4, tint 0.4 is #44B3E1 (Tab.Color). |

## In the editor

The tab menu (right click, Shift+F10 or the Menu key on a tab) has Excel's entries: Insert, Delete, Rename, Move
Left, Move Right, Tab Color, Hide, Unhide.
- **Right click.** Right-clicking a tab shows that sheet first, as in Excel.
- **New sheets.** The + button adds a sheet after the one shown, and Shift+F11 adds one before it.
- **Renaming.**
  - Double-click a tab to rename it in place.
  - A refused name brings up Excel's message. After OK the name box stays open to correct it.
  - After a rename, focus goes back to the sheet.
- **Moving.** Drag a tab to move the sheet. Move Left and Move Right do the same from the keyboard.
- **Deleting.** If the sheet has something in it, the editor asks first. The API's `deleteSheet` doesn't ask.
- **Tab colours.** The palette has the theme's ten colours (written as `theme="n"`), Excel's ten standard colours
  and No Color. Excel's palette also has lighter and darker shades of the theme colours; those aren't offered yet.

## What else changes in the file

- **Renames reach other parts too:**
  - formulas in conditional formats, data validations and their extensions
  - hyperlink locations
  - chart series (`c:f`)
  - pivot cache sources
- **A deleted sheet takes its parts with it:**
  - the parts that only it led to go (drawings, comments, tables, printer settings)
  - a part that something else still reaches stays
  - calcChain.xml is dropped, and the workbook is worked out again
- **Each sheet operation keeps the workbook view valid:**
  - the active tab stays on its sheet; if that sheet is gone or hidden, it moves to the nearest sheet shown
  - the first tab is kept within range
  - a hidden sheet loses `tabSelected`
- **docProps/app.xml isn't updated.** It still lists the sheet names as they were. Excel rewrites it when it saves and doesn't use it when it opens a file.

## Undo

Each sheet operation (insert, rename, move, hide or show, tab colour, delete) is one step of the undo history: Ctrl+Z
puts the workbook back with its sheets as they were (`Workbook.snapshot` keeps the sheet list), shows the sheet that
was shown, and Ctrl+Y does it again. Excel can't undo deleting a sheet, and warns before it deletes one; here the
deletion can be undone, a difference kept on purpose.

## Not yet

- Copying a sheet (Move or Copy... with "Create a copy").
- Grouping sheets (selecting several tabs).
- References across several sheets (`Sheet1:Sheet3!A1`). The formula engine doesn't read them yet.
