# Defined names

Names made, changed and deleted as Excel's Formulas > Defined Names does it (`src/ooxml/names.js`, the boxes in
`src/editor/namesbox.js`), and a selection named from the Name Box.

## In the editor

- **Formulas > Name Manager** (Ctrl+F3): the names with what each refers to, its scope and comment (Excel's own hidden
  and built-in names, such as a filter's or a print area's, left out, as Excel leaves them out). New..., Edit...,
  Delete (after Excel's question) and Close.
- **Formulas > Define Name**, or New... in the Name Manager: Excel's New Name box, the selection as what the name
  refers to, the workbook as its scope (or a sheet), a comment.
- **The Name Box:** a new name typed there names the selection, workbook-wide, as Excel's does. When the selection is
  exactly a name's range it shows the name instead of the cell (below), and its drop-down arrow (or Alt+Down) lists
  the names to go to.
- **Formulas > Use in Formula** and **Paste Names** (F3): the names the sheet sees (below), the one chosen typed at
  the caret of the cell being edited, or the active cell opened with "=Name". Paste Names' box has Paste List, which
  writes them from the active cell down as Excel's does (below); not while a cell is being edited, as Excel's is then
  dimmed. With no names, F3 opens nothing.
- **Formulas > Create from Selection** (Ctrl+Shift+F3): Excel's Create Names from Selection box (Top row, Left column,
  Bottom row, Right column) names the cells beside each label, as Excel's does (below). A name there already is asked
  about first (Yes replaces it, No keeps it, Cancel makes none); all the names made are one step of the history.
- A name Excel wouldn't take, or one there already is in that scope, gets Excel's words in the box, which stays open.
- Formulas that use a name are worked out again when it is defined, changed or deleted (`#NAME?` until it exists, and
  again after it is deleted). Renamed, the formulas that use it follow, except on a sheet with a name of its own the
  same, whose formulas read that one.
- Each change is one step that Ctrl+Z takes back. The agent and the MCP server define, change and delete names too
  (docs/AGENT.md).

## Excel's rules (measured)

`tools/excel-compare/excel-names.ps1` added 39 candidate names to a new workbook with `Names.Add` in a hidden Excel and
saved names added out of order (`test/names-excel.json`).

| Rule | Seen |
|---|---|
| A letter or _ first, then letters, digits, _, ., ? and \ | `Rate`, `_x`, `Tax.Rate`, `Rate?`, `AB12C`, `_`, `a\b`, `x\`, `x\y\`, `é`, `Ünits`: taken. `1abc`, `my rate`, `Q1-Q2`, `Total!`, `Sales%`, `x$`: refused. |
| \ alone, but not \ followed by more | `\` taken, `\x` refused. |
| Not a cell's address | `A1`, `a1`, `XFD1` refused; `XFE1` (past the last column) taken. |
| Not anything R1C1 reads | `R`, `C`, `r`, `c`, `RC`, `R2`, `C3`, `RC1`, `R1C`, `R1C1` refused. |
| Not TRUE or FALSE, nor Excel's own names | `TRUE`, `Print_Area` refused. |
| At most 255 characters | 255 taken, 256 refused. |
| One name per scope | The same name is allowed once for the workbook and once on each sheet. |
| The file's order | By name, capitals aside, a built-in name (`_xlnm.Print_Area`) by its own name; a sheet's before the workbook's of the same name; `hidden="1"`, `localSheetId`, the text without "=". 9 names saved here: the same `definedNames` as Excel's. |

`test/names.test.js` checks every verdict (all 39 the same), the saved order, and Define, Edit and Delete with the
formulas that use them.

## The Name Box (measured, as far as COM reaches)

COM can't read the Name Box: in a hidden Excel its Edit control stays empty (WM_GETTEXT read "" with a named range
selected, and its list held no items). What Excel reports for the selection itself can be read: `Selection.Name`, the
name whose range the selection is. `tools/excel-compare/excel-namebox.ps1` selected ranges in a hidden Excel 16 under
names of each kind, 35 cases (`test/namebox-excel.json`); `selectionName` in `src/ooxml/goto.js` gives the same name
for all 31 without a table (`test/namebox.test.js`), and the Name Box shows it.

| Rule | Seen |
|---|---|
| Exactly the name's range, wherever the active cell is | B2:C4 under =Sheet1!$B$2:$C$4: Rate, with C3 active too. B2:C3, B2:C5 and B2 alone: none. One cell (E5), whole columns ($B:$B) and rows ($3:$5) the same. |
| A name of ranges: the same ranges in the same order | A1 then C3:D4 under =Sheet1!$A$1,Sheet1!$C$3:$D$4: Multi. C3:D4 then A1, or A1 alone: none. |
| Absolute references only | A relative =Sheet1!B2 with B2 selected, and =OFFSET(Sheet1!$A$1,0,0,2,2) with A1:B2: none. |
| Of more than one, the first by name, capitals aside | Zed and Alpha (Zed added first): Alpha. beta and Alpha2: Alpha2. A workbook's Aw and the sheet's own Zown: Aw. |
| The same name: the sheet's own first | Dup of the workbook and Dup of Sheet1: Sheet1's. Mid of Sheet1 and of Sheet2: Sheet1's. |
| Hidden names and print areas count | A hidden Hid alone: Hid; hidden Ahid and Seen: Ahid. A print area over B2:C4, with Rate too: Print_Area. |
| Another sheet's own name counts | Sheet2's own Theirs over Sheet1!B2:C4, selected on Sheet1: Sheet2!Theirs (Sheet2's Aa before Sheet1's Bb too). |
| A table is no name | Its data rows, the whole of it, a column of it: none. A name over its data rows: that name. |
| Going to a name | Application.Goto from Sheet2: Rate selected B2:C4 on Sheet1, B2 active; Multi (=Sheet1!$C$3:$D$4,Sheet1!$A$1) C3:D4 and A1, A1 (the last range's first cell) active; Selection.Name then the name. Going to a name of ranges works here now too. |

COM gives a sheet's own name as "Sheet1!Own"; the Name Box shows the shown sheet's own without its sheet ("Own"), and
another sheet's with it ("Sheet2!Theirs"), as COM gives it: not measured, since the Name Box can't be read.

The drop-down: Excel's Name Box has an arrow that lists the names to go to (documented, not reachable over COM). Its
list here is Go To's (`goToNames`): the names standing for ranges the sheet sees, not Excel's own hidden ones, and the
tables, sorted. Which names Excel's list holds (hidden ones, other sheets' own) isn't measured; nor whether it shows a
table's name for a table's range, which `Selection.Name` doesn't report, so a table's range shows its cells here.

## Paste List (measured)

`Range.ListNames` in a hidden Excel, on two sheets of a workbook with names of the workbook's and of each sheet's own,
one hidden, a print area and a constant (`test/createnames.test.js`, the same lists here):

- Listed: the workbook's names and the sheet's own, one of its own hiding the workbook's of the same name; not hidden
  names, nor other sheets' own; the sheet's print area as `Print_Area`. Sorted by name, capitals aside.
- Written from the cell down: the name, and beside it what it refers to, as text (`=0.07`, `='My Data'!$C$4`, a
  shared string with no quote prefix), over what was there.

Use in Formula and Paste Names list the same names here (their lists aren't reachable over COM: not measured).

## Create from Selection (measured)

Three probes in a hidden Excel gave 86 labels to `Range.CreateNames` and saved the result (`src/ooxml/createnames.js`;
`test/createnames.test.js` has 81 text labels and the name Excel made of each, all 81 the same here).

| Rule | Seen |
|---|---|
| Text labels only | A number (2020, 3.5, -4, =1+1), TRUE typed as a value and an empty cell named nothing; the text '2020 named `_2020`. |
| Characters a name can't have: dropped at the ends, _ between | `  padded  ` padded, `Total!` Total, `Price ($)` Price, `(x)` x, `Unit Price` Unit_Price, `A B  C` A_B__C, `a!-!b` a___b, `$A$1` A_1, a line break _. |
| _ before a digit, . or ? first | `1st` _1st, `.dot` _.dot, `?q` _?q, ` 1x` _1x. |
| _ before R or C and a digit, or RC and a digit | `R1C1` _R1C1, `R1` _R1, `C3` _C3, `R2D2` _R2D2, `rc2x` _rc2x; `RCx`, `Rx` and `RCC` as they are. |
| _ after R, C, RC, a cell's address, TRUE and FALSE | `R` R_, `rc` rc_, `A1` A1_, `XFD1` XFD1_, `CR1` CR1_, the text TRUE TRUE_; `XFE1`, `A0` and `XFD1048577` as they are. |
| Still no name: none | `\x`, `\\`, `_` alone, `-`, `!!`. |
| At most 255 characters | 300 Ls named with 255. |
| The cells | Top row: those under each label; Left column: right of it; Bottom row: above; Right column: left of it; where two meet, the corner's label names all the cells inside (a grid's top-left `Grid` named C3:D5). Workbook-wide, absolute, with the sheet. |
| A name there already | Replaced, its capitals too (`RC_` then `rc_` left `rc_`). |
| No cells left to name | A single cell with Top row: refused. |

Excel's file opened the names made here the same way (8 of 8, each summing the same cells).

Left out here: a label `Print_Area` or `Print_Titles`, of which Excel makes a workbook-wide built-in name
(`_xlnm.Print_Area` with no sheet). Not measured, because COM doesn't reach the box: which sides it starts with
checked (here, Top row and Left column when they hold a label past the first cell, as Excel's Help describes), its
question before replacing a name ("Do you want to replace the existing definition of 'Qty'?"), and its words for a
selection with nothing to name.

The messages ("The syntax of this name isn't correct...", "The name entered already exists. Enter a unique name.")
are Excel's dialogs', which COM doesn't reach: not measured.

## Apply Names (measured)

Formulas > Define Name ▾ > Apply Names... puts names in place of the references to them (`applynames.js`). Measured
with Range.ApplyNames in a hidden Excel, each formula on its own: 29 formulas with Ignore Relative/Absolute on and 10
with it off, 16 on a grid of row and column names under three settings of the options; all 87 the same here
(`test/applynames.test.js`).

| Excel's rule | Excel's example |
|---|---|
| One cell selected: the whole sheet's formulas; more: the selection's. | Applied from H3, every formula on the sheet. |
| Names that refer to one area of this sheet, all $, are applied; a relative name, a constant, a formula name are not. | Rel (=Sheet1!B3): =B3 stays. Tax (=0.07): 0.07 stays. |
| Only references without a sheet's name. | =Sheet1!B2*2 and =Data!$A$1+1 stay, with names for both. |
| Ignore Relative/Absolute on: any of B2, $B$2, B$2, $C2; off: only the name's own $B$2. | =B$2*$C2: Qty*Price; off: stays. |
| A range that is a name's becomes it. In one that isn't, an end with a relative column and a cell's name becomes the name; a $ column's stays (on only). | B2:B5 Amounts; B2:B4 Qty:B4; A1:B2 A1:Qty; $B$2:$B$4 stays. |
| Use row and column names: a cell where a named column and a named row cross becomes the two, row first (Column row order: column first). From the cell's own row, the column's name alone; from its own column, the row's (the two Omit options). | =C3 in H9: South Jan; in C6: South; =C2 in H2: Jan. |
| That comes before the cell's own name; without a row and a column, the cell's own name. | =B2 in H2 with Qty and Amounts: Amounts; in H8: Qty. |
| Text in quotes and names already there stay. | =B2&"B2": Qty&"B2". |

What Excel does to a formula's value comes with it: =B2*C2 in H2 becoming Amounts*Price, it spills and shows
#SPILL! with something below, in Excel as here (Excel saves that one as an array formula; here it stays a plain one,
which Excel reads back as one cell, 20). In a formula entered as a dynamic array, Excel's row and column names left
B2 as Qty (implicit intersection doesn't apply there); here every formula is read as an opened file's are.

Not measured: which names Excel's box starts with chosen (all here); its messages when nothing can be applied (the
editor's own words); a one-cell range (C2:C2, which Excel made @Price, saved as SINGLE(Price)) is left as it is here.

## Names that refer to another workbook (measured)

A name's formula may read another workbook, through the link to it (docs/LINKS.md, `src/ooxml/extlinks.js`). Measured
over COM (`excel-extlinks.ps1`: Names.Add with references to another open workbook, the file saved), Excel writes the
link's number, `[1]Sheet1!$A$1:$B$2`, `'[1]Data Two'!$A$1` for a sheet's own name; `test/extlinks.test.js` makes the
same four names here (the workbook's ExtRate, ExtBlock and ExtConst, Sheet1's LocalExt) and Excel's `definedNames`
exactly.

- The Name Manager lists them with the workbook's file name, as typed: `=[oxesrc.xlsx]Sheet1!$A$1`. New Name and Edit
  Name take them typed so, and write the link's number; a workbook not linked yet becomes a link (read from an editor
  on the page that has it open, else not read until Update Values).
- Formulas using them are worked out from the values the link keeps: `=SUM(ExtBlock)` is 17.5, as Excel's.
- Kept as they are: a workbook opened and saved with such names, not edited, writes its `workbook.xml` and link parts
  byte for byte; one edited elsewhere (a cell typed) leaves them too (`test/extlinks.test.js`). Break Link leaves them,
  as Excel's does.
- Not here: a name's scope in the other workbook (`sheetId` in the link part) is read but not offered, and Excel's own
  `=Book.xlsx!Name` typing (here `[Book.xlsx]!Name`).

## Not yet

- The Name Manager's Value column and its filter; a name's scope changed (Excel doesn't either: it is fixed once made).
- A LAMBDA kept in a name and called by it.
