# Defined names

Names made, changed and deleted as Excel's Formulas > Defined Names does it (`src/ooxml/names.js`, the boxes in
`src/editor/namesbox.js`), and a selection named from the Name Box.

## In the editor

- **Formulas > Name Manager** (Ctrl+F3): the names with what each refers to, its scope and comment (Excel's own hidden
  and built-in names, such as a filter's or a print area's, left out, as Excel leaves them out). New..., Edit...,
  Delete (after Excel's question) and Close.
- **Formulas > Define Name**, or New... in the Name Manager: Excel's New Name box, the selection as what the name
  refers to, the workbook as its scope (or a sheet), a comment.
- **The Name Box:** a new name typed there names the selection, workbook-wide, as Excel's does.
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

## Not yet

- The Name Manager's Value column and its filter; a name's scope changed (Excel doesn't either: it is fixed once made).
- Names that refer to another workbook.
- A LAMBDA kept in a name and called by it.
