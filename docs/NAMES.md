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

The messages ("The syntax of this name isn't correct...", "The name entered already exists. Enter a unique name.")
are Excel's dialogs', which COM doesn't reach: not measured.

## Not yet

- The Name Manager's Value column and its filter; a name's scope changed (Excel doesn't either: it is fixed once made).
- Create from Selection, Use in Formula, and Apply Names.
- Names that refer to another workbook.
- A LAMBDA kept in a name and called by it.
