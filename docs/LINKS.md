# Hyperlinks

A sheet's hyperlinks: read and written as Excel writes them (`src/ooxml/links.js`), and followed from the grid as
Excel follows them (`src/editor/linkfollow.js`). Measured over COM: `Hyperlinks.Add` with an address, a place in the
workbook, a ScreenTip and text to display, on empty cells and on cells with something in them, and the file Excel
saved (test/links.test.js). A file with hyperlinks written here opens in Excel with them: their addresses, places,
ScreenTips and the Hyperlink style.

## In the editor

- A click on a hyperlink's cell follows it, as Excel's does: an address in a new browser tab (http, https, mailto
  and ftp; a host page can take it with the `onLink(url)` option instead), a place in the workbook gone to as the Name
  Box goes ('Sheet2'!B5 shows Sheet2 with B5 selected). A press held longer than half a second, or let go on another
  cell, selects the cell instead ("Click and hold to select this cell").
- Over a hyperlink's cell, the hand, and Excel's ScreenTip: its own tip, or where it goes, then " - Click once to
  follow. Click and hold to select this cell."
- An address a page can't open (a file's path, relative or not, or any other scheme) is refused with Excel's words
  when it can't open a link: "Cannot open the specified file."
- The cell's menu on a hyperlink (a right click, Shift+F10 or the menu key): Edit Hyperlink..., Open Hyperlink and Remove Hyperlink (as Home > Clear > Remove
  Hyperlinks); on another cell, Link.
- Insert > Link (Ctrl+K, `src/editor/linkbox.js`): Excel's Insert Hyperlink box, Edit Hyperlink on a cell with one.
  Link to: Existing File or Web Page (Address), Place in This Document (Type the cell reference, and the sheet or
  defined name under Cell Reference and Defined Names), E-mail Address (E-mail address, Subject: mailto:...?subject=);
  Text to display (the cell's text; "<<Selection in Document>>", fixed, on a number or a formula, as in Excel) and
  ScreenTip... (Set Hyperlink ScreenTip). Edit Hyperlink shows the link and adds Remove Link. A place in another
  sheet is written as a formula names it ('Notes & more'!C5, Data!A1). An address starting www. gets http:// before
  it; anything else stays as typed. Each is one step of Undo.
- On a protected sheet, a hyperlink is inserted only with Insert hyperlinks allowed and its cell unlocked (measured:
  docs/PROTECT.md).
- Home > Clear's Clear Hyperlinks and Remove Hyperlinks (docs/FORMATS.md), and hyperlinks following a sheet rename and
  rows and columns inserted or deleted (docs/ROWCOL.md), were there before.

- A HYPERLINK formula's cell (`src/calc/hyperlinkcell.js`) is followed too: its link_location worked out in the
  cell's place (=HYPERLINK("https://example.com/"&B1,"Shown") opens https://example.com/ and B1's text;
  "#Sheet2!B5" goes to Sheet2's B5), with the same hand and ScreenTip; its cell menu has Open Hyperlink. Typed, the
  cell gets the Hyperlink style when HYPERLINK is at the formula's top level, as Excel gives it (measured below).

## Excel's rules

| Rule | Seen |
|---|---|
| An address: a relationship of the sheet, `TargetMode="External"`, its Target the address as given. | `<hyperlink ref="A1" r:id="rId1"/>` |
| A place in the workbook: `location`, and `display` with the text to display. | `<hyperlink ref="A3" location="'Sheet2'!B5" display="Go there"/>` |
| A ScreenTip: `tooltip`. | `<hyperlink ref="A2" r:id="rId2" tooltip="Tip here"/>` |
| An empty cell: the text to display becomes its value, else the address. | A1: https://example.com/page; A2: Shown text |
| A cell with something in it keeps it; `display` then holds the address. | A5 "kept text": `display="https://example.com/"`; A6 42 stays a number |
| The cell gets the built-in Hyperlink style: underlined, the theme's colour 10, nothing marked applied in the cell's format. | `<cellStyle name="Hyperlink" xfId="1" builtinId="8"/>`, `<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="1"/>` |
| An address without a scheme is kept as given (a file's path). | `Target="example.com"` |

A HYPERLINK formula, measured over COM (the formula set, the file saved): no `<hyperlink>` is written, its value is
`friendly_name` (else `link_location`), and the cell gets the Hyperlink style when HYPERLINK is at the top level of
the formula: =HYPERLINK(...) and =1+HYPERLINK(...,5) did, =IF(TRUE,HYPERLINK(...)) didn't. A place in the workbook is
written "#Sheet2!B5". Which HYPERLINK a click follows when there are several wasn't measured (clicks can't be sent
over COM); the first, depth first, is followed here.

Excel also writes an `xr:uid` on each hyperlink; it is left out here (Excel doesn't need it).

## Not yet

- The Insert Hyperlink box's words and its http:// before www. were taken from Excel's box, not measured over COM
  (it can't read them); so were its refusals ("You must type a valid cell reference or select a location in this
  document."). Its Create New Document, Browse for File and Bookmark... aren't here.
- Half a second for a held click is a guess; Excel's own time wasn't measured.

# Links to other workbooks

A formula can read another workbook's cells, or one of its defined names; the workbook is then a link of this one
(`src/ooxml/extlinks.js`), as Excel keeps it. Measured over COM: Paste Link between two workbooks of made-up numbers,
names added that refer to the other one, a formula using its own defined name, the file saved with the source open and
again after it closed (`tools/excel-compare/excel-extlinks.ps1`; what Excel wrote reduced to
`test/extlinks-excel.json` by `extlinks-extract.mjs`, Excel's own file staying outside the repo since its paths are
the computer's), and Break Link on one of two links (`excel-breaklink.ps1`). `test/extlinks.test.js` builds the same
workbooks here: Excel's 14 formulas and values, its 4 names and its link part (sheets, names, values kept, file)
the same.

## In the editor

- Copy cells in one editor and use Paste Link (Home > Paste ▾, Paste Special) in another on the same page: formulas
  that read them, `=[1]Sheet1!A1` in the file, with the values copied kept in the link. Two editors on a page are two
  workbooks open at once in Excel, the last copy made in either shared between them.
- The formula bar and the cell being edited show the workbook by its file name: `=[oxesrc.xlsx]Sheet1!A1`,
  `='[oxesrc.xlsx]Data Two'!$A$1`, `=[oxesrc.xlsx]!SrcRate` for its defined name. Typed so, a formula is written with
  the link's number, the link made when there is none: with that workbook's sheets and values when it is open in an
  editor here, else as a link not read yet, its references #REF! until updated.
- The values kept are what formulas show (worked out from them, `=SUM(ExtBlock)` too) until **Data > Edit Links >
  Update Values**, which reads them again from the workbook open here by that name, or from a file chosen.
- **Data > Edit Links**: Excel's box, the links with their Source, Type (Worksheet), Update (A, automatic) and Status
  (OK when the workbook is open in an editor here; Unknown otherwise, as Excel says before Check Status), the chosen
  one's location; **Update Values**, **Change Source...** (another file: its name, and its values), **Break Link**
  (after Excel's question) and Close. Each is one step of Undo.
- Defined names may refer to another workbook: the Name Manager lists them with its file's name, and New Name and Edit
  Name take them typed so (docs/NAMES.md).

## Excel's rules (measured)

| Rule | Seen |
|---|---|
| A link is a part `xl/externalLinks/externalLinkN.xml` (content type `…spreadsheetml.externalLink+xml`), the workbook's `externalLink` relationship to it, listed in `<externalReferences>` after `<sheets>`; a formula writes its place in that list. | `<externalReferences><externalReference r:id="rId2"/></externalReferences>`, `[1]Sheet1!A1`. |
| The part: an `externalBook` whose relationship (`externalLinkPath`, `TargetMode="External"`) is the file, relative to this one; all of its sheets' names; its workbook-wide names, sheets quoted; and, sheet by sheet, the values of the cells read there. | `Target="oxesrc.xlsx"`; `<sheetName val="Data Two"/>`; `<definedName name="SrcRate" refersTo="='Sheet1'!$B$1"/>`; `<cell r="A2" t="str"><v>text</v></cell>`, `t="b"`, `t="e"`. |
| A blank cell read is kept as nothing; the formula shows 0. | `=[1]Sheet1!A3`: 0, no A3 in the part. |
| Paste Link: as within a workbook (docs/CLIPBOARD.md), each reference with the link and the sheet. | A1:C2 at D4: `[1]Sheet1!A1` ... `[1]Sheet1!C2`; one cell: `[1]Sheet1!$B$1`; `'[1]Data Two'!$A$1`. |
| A name referring to another workbook writes its link the same way. | `ExtBlock`: `[1]Sheet1!$A$1:$B$2`; a sheet's own `LocalExt`: `'[1]Data Two'!$A$1`. |
| Its own defined name: `[n]!Name`. | `=oxesrc.xlsx!SrcRate` typed: `[1]!SrcRate`, 2.5. |
| Break Link: each cell whose formula refers to the link becomes the value it shows; names keep their references, and the link stays with them; the other links keep their numbers. | `=[1]Sheet1!$A$1*2` is 10; `=SUM([1]…,[2]…)` is 14; `Na` stays `[1]Sheet1!$A$1`; `[2]Sheet1!$A$1` stays. |

Excel also writes the file's absolute path beside the relative one (`xxl21:alternateUrls`, `file:///C:\…`); it is
left out here (a page doesn't know where its files are), and a Change Source drops it. Excel's third sheet name in
the measured part (`oxesrc`, with `refreshError`) came from its reading `=oxesrc.xlsx!SrcRate` as a sheet before the
name: not made here.

## Not yet

- Excel's own typing of another workbook's name, `=Book.xlsx!Name`, and of a closed one's path, `='C:\Data\[Book.xlsx]Sheet1'!A1`
  (read, the path set aside): typed here as `[Book.xlsx]!Name` and `[Book.xlsx]Sheet1!A1`.
- Check Status, Open Source and Startup Prompt; links to other kinds of files (DDE, OLE); updating a link from a file
  on its own (a page can't open a file by its path: one is chosen).
- A link typed and then taken back with Undo stays in the file, unused.
- Paste Link's number format from another workbook (the measured copy was General).
