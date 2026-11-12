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
- The cell's menu on a hyperlink: Edit Hyperlink..., Open Hyperlink and Remove Hyperlink (as Home > Clear > Remove
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

Excel also writes an `xr:uid` on each hyperlink; it is left out here (Excel doesn't need it).

## Not yet

- The Insert Hyperlink box's words and its http:// before www. were taken from Excel's box, not measured over COM
  (it can't read them); so were its refusals ("You must type a valid cell reference or select a location in this
  document."). Its Create New Document, Browse for File and Bookmark... aren't here.
- The keyboard: the cell's menu opens with a right-click only (not Shift+F10 or the menu key), so Open Hyperlink
  needs the mouse.
- Half a second for a held click is a guess; Excel's own time wasn't measured.
- The HYPERLINK function's cells aren't followed.
