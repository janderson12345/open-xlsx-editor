# Cell formats

How the Home tab's commands change cells (`src/ooxml/formatops.js`) and how the change is written to the styles part
(`src/ooxml/stylesedit.js`). Each rule was measured by running Excel's own ribbon commands over COM
(`CommandBars.ExecuteMso` on a new workbook in a hidden instance; no keys sent, nothing saved but copies outside the
repository) and reading back what Excel showed and what it saved. The workbooks written here were then opened in Excel,
which reads back every font, fill, alignment, border, number format and cell style as written. `npm test`
(`test/formats.test.js`) repeats each step and expects Excel's result.

## How a change is stored

- A cell's format is an entry of `cellXfs`; the change makes a copy of the cell's entry with the change in it, and of
  the font, fill or border it points to. An equal entry already there is used instead, so formatting many cells with
  the same look adds one entry; a new one goes at the end and no index a cell uses ever moves.
- A formatted empty cell is kept as a `c` element with only its `s`, as Excel keeps it.
- New lines are automatic black (`indexed="64"`), as Excel writes the menu's borders.
- A workbook without a styles part gets one the first time it is formatted (Calibri 11, the two fills Excel always
  keeps, the Normal style).
- Rows without a fixed height grow when their font does: Excel fits them when it opens the file (20pt text: 26.25pt).

## Excel's rules

| Rule | Seen |
|---|---|
| Bold, Italic, Underline, Strikethrough and Wrap Text follow the active cell: on for every selected cell unless the active cell has it, then off for all. | A1 plain, B1 bold: both bold. A1 bold: both plain, whichever else is bold. |
| Underline is single. | `Font.Underline` 2. |
| Grow and Shrink Font set every selected cell to one step from the active cell's size, through 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72, and no further. A size between steps goes to the next one. | From 11: 12, 14 ... 72, 72. Down: 10, 9, 8, 8. 13 grows to 14. 11 and 20 both become 12. |
| Align Left, Center, Right, Top, Middle and Bottom set the alignment; pressing one again leaves it. | Center twice: still centred (through ExecuteMso; to be checked by hand on the ribbon). |
| Increase Indent adds a level and makes the cell left-aligned, unless it is right-aligned. Decrease Indent takes one away and leaves the alignment. | General or centred: left, 1. Right: right, 1. Up twice then down: 1, left. |
| A border goes on the cells along the selection's edge (Bottom: its bottom row). Inside lines change only with All Borders and No Border. | B2:C3 Outside: B2 left and top, C3 bottom and right. |
| A border drawn on the selection's edge takes away the neighbour's facing side, unless it is the same line. No Border takes them away too. | Top on B2: B1's medium bottom goes, a thin one stays. No Border on B2: B1's bottom and C2's left go. |
| Thick Outside Borders and Thick Bottom Border are stored as `medium`; Bottom Double Border as `double`. | Excel's saved copy. |
| Percent Style, Comma Style and the $ button (Accounting Number Format) apply Excel's built-in cell styles Percent (`builtinId` 5, format 9), Comma (3, format 43) and Currency (4, format 44). The style is added to `cellStyleXfs` and `cellStyles` when the workbook hasn't got it, and the codes of formats 43 and 44 are written out. | `Style.Name` Percent, Comma, Currency; Excel's saved `styles.xml`. |
| Increase and Decrease Decimal change the active cell's format and give it to every selected cell. From General they start from the decimals the cell shows; each section of a format gets one place more or fewer; a date, a time, text or a General cell that isn't a number doesn't change. | 1.5: `0.00`, `0.000`. 2.123456789 (shown 2.123457) down: `0.00000`. `0.00%` down: `0.0%`, `0%`. `#,##0.00 "kg";[Red]-#,##0.00` up: both sections. `m/d/yyyy` unchanged. |
| Font Color and Fill Color start as red (`FFFF0000`) and yellow (`FFFFFF00`, solid, background `indexed="64"`). | `Font.Color` 255, `Interior.Color` 65535. |

## Merging (`src/ooxml/mergeops.js`)

| Rule | Seen |
|---|---|
| Merging keeps one value, the first in reading order, and puts it in the top-left cell; the others are emptied. Excel asks first: "Merging cells only keeps the upper-left value and discards other values." | A1, B1, A2 filled in, A1:B2 merged: `a1` in A1, the rest empty. A1 and B1 empty: A2's `a2` moves to A1. |
| Every cell of the area gets the format of the cell whose value is kept (Excel writes them all with it). | A1 filled yellow, B2 bold: all four filled, none bold. A2's value kept: all four plain. |
| Merge & Center also centres; Merge Cells and Merge Across leave the alignment. Merge Across merges each row of the selection on its own. | A1:B2 across: A1:B1 and A2:B2, A2 keeping `a2`. |
| A selection that overlaps a merged range is merged with it into one. | A1:B2 merged, then B2:C3: A1:C3. |
| Merge & Center on a selection with a merged range in it unmerges, and sets the alignment back to General. Unmerge Cells leaves the alignment. | A1 after Merge & Center, pressed again: not merged, General. Unmerged: still centred. |

## Clear (`src/ooxml/clearops.js`)

Measured over COM: `Range.Clear`, `ClearContents` and `ClearFormats` on a small table with a formula, formats, a
border and two merged ranges, six cases (`tools/excel-compare/excel-steps.ps1` on `clear-cases.json`, from
`clear-cases.mjs`, giving `test/clear-excel.json`). `test/clear.test.js` expects Excel's cells and merged ranges
in each; cleared here and opened in Excel, the three that clear read back as Excel made them.

| Rule | Seen |
|---|---|
| Clear All empties the cells and gives them the Normal format; merged ranges inside are unmerged. | B2:E8: values, bold, fill and border gone, D8:E8 unmerged. |
| Clear Contents (and Delete) empties them; formats and merged ranges stay. | B2 keeps bold and 0.00. |
| Clear Formats gives them the Normal format and unmerges; the values stay. | B3 `=B2*2` kept, plain. |
| A selection over part of a merged range is refused, whichever Clear: "We can't do that to a merged cell." | B9:C10 with A9:B9 merged. |

A click on a merged cell selects the whole merged range in Excel; here a selection of one of its cells clears the
whole range in the same way.

### Comments and Notes, Hyperlinks (`src/ooxml/clearparts.js`)

Measured over COM: `Range.ClearComments`, `ClearHyperlinks`, `Hyperlinks.Delete` (Remove Hyperlinks), `Clear` and
`ClearContents` on a sheet with a hyperlink to an address in A1 (bold), a note in B1, a threaded comment in C1 and one
hyperlink over D1:E2, and the files Excel then saved. `test/clearparts.test.js` expects the same; cleared here and
opened in Excel, its hyperlinks, notes and comments are those.

| Rule | Seen |
|---|---|
| Clear Comments and Notes takes the notes and threaded comments (with their replies) of the cells, nothing else. | B1 alone: its note goes, C1's thread stays. A1:E2: the comments, drawing and threaded comments parts go. |
| Clear Hyperlinks takes the hyperlinks meeting the cells, each whole, and leaves the formats (the Hyperlink style stays). | D1 alone: D1:E2's hyperlink goes; A1's stays. |
| Remove Hyperlinks does that and takes the formats of the hyperlinks' cells away, theirs only. | A1 loses its bold, E2 its style; a bold F1 beside them stays bold. |
| Clear All takes them all with the cells; Clear Contents takes the hyperlinks (their formats stay) and leaves the notes; Clear Formats leaves both. | |

Excel drops a comments part it leaves empty; here it stays, empty, as after Delete Note.

## Format Cells (`src/editor/formatbox.js`, `formatops.js` formatCells)

Ctrl+1 opens Excel's box on the active cell's format; only what is changed in it is given to the selection, each other
setting staying as each cell has it. How Excel saves each setting was measured over COM
(`tools/excel-compare/excel-formatcells.ps1`: one cell per setting on a new workbook, saved, its `xf`, font, fill and
border read back; `test/formatcells-excel.json`). `test/formatcells.test.js` makes all 43 here and expects Excel's text.

| Rule | Seen |
|---|---|
| Rotation is `textRotation`: 1 to 90 degrees up as themselves, -1 to -90 as 91 to 180, stacked letters 255. | 45: 45; -90: 180; vertical text: 255. |
| An indented or Fill-aligned cell isn't rotated: Excel drops the rotation, whichever was set first. Centred text keeps it. | Indent 2 then 30 degrees, or 30 then indent: `horizontal="left" indent="2"`, no rotation. |
| Wrap text and shrink to fit can both be on in the file (the box offers one or the other). | `wrapText="1" shrinkToFit="1"`. |
| The alignment's attributes go in Excel's order: horizontal, vertical, textRotation, wrapText, indent, shrinkToFit, readingOrder. | `horizontal="left" indent="2" shrinkToFit="1"`. |
| A gradient is `gradientFill` with its stops; degree 0 is left out; a path gradient has `type="path"` and its rectangle. | `<gradientFill degree="90"><stop position="0">...` |
| A pattern's own colour is `fgColor`, the cell's colour `bgColor`; a shade without colours has none. | Reverse diagonal stripe, red on yellow: `fgColor` FF0000, `bgColor` FFFF00; 50% gray: `mediumGray` alone. |
| Diagonals are `diagonalUp` and `diagonalDown` on the border (in that order) and one `diagonal` line for both. | `<border diagonalUp="1" diagonalDown="1">...<diagonal style="thin">` |
| Superscript and subscript are the font's `vertAlign`; the accounting underlines `singleAccounting`, `doubleAccounting`. | |
| Protection: `locked="0"` when unlocked, `hidden="1"` when hidden, nothing when locked and shown. | |
| Over a range, the outline goes on its edge and the inside lines on both facing sides of each pair of cells. | D2:F4: E3 dashed all round, D2 thin left and top. |

Rows grow for turned text as Excel's AutoFit grows them (Calibri 11, measured with `Rows.AutoFit`): stacked letters a
line each and 1.5pt ('ab' 31.5, 'Stack' 76.5); text at an angle by its width up the angle and a line across it, within
2pt of Excel's 36 to 81pt for 15 to 90 degrees.

Not settled: a range's border over COM also took a matching thin line off the neighbour beyond its edge (G3's left
beside F3's new right), where the ribbon's Borders measured earlier keeps a matching line; this follows the ribbon.
The degrees the box's own Fill Effects variants give, and the font sizes Excel steps through when shrinking, are not
yet measured.

## A row's or a column's own format

Measured over COM (`Columns(2).Font.Bold`, `Range("D:E").Interior.Color`, `Rows(5).Font.Italic`, `Rows(7).Font.Size`,
and the sheet Excel saved; test/linestyles.test.js): a formatted column is `<col min max style="X">` (with the
default width), a formatted row `<row s="X" customFormat="1">`; the cells already there get the change, and where a
formatted row crosses a formatted column Excel makes an empty cell with both (B5: bold and italic). A row's font
made larger grows it (ht="18.75" for 14pt).

Here (`src/ooxml/sheet.js` lineStyle): an empty cell has its row's own format, else its column's; its fill and
borders are drawn from it, a cell typed there takes it (`src/ooxml/sheetedit.js`), and Format Cells > Protection on
a column counts for a protected sheet.

Formatting whole rows or columns here (any of the Home tab's formats, Format Cells, the borders) does as Excel does
(`src/ooxml/formatops.js` probeOf, `src/editor/commands.js` lineFormats): the columns' or rows' own formats are
written, the cells already there changed, and a cell made where a formatted row crosses a formatted column; empty
cells aren't made. The new format of each line is worked out on a row (column) past the used range that stands for
the rest of it, so the toggles and borders follow it as they follow a cell (an outline's left and right on the
first and last columns; top on row 1). One step of Undo. Excel's own run of Rows(5) italic, Columns(2) bold and
D:E filled was made here the same way, and Excel read it with the same formats, cell for cell (B100 bold, D5 italic
and filled, E900 filled).

The whole sheet (Ctrl+A), measured over COM (`Cells.Font.Bold` after a column widened, a row heightened, a row
italic and a column filled): every column gets its own format, its `<col>` elements split where a column already had
a width or a format of its own (`<col min="1" max="2"/><col min="3" max="3" customWidth="1"/><col min="4" max="4"/>
<col min="5" max="5"/><col min="6" max="16384"/>`, each with its style); the rows with a format of their own and the
cells there are changed; no cell is made, a row with only a height is left. Done here the same way, Excel read it
the same, cell for cell (XFD1048576 bold among them).

A row with a format of its own grows to its font as Excel's does (Rows(7).Font.Size = 14: 18.75pt), cells or not
(`src/editor/rowheights.js`); Excel fits a row without a height of its own when it opens the file, so none is
written.

Not yet: an outline's bottom on row 1,048,576 isn't drawn.

## Rich text

A cell's text in runs of several fonts (`src/ooxml/strings.js` stringItemRuns, drawn by `src/editor/cells.js`): each
run in its own font over the cell's, bold, italic, underline, strikethrough, size, colour, font and super- or
subscript. Measured over COM (`Characters(start, length).Font` set, the sharedStrings Excel saved,
test/richtext.test.js): each run an `<r>` with an `<rPr>` giving every property (the font's name as `rFont`), the
first run without one when it is the cell's own font; the cell's format takes the first run's font (a struck first
word made the cell's font struck). Shown as Excel shows them; Show Formulas, and a number format that changes the
text, show it plain.

A run's larger font grows its row as Excel's does (a 16pt run beside Aptos Narrow 11: 21pt, measured).

Part of a cell's text is formatted while it is edited, as in Excel: select some of the text in the cell editor and press
Ctrl+B, Ctrl+I, Ctrl+U or Ctrl+5 (or the Font group's Bold, Italic, Underline, Strikethrough); that part takes it when
the edit is entered (`src/editor/richmarks.js`, one font per character). A cell of rich text keeps all of its runs'
fonts while edited (bold, italic, underline, strikethrough, size, colour, font, super- and subscript). They are written
as Excel writes rich text (`strings.js` richRunsXml, `test/richtext.test.js`): a run in the cell's own font bare,
another with an rPr of what it changes and then the cell's size, colour, rFont, family and scheme; in the cell's inline
string. A run read from a file keeps its own colour (rgb, theme and tint, or indexed) and, in a font of its own, its
family. (Excel adds a family number for a run in another font, `family val="3"` for Courier New, which the editor can't
know for a font chosen in it and leaves out.)

While the cell is edited the text is drawn as it will be (`src/editor/richview.js`), as Excel's in-cell editor shows
it: each run in its font, updated as marks are toggled and as text is typed. The text box stays the one typed in (its
caret, selection, IME, undo and keys are the browser's own); while part of the text is formatted its own text is
hidden and the runs are drawn under it, with the caret and the selection placed by the drawn text, and a click or drag
placed by the drawn characters. A formula, a number, and text with nothing formatted are the text box's own plain
text, as before. The formula bar shows the text plain, as Excel's does.

Typed text takes a font as Excel's own insert gives it: `Characters(start, length).Insert(text)` over COM on a cell of
runs, each character's font read back (`tools/excel-compare/excel-richtype.ps1`, `test/richtype-excel.json`, checked by
`test/richmarks.test.js`):

| Where the text goes | Its font (measured) |
| --- | --- |
| Inside a run | the run's |
| Just after a bold run, or a red 16pt one | bold; red 16pt (the font before the caret) |
| Just before a bold run, or a red 16pt one | plain (the font before the caret) |
| At the very start, the first run bold | bold (the first character's) |
| In place of a selection | the selection's first character's (plain then bold: plain; bold then plain: bold) |
| At the very end, the last run bold | plain (the cell's own) in the object model |

The last row isn't followed: typing at the end goes on in the font before the caret like everywhere else, which isn't
measured (keystrokes can't be sent to Excel over COM, and the object model's insert past the last character may not be
what typing does). Undo inside the editor (Ctrl+Z, the text box's own) brings the fonts back with the text.

The Font group's font and size lists, Increase and Decrease Font Size and Font Color work on the selected part of the
text too, as they do in Excel's in-cell editor (`src/editor/acts.js` partOfText, `richmarks.js` setFont): that part
takes the font, the size (Grow and Shrink step from its first character's size, the cell's where it has none) or the
colour, the rest keeps its own; Automatic is written as theme colour 1, as a cell's own Automatic is. Focus going to
the font or size list keeps the edit open (`richtrack.js` keepsEdit) and comes back to it once the choice is made;
going anywhere else enters the edit as before. A new font name drops the run's theme scheme and family, which belonged
to the old one. With no text selected these format the cells, as before (the edit entered first).

The formula bar keeps a cell's runs through an edit there (`src/editor/formulabar.js`, the same per-character fonts,
`richtrack.js`): the text shown plain, as Excel's bar shows it, the typed text taking the font before the caret, and
Ctrl+B, I, U, 5 and the Font group formatting the part selected there. (Before, a rich cell edited in the bar lost its
runs.) Not measured: Excel's bar is driven by keystrokes, which COM can't send; this follows its in-cell editor.

A run's rPr children that aren't modelled (`outline`, `shadow`, `condense`, `extend`, `charset`, anything newer) are
kept as they came and written back with the run, in Excel's order (`strings.js` runFont `extra`; `test/richtext.test.js`).
A run's `family` is read from the file too.

Printed and exported to PDF, a rich cell is drawn as its runs (print.js draws cells as the grid does; `src/pdf.js`
takes each run's text node in its own computed font; `test/pdf.mjs`).

Not yet: Format Cells' Font tab for a selection inside the cell (Ctrl+1 while editing), Ctrl+B with nothing selected to
type on in bold, the Font group's controls showing the selected characters' font, the drawn view's up and down arrows
across lines of different heights (the caret moves by the text box's own lines), and copying runs.

## Format Painter

Format Painter copies the selection (outlined, as a copy is, the system clipboard left alone) and pastes its formats
on the next cells selected, through Paste Special > Formats (measured for `test/paste-excel.json`). Pressed twice
quickly it stays on; Esc, or pressing it again, ends it.

## Draw Borders and More Colors (`formatops.js` drawLines, `src/ooxml/recentcolors.js`, `src/editor/drawborders.js`)

Home > Font > Borders' Draw Borders section and the More Colors... of the Fill Color, Font Color and Line Color menus.
Measured in a hidden Excel over COM (`tools/excel-compare/excel-drawborders.ps1`, Excel's saved borders cell by cell
in `test/drawborders-excel.json`; `test/drawborders.test.js` draws each scene here and expects Excel's text).
Drawing is a mouse gesture COM can't make, so the lines a drag draws are measured as the ribbon's presets on the same
cells (`ExecuteMso` Outside Borders, All Borders, No Border); the modes through `GetPressedMso` after each step.

| Rule | Seen |
|---|---|
| Draw Border, Draw Border Grid and Erase Border are modes, one at a time: picking another switches to it, picking the one on again ends it. | Draw, then Grid: only Grid pressed; Erase twice: none. |
| A cell selected ends the mode; a value typed doesn't. Here the drag draws and selects nothing, and Esc ends the mode (Excel's documented key). | `Range.Select` after Draw Border: none pressed. `Value2` set: still pressed. |
| Draw Border draws the outline of the range dragged, Draw Border Grid every line in it, Erase Border takes every line off, the diagonals too. | Outside Borders on H2:J4; All Borders on H8:J10; No Border on N14:P16 took O15's diagonal off. |
| The neighbours' facing lines go, whatever they are; a neighbour's line the same as the one drawn stays beside one cell and goes beside a range of several. | H1's double bottom and K3's red left went. G24's thin right stayed beside H24; G3's beside H2:J4, H7's beside H8:J10 and H19's beside H20:I20 (Top Border) went. |
| A click within 4px of a gridline draws or erases just that line (elsewhere the cell's own lines), with the neighbour's facing side going as above. | Excel's documented behaviour; not measurable over COM. |
| Line Style's thirteen lines are written `hair`, `dotted`, `dashDotDot`, `dashDot`, `dashed`, `thin`, `mediumDashDotDot`, `slantDashDot`, `mediumDashDot`, `mediumDashed`, `medium`, `thick`, `double`; an automatic Line Color is `indexed="64"`. Thin and automatic to start with. | One cell's bottom per line, A20 to A32. |
| A custom colour (More Colors) is written `rgb="FFRRGGBB"`: a line's `color`, a fill's `fgColor` (solid, `bgColor indexed="64"`), a font's `color`. | RGB 18, 52, 86: `rgb="FF123456"` on C20, C21, C22. |
| Recent Colors are the workbook's `<colors><mruColors>` in styles.xml (after `tableStyles`); Excel saved such a list back as it was. A colour set over COM added none. Here a More Colors colour goes first, once, at most ten (the schema's limit), and isn't taken back by Undo (the format it came with is). | Two colours put in the saved file: the same `<colors>` after Excel's save. |

Not measured (the Colors box and the Line Color and Line Style menus can't be driven over COM): which end of
`mruColors` Excel puts the latest at; that picking a line colour or style turns Draw Border on (Excel's documented
behaviour, followed here); whether the Line Color and Line Style also go to the Borders presets (here they don't).
Excel's Standard (hexagon) tab of the Colors box is not offered; the Custom tab's Red, Green, Blue and Hex are.

Found in passing, not changed here: over a range of several cells the ribbon's presets also take a neighbour's same
line away (the rows above follow it for Draw Borders only), and No Border takes the diagonals off, where the presets
here keep both.

## Cell Styles (`src/ooxml/cellstyleedit.js`, `src/editor/cellstylegallery.js`)

Home > Cell Styles has Excel's 47 built-in cell styles under its gallery's headings (Good, Bad and Neutral; Data and
Model; Titles and Headings; Themed Cell Styles; Number Format; the workbook's own under Custom), each named and drawn
in its look. The styles are measured: each applied in a new workbook over COM (`Range.Style`,
`tools/excel-compare/excel-cellstyles.ps1`) and what Excel saved kept in `test/cellstyles-excel.json`;
`tools/excel-compare/cellstyles-data.mjs` writes them into `src/engine/cellstyles.js`. The headings and their order
are Excel's gallery as it shows them (COM doesn't reach the gallery).

What Excel writes, and the editor with it (`test/cellstyle.test.js`: 47 of 47 styles, and the two cases below, the same
style entry, cellStyle and cell format as Excel's):

- A style the workbook hasn't got is added: its font written as a new entry (the theme's body font, as the Normal
  style's font names it; Title in the heading font; Excel's 11pt styles at the Normal style's size), its fill and
  border, a `cellStyleXfs` entry marking with `applyX="0"` each part it doesn't apply (Good: number format, border,
  alignment and protection; Currency: all but its number format), and a `cellStyle` with its `builtinId`, the
  cellStyles kept in name order (Good before Normal, Normal before Title).
- The cell's format takes from the style each part it applies and keeps its own of the rest; `xfId` is the style's,
  and each part that isn't the style's is marked `applyX="1"`: Currency over a plain cell keeps the cell's font
  (`applyFont="1"`); Good over a bold red cell takes Good's font (the bold and red go); Heading 1 over Good takes Heading
  1's font and border and keeps Good's fill (`applyFill="1"`). Normal applies every part: the cell back to plain.
- Each is one step of the history; on a protected sheet refused as formatting is.

### New Cell Style and Modify (`src/ooxml/cellstyledefine.js`, `src/editor/stylebox.js`)

The gallery ends with New Cell Style..., and each style's menu (a right click, or the menu key) has Apply and
Modify..., as Excel's. Both open Excel's Style box: the style's name (Excel's "Style 1" to start with; a built-in
style's can't be changed), the six parts it includes (Number, Alignment, Font, Border, Fill, Protection, each with a
line saying what it is) and Format..., the Format Cells box on the style's own format. Measured over COM
(`Styles.Add(name, BasedOn)` on a formatted cell, the style's Include..., Font, Interior, NumberFormat and alignment
set, for a style of the workbook's own, a built-in one and Normal, with cells using them;
`tools/excel-compare/excel-newstyle.ps1`, `test/newstyle-excel.json`; `test/newstyle.test.js` makes each case here and
gets Excel's style entry, cellStyle and cell formats, 5 of 5):

| Rule | Seen |
|---|---|
| New Cell Style is by example: the style's `cellStyleXfs` entry takes the active cell's number format, font, fill and border ids and its alignment and protection; a part left out is marked `applyX="0"`, in Excel's order. | B2 bold red 14pt, yellow, bottom border, 0.00, centred: `<xf numFmtId="2" fontId="1" fillId="2" borderId="1"><alignment horizontal="center"/></xf>`; border and alignment left out: `applyBorder="0" applyAlignment="0"`. |
| Its `cellStyle` has a name and no builtinId, among the others in name order. The cell keeps its own style. | `<cellStyle name="Mine" xfId="1"/>` before Normal; B2's style still Normal. |
| Modify changes the style's entry in place: a part changed is a new font, fill or border (or an equal one found). | Mine not bold, green, 0%, left: `<xf numFmtId="9" fontId="3" fillId="3" borderId="1"><alignment horizontal="left"/></xf>`. |
| Every cell format of the style then takes, in place, each changed part the style includes and the format doesn't hold as its own (`applyX="1"`). | D4 (style only) took all of it; D5 (italic of its own, `applyFont="1"`) kept its bold italic font and took the fill, format and alignment; B2's own format didn't change. |
| A part the style leaves out isn't given to its cells, even when changed. A built-in style changed is marked `customBuiltin="1"`. | Good made bold with a border: its cells bold, their border unchanged (`applyBorder="0"` kept). |
| Normal is changed the same way: every cell of no format of its own with it, a cell's own font kept. | Normal at 14pt: B2 (s 0) 14pt; C3 (bold of its own) 11pt. |

Here a style's name is refused when empty or already a style's (a built-in one's too, in any case: "A style with this
name already exists."). New and Modify are each one step of the history (a snapshot with the styles part,
`workbookstate.js`), refused on a protected sheet as formatting is. A file saved after both opened in Excel itself
without repair, with the styles and cells as written (Mine 16pt, I40 of it 16pt, J40's own italic font kept at 11pt).

Not measured: whether Excel's box gives the new style to the selection when OK is pressed (COM's Styles.Add doesn't;
the editor doesn't, which Microsoft's steps for New Cell Style describe). Excel rewrote Normal's font in place (font 0)
where the editor writes a new font entry, the same look; Excel's own Format... changes also show in its box's
Format Cells as they are made, where the editor's second Format... shows the font, number format and alignment chosen
so far but the fill and border as they were.

### Delete, Duplicate and Merge Styles (`src/ooxml/cellstylemanage.js`)

Each style's menu in the gallery has Duplicate... and Delete after Modify... (Normal's has no Delete: Excel refuses it,
"Unable to get the Delete property of the Style class" over COM), and the gallery ends with Merge Styles..., Excel's box
listing the other workbooks open on the page (`pagebooks.js`) and those the host gives (`styleBooks`). Measured over
COM (`Style.Delete`, `Styles.Merge` and the saved `styles.xml`; `tools/excel-compare/excel-stylesmore.ps1`,
`test/stylesmore-excel.json`; `test/stylesmore.test.js` makes each case here and gets Excel's style entries,
cellStyles and cell formats, 11 of 11):

| Rule | Seen |
|---|---|
| Delete a style of the workbook's own: its `cellStyle` and its `cellStyleXfs` entry go, the entries after it a place down with every `xfId` pointing at them. | Mine deleted: `cellStyleXfs` back to Normal's alone, `<cellStyle name="Mine">` gone; unused, the same. |
| Each cell format of the deleted style goes back to Normal in place: `xfId="0"`, Normal's number format, font, fill, border, alignment and protection, but for each part it marks as its own (`applyX="1"`), kept. The cells keep their `s`. | D4 (Mine alone): `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`; D5 (italic of its own, `applyFont="1"`): its bold italic red 14pt font kept, Mine's fill, border, 0.00 and centring gone. |
| Delete a built-in style: it stays, its `cellStyle` marked `hidden="1"`, its entry as it was; it leaves the gallery and can't be given to a cell again. Its cells go back to Normal as above. One the workbook hasn't got is added as applying it adds it, hidden. | Good deleted: `<cellStyle name="Good" xfId="2" builtinId="26" hidden="1"/>`; B2 Normal, B3 kept its red font; `Range.Style = "Good"` then fails. |
| Duplicate: Excel's Style box on a copy named "<name> 2" (the first number free after it); OK makes a style of the workbook's own (no `builtinId`), its entry the style's, with each part it leaves out Normal's; its `cellStyle` in name order. A built-in style the workbook hasn't got isn't added: the copy has the entry adding it would write. | Mine (border left out) duplicated: `<xf numFmtId="2" fontId="1" fillId="2" borderId="0" applyBorder="0"><alignment horizontal="center"/></xf>`; Good: "Good 2" with Good's green font and fill, no Good. |
| Merge Styles takes each style of the other workbook (not a hidden one): its font, fill and border written into this one (an equal one reused), its number format by its code (a new id from 164 up). One this workbook hasn't got is added, its `cellStyle` in name order (a built-in one with its `builtinId`, and `customBuiltin` if the other has it). | Other (bold, `0.0" kg"`, green) and Third (italic, the rest left out) added; `0.0" kg"` as `numFmtId="165"` after the workbook's own 164. |
| A style of the same name (a built-in one by its `builtinId`) asks Excel's "Merge styles that have the same names?": Yes puts the other's entry in place of this one's as Modify puts a change, the cells using it taking each part they don't hold as their own; a built-in style keeps `customBuiltin="1"` if either had it. No keeps this workbook's; Cancel merges nothing. | Mine replaced by the other's green italic 12pt, dashed, 0%, left: D4 took all of it, D5 kept its own font; Good made bold there: B6 bold; Normal at 14pt there: C3 14pt; Title (customBuiltin here) the other's plain Title, still `customBuiltin="1"`. |

Here the question is asked only when a style of the same name differs (COM's `Styles.Merge` answers it Yes without
asking, so when Excel asks wasn't seen; No was not measured, as only Excel's own box gives it). Each is one step of
the history (a snapshot with the styles part), refused on a protected sheet as formatting is. Excel numbers
`cellStyleXfs` anew when it saves (Normal first, then the styles by name, hidden ones last: the deleted Good moved after
Title); the editor keeps the entries where they are, the `xfId`s pointing at the same styles. The editor writes no
`xr:uid` on a `cellStyle`, where Excel gives each style of its own one. Duplicate's box isn't reached over COM: the copy
was made in Excel as its box makes it, by example from a cell of the style alone (`Styles.Add(name, cell)`) with the
style's Include... flags. A merged Normal is written as a new font where Excel rewrote font 0 in place, the same look,
as with Modify; Excel added Bad at 11pt after Normal was made 14pt, where the editor adds a built-in style at Normal's
size.
