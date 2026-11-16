# Find and Replace

How cells are searched and replaced (`src/ooxml/findops.js`, the bar in `src/editor/findbar.js`), and what Excel was
seen to do. Each rule was measured over COM (`Range.Find` and `FindNext`, `Cells.Replace`, on a new workbook in a
hidden instance). `npm test` (`test/find.test.js`) repeats each search and replacement and expects Excel's result.
After replacements across formulas.xlsx, Excel works out all 399 formulas as the editor does.

## In the editor

- Ctrl+F opens the bar under the ribbon (as open-docx-editor's), Ctrl+H with its Replace row; Home > Editing > Find &
  Select has both. Enter finds the next match, Shift+Enter the previous; Esc closes the bar and goes back to the sheet.
- Options as Excel's dialog has them: Match case, Match entire cell contents, Within (Sheet or Workbook), Search (By
  Rows or By Columns), Look in (Formulas, Values, Notes or Comments; Replace always looks in formulas, as Excel's does).
- Find All says how many cells match and lists them as Excel's dialog does (sheet, cell, value, formula), each going
  to its cell when clicked; Replace replaces the active cell when it matches and goes on to the next;
  Replace All replaces every match as one step of the undo history.
- The bar keeps focus while the matches are selected, as Excel's dialog does; its status line is read out by screen
  readers.

## Excel's rules

| Rule | Seen |
|---|---|
| A search starts at the cell after the active one and goes round to the start; the active cell is found last. | "apple" from A1: B1, C1, A1. From B1: C1, A1, B1. |
| Look in Formulas searches the text of the formula bar: a formula with its `=`, a percentage as a percentage, a date as a date. | `50%` finds 0.5 shown as 50%, `0.5` finds nothing; `1/1/2024` finds a date, `45292` doesn't; `=` finds the formulas. |
| Look in Values searches the text cells show. | `1` finds `=A2*2` showing 1 and the date 1/1/2024, not `=A1+A10`. `A2` finds nothing. |
| Case is ignored unless Match case. Match entire cell contents compares the whole text. | `Apple` with Match case: B1 only. `apple` whole: A1 only. |
| `*` stands for any characters, `?` for one, `~` makes the next character plain. | `a*b` whole: `a*b` and `axxb`; `a~*b`: `a*b` only; `a?xb`: `axxb`. |
| Replace changes every match within a cell (or the whole cell), and the result is entered as if typed: a number becomes a number, `=...` a formula, nothing empties the cell. | `apple` with `pear`: `pear pie`, `pinepear`. `1` with `9`: 12 becomes the number 92. `12` with `=1+1`: a formula. |
| In a formula, replacing is plain text: references aren't recognised. | `A1` with `B9`: `=A1+A10` becomes `=B9+B90`. |
| Replace doesn't look in values: a formula's result isn't replaced. | `apple` doesn't change `=A1&"x"`, which shows `applex`. |
| Look in Notes searches the notes' text, any capitals; Look in Comments the threaded comments'; neither finds the other or the cells, and Formulas finds neither. | `apple` in notes: A1 ("Check the apple stock") and E5 ("APPLES"); in comments: C3; in formulas: D4 alone (`tools/excel-compare/excel-find2.ps1`, `test/find2-excel.json`). |

The messages are Excel 365's: "We couldn't find what you were looking for. Select Options for more ways to search."
and "All done. We made 3 replacements." (Excel shows them in a message box; here they are in the bar's status line.
They come from Excel's dialog, which COM doesn't reach, so their wording isn't measured.)

## Formats (Find Format and Replace Format)

Beside Find what and Replace with: Format... (the format to find, or the one a replacement gives: number format, bold,
italic, underline, strikethrough, font colour, fill, a border), From Cell... (Excel's Choose Format From Cell: the next
cell clicked on the sheet, or the one Enter is pressed on, gives its whole format; Esc gives up) and Clear. The format
chosen is shown beside the box, with its fill, ink and weight. `src/ooxml/findformat.js` resolves each cell's format to
what it looks like (theme colours and tints worked out, automatic ink as black, the theme's font for a scheme font) and
compares that.

Measured two ways. Excel's own dialog (used in its window; `test/find-format-excel.json`), as COM can't run a format
search with FindNext or an empty Replace (`excel-find2.ps1`). And `Application.FindFormat` with `Range.Find`, each
match found from the one before (`tools/excel-compare/excel-findformat.ps1`, 24 cases, `test/find-format2-excel.json`,
`test/findformat.test.js`): 19 of them as Excel, the other five not measurable that way (below).

| Rule | Seen |
|---|---|
| A cell matches when it has every property the find format sets; the others are free. | Bold alone: the bold cells, whatever else. Arial: the Arial cells of 11pt and of 14pt. |
| Nothing in Find what, a find format: every filled cell of that format. | Find All: A1, B1, A3, the bold ones, in rows. |
| Colours compare as they look: a theme colour and the same RGB match each other, a tint of the theme colour doesn't. | Accent 1 (theme) or 156082 (RGB): both cells, not Accent 1 lighter 40%. Fills the same (Accent 2, E97132). |
| Automatic ink is black. | Black (RGB 0 or Text 1): the automatic cell, the RGB black one and the Text 1 one. |
| "General" finds the cells with no number format of their own; a code finds that code only. | `0.00`: not `#,##0.00`. |
| Alignment and protection compare one property each. | Centred: centred and centred-and-wrapped. Indent 2. Unlocked. |
| A border line in a colour finds that side drawn so. | Bottom thin red: the red one only. |
| Replace All replaces in the cells of the find format alone, and gives them the replace format, keeping their own. | apple to fruit, bold ones only, a yellow fill: 2 replaced, still bold; the plain apple and the bold pear as they were. |
| Nothing in either box: only the format replaced, the text kept. | 3 replaced: the three bold cells yellow, their text as it was. |

Not measurable over COM, so done as Excel's dialog describes them: a border asked for without a colour (COM's
`FindFormat.Borders` found nothing at all, not even Excel's own thin line) here matches the line in any colour; Choose
Format From Cell copies every property of the cell (COM has no such call: setting each property from the cell found two
of the three look-alikes and one that was italic, as COM takes `Italic = False` for "either"), so here only cells that
look the same in every property match, and as a Replace Format the cell takes that format whole, theme colours kept.

## Not yet

A find format's font name, size, alignment and protection can be chosen only from a cell, not in the Format... box.
(Go To Special: docs/GOTOSPECIAL.md.) (Replacing within a date now works as in Excel: docs/ENTRY.md.)

## Found on the way

Replacing 5 with 7 in formulas.xlsx left a column unsorted (0, 70, 60, 70, 80), where Excel's approximate `VLOOKUP` and
`MATCH` gave answers the engine didn't: Excel searches such lists its own way. The engine now does the same
(docs/CALC.md), and after those replacements Excel works out all 399 formulas as the editor does.
