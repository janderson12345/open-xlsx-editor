# Co-editing

`src/collab.js` (`import { connectWorkbook } from 'open-xlsx-editor/collab'`) lets several editors work on one
workbook at the same time over [Yjs](https://github.com/yjs/yjs) (MIT). You bring the Yjs provider (y-websocket,
y-webrtc and so on) for the transport. Yjs is an optional peer dependency: only a page that co-edits needs it.

```js
import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import { mountXlsxEditor } from 'open-xlsx-editor';
import { connectWorkbook } from 'open-xlsx-editor/collab';

const editor = mountXlsxEditor(host, { bytes });            // the first in the room brings the workbook
const ydoc = new Y.Doc();
const provider = new WebsocketProvider(url, 'budget-2027', ydoc);
const room = await connectWorkbook(editor, ydoc, {
  ready: new Promise((r) => provider.once('synced', r)),   // wait for the room before deciding whose workbook it is
  awareness: provider.awareness,
  user: { name: 'Avery', color: '#2b6cdf' },
});
// room.stop() leaves the room.
```

In a room that already has a workbook, that one opens. In an empty room, the workbook open in this editor becomes the
room's. Opening another file afterwards makes that file the room's.

## What is shared, and how

Every change goes to the others as what it is, piece by piece, except the few kinds the room doesn't hold that way
(last row of the table), which go as the whole workbook.

| Change | How it travels |
|---|---|
| Cell contents: typing, Delete, pasting values, fill, Replace, AutoSum, the agent's writes, and their undo | Each cell's formula or value. Edits to different cells at the same moment all arrive; when two people change the same cell, the last change wins. Formulas are worked out again in each editor. |
| Cell formats: font, fill, borders, alignment, number format, Clear Formats, the Format Painter | Each cell's format, by what it is (its font, fill, border and number format written out), not by its index in one editor's styles part. Each editor finds an equal format of its own or adds one at the end of its styles part, as formatting the cell there would, so indexes cells already use never move. |
| Rows and columns: heights, widths, hidden and shown, AutoFit, a whole row's or column's format, grouping (outline levels), the sheet's default width and height | Each row's and column's own attributes, as the file has them. |
| Rows and columns inserted and deleted | As operations: "3 rows after row 12", "rows 4 to 6 deleted". Each editor makes them with its own Insert and Delete code, so formulas, merged ranges, notes, charts and tables move the same way everywhere. |
| Sheets added, renamed, moved, hidden and shown, coloured, deleted | The sheet list (its order) and each sheet's name, visibility and tab colour. |
| Undo and Redo of any of these | Piece by piece too: an insert's Undo deletes the rows it made, a delete's Undo inserts the rows again with their cells and formats, a size's Undo puts the sizes back. What others changed meanwhile stays. |
| Each person's selection | Shown to the others as a cursor (below), through the provider's awareness, never in the workbook. |
| Freeze panes, zoom | Not shared: each person's own view, as in Excel. |
| Anything else: a cell style the workbook hadn't used yet (Cell Styles, Percent, Comma, Currency: the first use adds it to the styles part), merges, conditional formats, validation, tables, filters and sorts' saved state, defined names, notes, links, charts, pictures, PivotTables, protection, a sheet copied, Undo of a sheet deleted | As the whole workbook. The editor that made the change saves and publishes the bytes as a new epoch of the room, and the others open them on the sheet and cell they were on. Their undo history starts again. |

### Rows and columns that keep who they are

Cells, formats and row heights are kept in the room under the identity of their row and column, not under their
number: a row of the room's workbook by its number in that workbook, a row inserted since by the insert that made it
and its place in it. When one person types in A7 while another inserts two rows at the top, the typing lands in the
row that held A7 when it was typed (A9 by then) in both editors. The order of the rows is a replicated sequence
(RGA): an insert goes after the row above it, past rows a later insert put there, and a deleted row stays as a
marker, so two people inserting rows at the same place at the same time both get theirs, in the same order
everywhere, and an insert next to a row someone else deleted still finds its place.

A formula typed while another person inserts or deletes rows is moved by the inserts and deletes its writer hadn't
seen, as a formula already in the workbook is moved by them: `=B4+1` typed as someone inserts a row above row 4 reads
`=B5+1` in both editors.

### The room

A `Y.Map` named `oxe-files` holds the workbook's bytes (`xlsx`) as they were when the room's current epoch began, and
the epoch's id (`epoch`). `oxe-rooms` holds, under the epoch's id, a `Y.Map` with:

| Part | What it holds |
|---|---|
| `sheets` (`Y.Array`) | The sheets' ids in order: `s<sheetId>` for the epoch's own, `<client>.<n>` for one added since. |
| `props` (`Y.Map`) | `<sheet>/name`, `<sheet>/state`, `<sheet>/color`, `<sheet>/gone`. |
| `lines` (`Y.Array`) | Rows and columns inserted (`{ kind: 'ins', after, n }`) and deleted (`{ kind: 'del', runs }`), by identity, each with its id, Lamport time, client and sheet. |
| `cells` (`Y.Map` of `Y.Map`s) | For each sheet: `v <row>,<col>` its contents (`{ f }` or `{ v }`), `s <row>,<col>` its format's key, `r <row>` and `c <col>` a line's own attributes, `fmt` the sheet's defaults. |
| `styles` (`Y.Map`) | Each format's key to its description. |

A new epoch starts with an empty map; the old epoch's is deleted, so changes made in it at the same moment can't land
in the wrong place. What another editor writes is checked before any of it reaches the workbook (`src/collab-check.js`:
XML names, text values, whole numbers within the sheet, hex colours); an entry that isn't what this version writes is
left out.

## Others' cursors

With `awareness`, each editor sends its selection (the sheet, by its id, the active cell and every range) when it
changes, at most every 100 ms with the last move always sent, as the `cursor` field of its awareness state beside
`user`. Each editor draws the others' on the sheet it shows, the way Excel for the web shows co-authors (Microsoft's
support article "Collaborate on Excel workbooks at the same time with co-authoring"): a coloured outline round the
other person's active cell (or the merged range it is in), their other selected cells lightly shaded in that colour,
and their name on a small flag in the colour at the cell's top-right corner (below the cell on row 1, where the
column headings would cover it). Excel for the web shows the name when the pointer rests on the outline; here the
flag stays shown, as the grid's overlay doesn't take the pointer. COM drives desktop Excel, which has no co-authors'
cursors to measure, so this follows Excel's documented look rather than a measurement.

- The colour is `user.color` (a `#rgb` or `#rrggbb` colour; anything else is ignored), or one of eight given out by
  awareness client id, each at least 4.5:1 against the sheet's white paper (which stays white in dark mode). The
  name on the flag is black or white, whichever contrasts more with the colour.
- This editor's own selection is drawn above the others'.
- A cursor goes when its editor moves to another sheet, calls `room.stop()`, or its provider drops it from awareness
  (y-protocols does so when the connection closes, or after 30 seconds without word).
- Without `awareness` nothing is sent or drawn.

## Limits

- Undo takes back this editor's own changes only. When another editor's change is shared as the whole workbook, this
  editor opens it and its undo history starts again.
- Undo of rows deleted puts their cells and formats back, but not formulas elsewhere that read them: those keep the
  #REF! the delete gave them (Excel's own Undo restores them).
- Two people deleting and inserting rows at the same place at the same moment end with the same rows in the same
  order, but a formula reading a range across both may be moved differently in each (one editor deleted then
  inserted, the other inserted then deleted).
- Cells inserted or deleted with the rest of their columns or rows shifting go as the cells that moved (each one's
  new contents and format), not as an operation: a cell typed in the shifted range at the same moment keeps its
  address. When merges, rules or validation move with them, they go as the whole workbook.
- The colour menus' Recent Colors (kept in the styles part) stay each editor's own.
- A cell's rich text (runs of different fonts) and an error typed as a value go across as plain text.
- A sheet added in two editors gets a different sheetId in each, so a cursor on it may not be drawn in the other.
- Comparing an editor's workbook with what it last shared takes one pass over each sheet it has open after each
  change that isn't typing: about 42 ms for a sheet of 100,000 cells (test/fixtures/large.xlsx, Node 26). Typing
  compares only the cells typed.
- Cursors are drawn in the sheet's own overlay, not in frozen panes: another person's cell inside the frozen rows or
  columns is hidden under the pane.
- The flag is in the grid's overlay, which assistive technology doesn't read, as with this editor's own selection
  outline; a screen reader doesn't hear where the others are.
- Notes and threaded comments are not co-edited piece by piece: a change to one goes as the whole workbook.
- A workbook saved with a password is shared encrypted, so each editor joining asks for the password. Whether co-editing
  an encrypted workbook works end to end hasn't been tested.
- A room written by an earlier version of this module (no epoch) is opened and then shared again in this version's
  form.

## Tested

`test/collab.test.js` runs two editors' workbooks in Node, each with its own Y.Doc, relayed to each other as a
provider would and held back to make changes at the same moment. It checks that:

- the line sequence puts inserts after one row in one order whatever order they are applied in, passes deleted rows
  over and still places an insert next to one
- contents typed go across cell by cell, in the same epoch
- a format added in each at once goes across as the format, at a different index in each
- row heights, column widths, hidden columns and a row's own format go across
- rows inserted and a column deleted go across as operations, formulas moved and reading #REF! as in the editor that
  made them, typing in step after them
- two editors inserting rows at the same place at once both get theirs, in one order
- one editor formatting, typing and resizing while the other inserts and deletes rows above: the same cells in both
- a formula typed while the other inserts a row reads the cell it was typed to read
- sheets added, renamed, moved, hidden, coloured and deleted go across, two added at once both arrive, and typing on a
  new sheet goes across
- an editor joining later opens the room with every change since
- a merge goes as the whole workbook, in a new epoch, and changes go piece by piece again after it
- rows grouped go piece by piece; a cell style the workbook hadn't used yet goes as the whole workbook
- room entries that aren't what this version writes (a row entry naming its own place, a bad attribute or element
  name, contents that aren't a value, a colour that isn't hex, an insert of a billion rows) are left out

Most of these end with the two workbooks compared whole (every sheet's cells, formats by description, rows, columns,
defaults, names, states and tab colours): equal.

`test/collab.mjs` runs two editors, each in its own browser page, with their Y.Docs relayed to each other as a
provider would (and held back for changes made at once). It checks that:

- the second editor opens the first's workbook
- a number typed in the first shows in the second
- a formula typed in the second is worked out in the first
- two cells typed at the same moment arrive in both
- a row inserted in the first moves the second's cells down, typing stays in step after it, and undoing it moves them
  back in both
- bold applied in the first shows in the second with its contents kept, and none of these changed the room's epoch
- a row height and a hidden column go across, and so does their Undo
- a row inserted at the same place in both at once, and typed in: both rows in both, in one order
- a cell made italic in the first while the second inserts a row at the top: italic, a row down, in both
- a sheet added, renamed, coloured and moved first in the first and another hidden in the second: the same tabs in
  both, the tab colour too; the move undone in the first moves it back in both, and redone moves it first again
- a row deleted in the second and its Undo: the row's cells back in both, the first's typing below moved with its
  row; the epoch unchanged through all of it
- a merge goes as the whole workbook (a new epoch), and typing goes on after it
- a range selected in the first is drawn in the second: the active cell outlined at its place in the first's colour,
  the range shaded, the first's name on the flag in white on the dark red; the first doesn't draw its own cursor
- axe finds no violations with a flag shown, light and dark (axe leaves text in the overlay's flags as "incomplete",
  so the flag's contrast is checked by `test/collab-peers.test.js` instead: every colour given out, and its text)
- the cursor leaves the second's grid when the first shows another sheet, comes back with it, and goes when the first
  leaves the room

The test uses a stand-in for y-protocols' Awareness (the four members `connectWorkbook` uses), relayed between the
pages as a provider would.
