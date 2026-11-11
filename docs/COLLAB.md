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

| Change | How it travels |
|---|---|
| Cell contents: typing, Delete, pasting values, fill, Replace, AutoSum, the agent's writes, and their undo | Live, cell by cell: a Yjs map keyed by sheet id, row and column, holding the formula or the value. Edits to different cells at the same moment all arrive. When two people change the same cell, the last change wins. Formulas are worked out again in each editor. |
| Everything else: formats, merges, rows and columns inserted or deleted, sheets, filters, widths and heights | As the whole workbook. The editor that made the change saves and publishes the bytes, and the others reopen them on the sheet and cell they were on. Contents typed elsewhere at the same time are applied on top. |
| Each person's selection | Shown to the others as a cursor (below), through the provider's awareness, never in the workbook. |
| Freeze panes, zoom | Not shared: each person's own view, as in Excel. |

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

- Undo takes back this editor's own changes only. When another editor's change is shared as the whole workbook,
  this editor reopens it and its undo history starts again.
- When rows or columns are inserted or deleted in one editor while another is typing in a cell below them, the typed
  cell keeps its address. It isn't moved with the rows, because cells are keyed by position.
- Cursors are drawn in the sheet's own overlay, not in frozen panes: another person's cell inside the frozen rows or
  columns is hidden under the pane.
- The flag is in the grid's overlay, which assistive technology doesn't read, as with this editor's own selection
  outline; a screen reader doesn't hear where the others are.
- Notes and threaded comments are not co-edited yet: one editor's changes to them don't reach the others.
- A workbook saved with a password is shared encrypted, so each editor joining asks for the password. Whether co-editing
  an encrypted workbook works end to end hasn't been tested.

## Tested

`test/collab.mjs` runs two editors, each in its own browser page, with their Y.Docs relayed to each other as a
provider would. It checks that:

- the second editor opens the first's workbook
- a number typed in the first shows in the second
- a formula typed in the second is worked out in the first
- two cells typed at the same moment arrive in both
- a row inserted in the first moves the second's cells down, typing stays in step after it, and undoing it moves them
  back in both
- bold applied in the first shows in the second with its contents kept
- a range selected in the first is drawn in the second: the active cell outlined at its place in the first's colour,
  the range shaded, the first's name on the flag in white on the dark red; the first doesn't draw its own cursor
- axe finds no violations with a flag shown, light and dark (axe leaves text in the overlay's flags as "incomplete",
  so the flag's contrast is checked by `test/collab-peers.test.js` instead: every colour given out, and its text)
- the cursor leaves the second's grid when the first shows another sheet, comes back with it, and goes when the first
  leaves the room

The test uses a stand-in for y-protocols' Awareness (the four members `connectWorkbook` uses), relayed between the
pages as a provider would.
