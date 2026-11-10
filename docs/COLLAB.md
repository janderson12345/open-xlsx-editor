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
| Freeze panes, zoom, the selection | Not shared: each person's own view, as in Excel. |

## Limits

- Undo takes back this editor's own changes only. When another editor's change is shared as the whole workbook,
  this editor reopens it and its undo history starts again.
- When rows or columns are inserted or deleted in one editor while another is typing in a cell below them, the typed
  cell keeps its address. It isn't moved with the rows, because cells are keyed by position.
- Other people's cursors aren't drawn yet. `awareness` carries the user's name and colour for a host page to show.
- Notes and threaded comments are read-only in the editor, so they are not co-edited.
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
