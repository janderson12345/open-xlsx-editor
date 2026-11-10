// Real-time co-editing over Yjs (MIT): import { connectWorkbook } from 'open-xlsx-editor/collab'. Bring your own Yjs
// provider (y-websocket, y-webrtc...) for the transport. The room holds the workbook's bytes and, cell by cell, the
// contents typed since: { "<sheetId>!<row>,<col>": { formula } or { value } }. Contents edited in different cells
// at once all arrive; the last change to the same cell wins.
//
// What's shared live: cell contents (typing, pasting values, Delete, fill, Replace, the agent's writes and their undo).
// Anything else (formats, merges, rows and columns inserted or deleted, sheets, filters, sizes) is shared as the whole
// workbook: the editor that made it saves and publishes the bytes, and the others reopen them, on the sheet they were
// on. Freeze panes and zoom stay each person's own, as in Excel. Undo takes back this editor's own changes.
const ORIGIN = 'oxe-local';

/**
 * Shares the workbook open in `api` (mountXlsxEditor's) through `ydoc`, the Y.Doc your provider syncs. In a room that
 * has a workbook, that one is opened here; in an empty room, the one open here becomes the room's. `ready`: a promise
 * that settles once the provider has synced (y-websocket: `new Promise((r) => provider.once('synced', r))`), waited
 * for before deciding which. `awareness` (the provider's) gets `user` ({ name, color }). Opening another file in the
 * editor afterwards makes it the room's. Returns { stop() }.
 */
export async function connectWorkbook(api, ydoc, { ready = null, awareness = null, user = { name: 'Author', color: '#2b6cdf' } } = {}) {
  await ready;
  const files = ydoc.getMap('oxe-files');
  const cells = ydoc.getMap('oxe-cells');
  if (awareness) awareness.setLocalStateField('user', user);
  const open = api.open;
  let quiet = false; // (true while the room's changes are made here)
  let pending = false;

  const sheetById = (id) => api.workbook.sheets.findIndex((s) => s.sheetId === id);
  const keyOf = (si, r, c) => `${api.workbook.sheets[si].sheetId}!${r},${c}`;
  const parse = (key) => { const m = /^(\d+)!(\d+),(\d+)$/.exec(key); return m && { id: Number(m[1]), r: Number(m[2]), c: Number(m[3]) }; };

  /** Make the room's cell contents `keys` here. */
  function take(keys) {
    const bySheet = new Map();
    for (const key of keys) {
      const at = parse(key);
      const v = cells.get(key);
      if (!at || v === undefined) continue;
      const si = sheetById(at.id);
      if (si < 0) continue;
      if (!bySheet.has(si)) bySheet.set(si, []);
      bySheet.get(si).push(v.formula != null ? { r: at.r, c: at.c, formula: v.formula } : { r: at.r, c: at.c, value: v.value ?? null });
    }
    quiet = true;
    try {
      for (const [si, contents] of bySheet) {
        try { api.command('remoteContents', si, contents); } catch { /* (a cell another editor changed that is now part of an array here: left as it is) */ }
      }
    } finally { quiet = false; }
  }

  /** Open the room's workbook here, staying on the sheet shown (when it's still there), then its contents since. */
  function openRoom() {
    const shownId = api.workbook ? api.workbook.sheets[api.sheetIndex]?.sheetId : null;
    const sel = api.selection;
    quiet = true;
    try {
      open(files.get('xlsx'), api.fileName);
      const k = shownId == null ? -1 : sheetById(shownId);
      if (k >= 0 && k !== api.sheetIndex) api.showSheet(k);
      if (sel) api.command('select', sel.ar, sel.ac);
    } finally { quiet = false; }
    take([...cells.keys()]);
  }

  /** Publish the workbook here as the room's: its bytes, the contents so far in them. */
  function publish() {
    pending = false;
    if (!api.workbook) return;
    const bytes = api.save();
    ydoc.transact(() => {
      cells.clear();
      files.set('xlsx', bytes);
    }, ORIGIN);
  }

  if (files.get('xlsx')) openRoom();
  else publish();

  const onFiles = (e, tr) => { if (tr.origin !== ORIGIN && e.keysChanged.has('xlsx')) openRoom(); };
  const onCells = (e, tr) => {
    if (tr.origin === ORIGIN) return;
    // (Keys deleted are the room's contents going into a published workbook: that workbook comes with them.)
    take([...e.keysChanged].filter((k) => cells.has(k)));
  };
  files.observe(onFiles);
  cells.observe(onCells);
  const stopEdits = api.onEdit((e) => {
    if (quiet) return;
    if (e.book) {
      // (Once per change, after the editor has finished it: one command can make several.)
      if (!pending) { pending = true; queueMicrotask(publish); }
      return;
    }
    ydoc.transact(() => {
      for (const { si, contents } of e.cells) {
        for (const ch of contents) cells.set(keyOf(si, ch.r, ch.c), ch.formula != null ? { formula: ch.formula } : { value: ch.value ?? null });
      }
    }, ORIGIN);
  });
  api.open = (bytes, fileName = null) => {
    const out = open(bytes, fileName);
    publish();
    return out;
  };

  return {
    ydoc,
    stop() {
      files.unobserve(onFiles);
      cells.unobserve(onCells);
      stopEdits();
      api.open = open;
    },
  };
}
