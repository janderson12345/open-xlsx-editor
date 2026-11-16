// Real-time co-editing over Yjs (MIT): import { connectWorkbook } from 'open-xlsx-editor/collab'. Bring your own Yjs
// provider (y-websocket, y-webrtc...) for the transport.
//
// The room holds the workbook's bytes as they were when it became the room's, and every change since, piece by piece
// (collab-room.js has the room's parts): cell contents and formats, rows' and columns' heights, widths, hiding and
// formats, rows and columns inserted and deleted, and the sheets added, renamed, moved, hidden, coloured and deleted.
// Rows and columns are known by an identity that stays with them (collab-seq.js), so a cell typed or formatted in one
// editor while another inserts rows above it lands in the same cell in both, and two editors inserting at once both
// get their rows. Formats go by what they are (collab-styles.js), not by their index in one editor's styles part.
// Each editor catches its own changes as they are made (collab-capture.js) and makes the others' with the code its
// own commands use.
//
// A change the room doesn't hold piece by piece (merges, rules, tables, filters, names, notes, drawings, cells shifted
// within columns) is shared as the whole workbook: the editor that made it saves and publishes the bytes as a new
// epoch of the room, and the others open them, on the sheet they were on. Freeze panes and zoom stay each person's
// own, as in Excel. Undo takes back this editor's own changes, piece by piece where the room has them so.
//
// With the provider's awareness, each editor's selection goes to the others, who draw it on the grid in that person's
// colour with their name (collab-peers.js).
import { shareCursors } from './collab-peers.js';
import { styleKeys } from './collab-styles.js';
import { newRoom, refreshAll, refreshSheet, patchCells, indexOfUid, applyOp, applyEntries, reconcileSheets, replayRoom } from './collab-room.js';
import { flush, flushCells, wrapWorkbook, unwrapWorkbook } from './collab-capture.js';
import { bookPrint, sheetPrint } from './collab-state.js';

const ORIGIN = 'oxe-local';

/** The state of one connection, made fresh for each workbook opened: what collab-room.js and collab-capture.js share. */
function attach(st, wb, room, epoch) {
  if (st.wb && st.wb !== wb) unwrapWorkbook(st.wb);
  const own = (id) => (typeof id === 'string' && id.startsWith(`${st.cid}.`) ? Number(id.slice(String(st.cid).length + 1)) : 0);
  let seqNo = 0;
  if (room) {
    for (const op of room.get('lines').toArray()) if (op.cid === st.cid) seqNo = Math.max(seqNo, op.seq);
    for (const uid of room.get('sheets').toArray()) seqNo = Math.max(seqNo, own(uid));
  }
  const fresh = st.wb !== wb;
  Object.assign(st, {
    wb, room, epoch, seqNo, lamport: 0, known: {}, applied: new Set(), order: [],
    uids: new Map(wb.sheets.map((s) => [s.sheetId, `s${s.sheetId}`])), seqs: new Map(), shadow: new Map(), prints: new Map(),
    keys: styleKeys(wb), since: st.journal.length, lastSnap: null,
  });
  if (fresh) { st.snapMeta = new WeakMap(); wrapWorkbook(st); }
}

/**
 * Shares the workbook open in `api` (mountXlsxEditor's) through `ydoc`, the Y.Doc your provider syncs. In a room that
 * has a workbook, that one is opened here; in an empty room, the one open here becomes the room's. `ready`: a promise
 * that settles once the provider has synced (y-websocket: `new Promise((r) => provider.once('synced', r))`), waited
 * for before deciding which. `awareness` (the provider's) gets `user` ({ name, color }) and this editor's cursor, and
 * the others' cursors are drawn from it. Opening another file in the editor afterwards makes it the room's. Returns
 * { stop() }.
 */
export async function connectWorkbook(api, ydoc, { ready = null, awareness = null, user = { name: 'Author', color: '#2b6cdf' } } = {}) {
  await ready;
  const files = ydoc.getMap('oxe-files');
  const rooms = ydoc.getMap('oxe-rooms');
  // (The Y.Map and Y.Array of the Yjs this doc was made with, so nested parts never mix two copies of Yjs.)
  const Y = { Map: rooms.constructor, Array: ydoc.getArray('oxe-types').constructor };
  let cursors = null; // (others' cursors, once the room's workbook is open here)
  if (awareness) awareness.setLocalStateField('user', user);
  const open = api.open;
  const st = {
    ydoc, Y, origin: ORIGIN, cid: ydoc.clientID, wb: null, room: null, epoch: null, journal: [], quiet: 0, depth: 0, tick: 0,
    host: {
      publish: () => publish(),
      remote: (run, sheets) => api.command('remoteChange', run, { sheets }),
    },
  };

  /** Publish the workbook here as the room's: its bytes, as a new epoch with nothing changed since. */
  function publish() {
    if (!api.workbook || st.publishing) return;
    st.publishing = true;
    try {
      const bytes = api.save();
      const epoch = `${st.cid}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
      let room;
      ydoc.transact(() => {
        room = newRoom(st, api.workbook);
        for (const k of [...rooms.keys()]) rooms.delete(k);
        rooms.set(epoch, room);
        files.set('xlsx', bytes);
        files.set('epoch', epoch);
      }, ORIGIN);
      attach(st, api.workbook, room, epoch);
      refreshAll(st);
    } finally { st.publishing = false; }
  }

  /** Open the room's workbook here, with every change made since, staying on the sheet shown (when it's still there). */
  function openRoom() {
    const room = rooms.get(files.get('epoch'));
    const shownId = api.workbook ? api.workbook.sheets[api.sheetIndex]?.sheetId : null;
    const shownUid = shownId == null ? null : st.uids?.get(shownId);
    const sel = api.selection;
    st.quiet++;
    try {
      open(files.get('xlsx'), api.fileName);
      attach(st, api.workbook, room || null, files.get('epoch'));
      if (room) api.command('remoteChange', () => { replayRoom(st); refreshAll(st); }, { sheets: true });
      const k = shownUid == null ? -1 : indexOfUid(st, shownUid);
      if (k >= 0 && k !== api.sheetIndex && api.workbook.canShow(k)) api.showSheet(k);
      if (sel) api.command('select', sel.ar, sel.ac);
    } finally { st.quiet--; }
    // (A room written by an older version of this file, without its parts: made this version's.)
    if (!room) publish();
    cursors?.redraw();
  }

  /** Others' changes to the room (`events`, observeDeep's), made here. */
  function onRoom(events, tr) {
    if (tr.origin === ORIGIN || !st.room) return;
    let sheets = false;
    let lines = false;
    const cells = new Map();
    for (const e of events) {
      const path = e.path;
      if (path[0] !== st.epoch || path.length < 2) continue;
      if (path[1] === 'sheets' || path[1] === 'props') sheets = true;
      else if (path[1] === 'lines') lines = true;
      else if (path[1] === 'cells' && path.length === 2) for (const uid of e.keysChanged) cells.set(uid, null);
      else if (path[1] === 'cells' && path.length === 3) {
        const uid = path[2];
        if (cells.get(uid) === null) continue;
        if (!cells.has(uid)) cells.set(uid, new Set());
        for (const k of e.keysChanged) cells.get(uid).add(k);
      }
    }
    if (!sheets && !lines && !cells.size) return;
    flush(st);
    const done = new Map();
    st.host.remote(() => {
      st.quiet++;
      try {
        if (sheets) reconcileSheets(st);
        if (lines) for (const op of st.room.get('lines').toArray()) applyOp(st, op);
        const data = st.room.get('cells');
        for (const [uid, keys] of cells) if (data.has(uid)) done.set(uid, applyEntries(st, uid, keys ? [...keys] : [...data.get(uid).keys()]));
      } finally { st.quiet--; }
      if (sheets || lines) refreshAll(st);
      else {
        for (const [uid, out] of done) {
          const i = indexOfUid(st, uid);
          if (i < 0 || !st.wb._parsed.has(st.wb.sheets[i])) continue;
          if (out.lines) refreshSheet(st, i); else patchCells(st, i, out.cells);
          st.prints.set(uid, sheetPrint(st.wb, i));
        }
        st.book = bookPrint(st.wb);
      }
    }, sheets);
    cursors?.redraw();
  }

  if (files.get('xlsx')) openRoom();
  else publish();
  if (awareness) cursors = shareCursors(api, awareness);

  const onFiles = (e, tr) => { if (tr.origin !== ORIGIN && (e.keysChanged.has('xlsx') || e.keysChanged.has('epoch'))) openRoom(); };
  files.observe(onFiles);
  rooms.observeDeep(onRoom);
  let pending = false;
  const stopEdits = api.onEdit((e) => {
    if (st.quiet) return;
    if (api.workbook !== st.wb) return;
    if (e.book) {
      // (Once per change, after the editor has finished it: one command can make several.)
      if (!pending) { pending = true; queueMicrotask(() => { pending = false; flush(st); }); }
      return;
    }
    flushCells(st, e.cells || []);
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
      rooms.unobserveDeep(onRoom);
      stopEdits();
      cursors?.stop();
      if (st.wb) unwrapWorkbook(st.wb);
      api.open = open;
    },
  };
}
