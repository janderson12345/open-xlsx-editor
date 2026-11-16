// A stand-in for mountXlsxEditor's api, for co-editing tests in Node (collab.test.js): a Workbook, the members
// connectWorkbook uses, and `edit(fn, e)` to change the workbook as a command would and say so (onEdit). Two hosts'
// Y.Docs are relayed to each other as a provider would, and the relay can be held to make changes at the same time.
import './setup.js';
import * as Y from 'yjs';
import { Workbook } from '../src/ooxml/workbook.js';
import { connectWorkbook } from '../src/collab.js';
import { digestSheet } from '../src/collab-state.js';
import { describeStyle } from '../src/collab-styles.js';
import { tabColorOf } from '../src/collab-room.js';

export function host(bytes = null) {
  const edits = new Set();
  const api = {
    workbook: null, sheetIndex: 0, fileName: null, selection: null,
    open(b) { api.workbook = Workbook.fromBytes(b); api.sheetIndex = 0; return api.workbook; },
    save() { return api.workbook.toBytes(); },
    showSheet(k) { api.sheetIndex = k; },
    onEdit(fn) { edits.add(fn); return () => edits.delete(fn); },
    command(name, ...args) {
      if (name === 'remoteChange') { args[0](api.workbook); return true; }
      if (name === 'select') return true;
      throw new Error(`no command ${name}`);
    },
    /** A command's change: `fn(wb)`, then onEdit with `e` (a change of anything but contents by default). */
    edit(fn, e = { book: true }) { const out = fn(api.workbook); for (const f of [...edits]) f(e); return out; },
    /** Contents typed: [{ r, c, value } or { r, c, formula }] of sheet si, as typing says it. */
    type(si, contents) { api.workbook.apply(si, contents); for (const f of [...edits]) f({ cells: [{ si, contents }] }); },
  };
  if (bytes) api.open(bytes);
  return api;
}

/** Y.Docs `docs` relayed to each other; `hold()` keeps updates back until `release()`, as a slow connection would. */
export function relay(...docs) {
  let held = null;
  for (const d of docs) {
    d.on('update', (u, origin) => {
      if (origin === 'relay') return;
      for (const o of docs) {
        if (o === d) continue;
        if (held) held.push([o, u]); else Y.applyUpdate(o, u, 'relay');
      }
    });
  }
  return {
    hold() { held = []; },
    release() { const list = held; held = null; for (const [o, u] of list) Y.applyUpdate(o, u, 'relay'); },
  };
}

/** Two (or `n`) hosts in one room, the first bringing `bytes`. */
export async function room(bytes, n = 2) {
  const hosts = [host(bytes), ...Array.from({ length: n - 1 }, () => host())];
  const docs = hosts.map(() => new Y.Doc());
  const net = relay(...docs);
  const rooms = [];
  for (let k = 0; k < n; k++) rooms.push(await connectWorkbook(hosts[k], docs[k]));
  return { hosts, docs, net, rooms };
}

/** Wait for the microtasks a change queues (shared once the command is done). */
export const settle = () => new Promise((r) => setTimeout(r, 0));

/** What a host's workbook holds, in terms that don't depend on its own format indexes: each sheet's name, state, tab
 *  colour, cells (contents and format described), rows and columns (format described). */
export function view(api) {
  const wb = api.workbook;
  const fmt = (s) => JSON.stringify(describeStyle(wb, Number(s) || 0));
  return wb.sheets.map((s, i) => {
    if (s.kind !== 'worksheet') return { name: s.name };
    const d = digestSheet(wb, i);
    const cells = [...d.cells].sort().map(([k, v]) => [k, v.c, v.s ? fmt(v.s) : null]);
    const line = (m, key) => [...m].sort((a, b) => a[0] - b[0]).map(([n, a]) => [n, { ...a, ...(a[key] != null ? { [key]: fmt(a[key]) } : {}) }]);
    return { name: s.name, state: s.state, color: tabColorOf(wb, i), cells, rows: line(d.rows, 's'), cols: line(d.cols, 'style'), format: d.format };
  });
}

/** The text of cell (r, c) of sheet si: its formula or its value. */
export function cellOf(api, si, r, c) {
  const cell = api.workbook.sheet(si).cells.get(r)?.get(c);
  return cell ? (cell.f ? `=${cell.f}` : cell.v) : null;
}
