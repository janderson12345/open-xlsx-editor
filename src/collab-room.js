// The room of a co-edited workbook (collab.js), piece by piece. A room holds the workbook's bytes as it was when it
// became the room's (its epoch), and since then:
//
//   sheets  Y.Array   the sheets' ids in order ('s<sheetId>' for the epoch's own, '<client>.<n>' for one added since)
//   props   Y.Map     '<sheet>/name', '/state', '/color', '/gone': each sheet's name, visibility, tab colour, deletion
//   lines   Y.Array   rows and columns inserted and deleted, in order: { kind: 'ins', after, n } or { kind: 'del',
//                     runs }, by line identity (collab-seq.js), each with its id, Lamport time and client
//   cells   Y.Map     a Y.Map for each sheet: 'v <row>,<col>' its contents, 's <row>,<col>' its format's key,
//                     'r <row>' and 'c <col>' a line's own attributes, 'fmt' the sheet's defaults; by line identity
//   styles  Y.Map     each format's key to its description (collab-styles.js)
//
// Each editor keeps a shadow of what it last shared or took (collab-state.js's digest) and, after a change of its
// own, writes what differs; what the others write it makes in its workbook with the same code its own commands use.
// `st` is the shared state of one connection, as one object (collab.js makes it).
import { frameOf } from './ooxml/sheetcells.js';
import { LineSeq } from './collab-seq.js';
import { contentsOf, digestSheet, sheetPrint, bookPrint, samePrint, sameContents, sameAttrs, writeRows, writeCols, writeFormat } from './collab-state.js';
import { shiftRefs } from './calc/tokens.js';
import { checkOp, checkAttrs, checkStyle, checkContents, checkColor, isToken } from './collab-check.js';

const EMPTY = { cells: new Map(), rows: new Map(), cols: new Map(), format: {} };
/** A row's and a column's attributes that say where it is: never taken from the room. */
const ROW_OWN = new Set(['r', 'spans']);
const COL_OWN = new Set(['min', 'max']);

export const isGrid = (wb, i) => { const s = wb.sheets[i]; return !!s && s.kind === 'worksheet' && !!s.path && wb.pkg.has(s.path); };
export const uidAt = (st, i) => st.uids.get(st.wb.sheets[i]?.sheetId);
export const indexOfUid = (st, uid) => st.wb.sheets.findIndex((s) => st.uids.get(s.sheetId) === uid);
export function seqOf(st, uid, axis) {
  const k = `${uid}|${axis}`;
  if (!st.seqs.has(k)) st.seqs.set(k, new LineSeq());
  return st.seqs.get(k);
}
const part = (st, name) => st.room.get(name);
/** The room's Y.Map of a sheet's cells and lines (made with the sheet: two editors never make one at once). */
export function dataOf(st, uid) {
  const cells = part(st, 'cells');
  if (!cells.has(uid)) cells.set(uid, new st.Y.Map());
  return cells.get(uid);
}

/** A new room for epoch `epoch`: its parts, the sheets of `wb` (their ids) and a cells map for each. */
export function newRoom(st, wb) {
  const room = new st.Y.Map();
  const sheets = new st.Y.Array();
  const cells = new st.Y.Map();
  sheets.push(wb.sheets.map((s) => `s${s.sheetId}`));
  for (const s of wb.sheets) cells.set(`s${s.sheetId}`, new st.Y.Map());
  for (const [k, v] of [['sheets', sheets], ['props', new st.Y.Map()], ['lines', new st.Y.Array()], ['cells', cells], ['styles', new st.Y.Map()]]) room.set(k, v);
  return room;
}

/** What this editor holds, taken as shared: every parsed sheet's digest and print, the sheet list, the workbook's print. */
export function refreshAll(st) {
  st.digesting = true;
  try {
    for (let i = 0; i < st.wb.sheets.length; i++) if (st.wb._parsed.has(st.wb.sheets[i]) && isGrid(st.wb, i)) refreshSheet(st, i);
  } finally { st.digesting = false; }
  st.list = st.wb.sheets.map((s) => s.sheetId).join();
  st.book = bookPrint(st.wb);
}
export function refreshSheet(st, i) {
  const was = st.digesting;
  st.digesting = true;
  try {
    st.shadow.set(uidAt(st, i), digestSheet(st.wb, i));
    st.prints.set(uidAt(st, i), sheetPrint(st.wb, i));
  } finally { st.digesting = was; }
}

/** The shadow of sheet i taken again at cells `cells` ([{ r, c }]) only: after the room's contents and formats were
 *  made there (a whole sheet's digest per keystroke would cost a pass over all its cells). */
export function patchCells(st, i, cells) {
  const sh = st.shadow.get(uidAt(st, i));
  if (!sh) return refreshSheet(st, i);
  const { model } = st.wb._part(i);
  for (const { r, c } of cells) {
    const cell = model.cells.get(r)?.get(c);
    const contents = contentsOf(st.wb, i, model, r, c, cell);
    if ((contents?.v == null && contents?.f == null) && !cell?.s) sh.cells.delete(`${r},${c}`);
    else sh.cells.set(`${r},${c}`, { c: contents, s: cell?.s || 0 });
  }
}

/** A format's key, its description put in the room the first time. */
function styleRef(st, s) {
  const { key, desc } = st.keys.keyOf(s);
  const styles = part(st, 'styles');
  if (!styles.has(key)) styles.set(key, desc);
  return key;
}
/** A line's attributes as the room keeps them: its format (rows' `s`, columns' `style`) as the format's key, `x`. */
function lineOut(st, attrs, axis) {
  const name = axis === 'row' ? 's' : 'style';
  if (!attrs || attrs[name] == null) return { ...attrs };
  const { [name]: s, ...rest } = attrs;
  return { ...rest, x: styleRef(st, Number(s)) };
}
function lineIn(st, attrs, axis) {
  if (!attrs || attrs.x == null) return { ...attrs };
  const { x, ...rest } = attrs;
  return { ...rest, [axis === 'row' ? 's' : 'style']: String(styleIn(st, x)) };
}

/** Whether this editor's workbook changed in a way the room doesn't hold piece by piece (then shared whole). */
export function unmodelled(st) {
  if (st.wb.sheets.map((s) => s.sheetId).join() !== st.list) return true;
  if (!samePrint(st.book, bookPrint(st.wb))) return true;
  for (let i = 0; i < st.wb.sheets.length; i++) {
    if (!st.wb._parsed.has(st.wb.sheets[i]) || !isGrid(st.wb, i)) continue;
    const was = st.prints.get(uidAt(st, i));
    if (was != null && was !== sheetPrint(st.wb, i)) return true;
  }
  return false;
}

/** Everything that differs between sheet i and its shadow, written into the room; the change, by line identity, for
 *  the history ({ type: 'state', uid, cells, rows, cols, format }). */
export function diffSheet(st, i) {
  const uid = uidAt(st, i);
  st.digesting = true;
  let now;
  try { now = digestSheet(st.wb, i); } finally { st.digesting = false; }
  const was = st.shadow.get(uid) || EMPTY;
  const data = dataOf(st, uid);
  const R = seqOf(st, uid, 'row');
  const C = seqOf(st, uid, 'col');
  const rec = { type: 'state', uid, cells: [], rows: [], cols: [], format: null };
  for (const key of new Set([...was.cells.keys(), ...now.cells.keys()])) {
    const a = was.cells.get(key);
    const b = now.cells.get(key);
    const [r, c] = key.split(',').map(Number);
    const id = `${R.tokAt(r)},${C.tokAt(c)}`;
    const ca = a ? a.c : { v: null };
    const cb = b ? b.c : { v: null };
    if (cb !== null && ca !== null && !sameContents(ca, cb)) {
      data.set(`v ${id}`, cb.f != null ? { f: cb.f, ...(cb.legacy ? { legacy: true } : {}), k: { ...st.known } } : { v: cb.v });
      rec.cells.push({ id, kind: 'v', before: ca, after: cb });
    }
    const sa = a?.s || 0;
    const sb = b?.s || 0;
    if (sa !== sb) {
      data.set(`s ${id}`, styleRef(st, sb));
      rec.cells.push({ id, kind: 's', before: sa, after: sb });
    }
  }
  for (const [axis, seq, list, k] of [['row', R, rec.rows, 'rows'], ['col', C, rec.cols, 'cols']]) {
    for (const n of new Set([...was[k].keys(), ...now[k].keys()])) {
      const a = was[k].get(n) || {};
      const b = now[k].get(n) || {};
      if (sameAttrs(a, b)) continue;
      const tok = seq.tokAt(n);
      data.set(`${axis === 'row' ? 'r' : 'c'} ${tok}`, lineOut(st, b, axis));
      list.push({ tok, before: a, after: b });
    }
  }
  if (!sameAttrs(was.format, now.format)) {
    data.set('fmt', { ...now.format });
    rec.format = { before: was.format, after: now.format };
  }
  st.shadow.set(uid, now);
  return rec.cells.length || rec.rows.length || rec.cols.length || rec.format ? rec : null;
}

/** A formula taken from the room made to read what it read when written: moved by the inserts and deletes its writer
 *  hadn't seen (`k`, the latest of each client's it had), as a formula already in the workbook was moved by them. */
function caughtUp(st, f, k, here) {
  for (const { op, forms } of st.order) {
    if (op.seq <= (k?.[op.cid] ?? 0)) continue;
    const at = indexOfUid(st, op.sheet);
    if (at < 0) continue;
    for (const { axis, at: pos, n } of forms) f = shiftRefs(f, { here, sheet: st.wb.sheets[at].name, axis, at: pos, n });
  }
  return f;
}

/** An insert or delete of lines (the room's or this editor's own replayed) made in the workbook, with the code the
 *  editor's own Insert and Delete use. */
export function applyOp(st, op) {
  if (!checkOp(op) || st.applied.has(op.id)) return;
  st.applied.add(op.id);
  st.known[op.cid] = Math.max(st.known[op.cid] || 0, op.seq);
  st.lamport = Math.max(st.lamport, op.ts);
  const seq = seqOf(st, op.sheet, op.axis);
  const si = indexOfUid(st, op.sheet);
  const forms = [];
  if (op.kind === 'ins') {
    const pos = seq.insert(op);
    if (pos >= 0) forms.push({ axis: op.axis, at: pos, n: op.n });
  } else {
    for (const { pos, n } of seq.remove(op.runs)) forms.push({ axis: op.axis, at: pos, n: -n });
  }
  st.order.push({ op, forms });
  if (si < 0 || !isGrid(st.wb, si)) return;
  for (const f of forms) {
    try { st.raw.changeLines(si, f.axis, f.at, f.n); } catch { /* (refused here, as Excel would: a table or an array in the way) */ }
  }
}

/** The index of the format the room has under `key`: this workbook's equal one or one added; the room's description
 *  checked first (anything else: the Normal format). */
function styleIn(st, key) {
  const desc = typeof key === 'string' ? part(st, 'styles').get(key) : null;
  return checkStyle(desc) ? st.keys.indexOf(key, desc) : 0;
}

/** The room's entries `keys` of sheet `uid` made in the workbook: contents, formats, lines' attributes, defaults.
 *  Returns what was changed: { cells: [{ r, c }] } with `lines` set when rows, columns or defaults were. */
export function applyEntries(st, uid, keys) {
  const out = { cells: [], lines: false };
  const si = indexOfUid(st, uid);
  if (si < 0 || !isGrid(st.wb, si)) return out;
  const data = dataOf(st, uid);
  const R = seqOf(st, uid, 'row');
  const C = seqOf(st, uid, 'col');
  const contents = [];
  const styles = [];
  const rows = [];
  const cols = [];
  const here = st.wb.sheets[si].name;
  const place = (id) => {
    const [rt, ct] = id.split(',');
    if (!isToken(rt) || !isToken(ct)) return null;
    const r = R.posOf(rt);
    const c = C.posOf(ct);
    return r?.alive && c?.alive ? { r: r.pos, c: c.pos } : null;
  };
  for (const key of keys) {
    const v = data.get(key);
    if (v === undefined) continue;
    const sp = key.indexOf(' ');
    const kind = sp < 0 ? key : key.slice(0, sp);
    const rest = key.slice(sp + 1);
    if (kind === 'v' || kind === 's') {
      const at = place(rest);
      if (!at) continue;
      if (kind === 's') styles.push({ ...at, s: styleIn(st, v) });
      else if (checkContents(v)) contents.push(v.f != null ? { ...at, formula: caughtUp(st, v.f, v.k, here), ...(v.legacy ? { legacy: true } : {}) } : { ...at, value: v.v ?? null });
      out.cells.push(at);
    } else if (kind === 'r' || kind === 'c') {
      const p = isToken(rest) ? (kind === 'r' ? R : C).posOf(rest) : null;
      const attrs = checkAttrs(v, kind === 'r' ? ROW_OWN : COL_OWN);
      if (!p?.alive || !attrs) continue;
      if (kind === 'r') rows.push({ r: p.pos, attrs: lineIn(st, attrs, 'row') }); else cols.push({ c: p.pos, attrs: lineIn(st, attrs, 'col') });
    } else if (kind === 'fmt') {
      const attrs = checkAttrs(v);
      if (attrs) writeFormat(st.wb, si, attrs);
      out.lines = true;
    }
  }
  writeRows(st.wb, si, rows);
  writeCols(st.wb, si, cols);
  if (rows.length || cols.length) out.lines = true;
  if (styles.length) st.wb.setStyles(si, styles);
  for (const ch of contents) {
    try { st.wb.apply(si, [ch]); } catch { /* (a cell another editor changed that is now part of an array here, or a formula that doesn't parse: left as it is) */ }
  }
  return out;
}

/** The room's sheet list and props made this workbook's: sheets added, deleted, moved, renamed, hidden, coloured. */
export function reconcileSheets(st) {
  const props = part(st, 'props');
  const seen = new Set();
  const want = [];
  for (const uid of part(st, 'sheets').toArray()) {
    if (seen.has(uid) || props.get(`${uid}/gone`)) continue;
    seen.add(uid);
    want.push(uid);
  }
  const wb = st.wb;
  const raw = st.raw;
  for (let i = wb.sheets.length - 1; i >= 0; i--) {
    if (want.includes(uidAt(st, i))) continue;
    try { raw.deleteSheet(i); } catch { /* (the last sheet shown: kept) */ }
  }
  want.forEach((uid, k) => {
    if (indexOfUid(st, uid) >= 0) return;
    const at = Math.min(k, wb.sheets.length);
    const given = props.get(`${uid}/name`);
    try { raw.addSheet(at, typeof given === 'string' ? given : null); } catch { raw.addSheet(at, null); }
    st.uids.set(wb.sheets[at].sheetId, uid);
  });
  want.forEach((uid, k) => {
    const i = indexOfUid(st, uid);
    if (i >= 0 && i !== k && k < wb.sheets.length) raw.moveSheet(i, k);
  });
  for (const uid of want) {
    const i = indexOfUid(st, uid);
    if (i < 0) continue;
    const name = props.get(`${uid}/name`);
    if (typeof name === 'string' && name !== wb.sheets[i].name) {
      // (Two editors giving two sheets one name at once: the second gets a number, as a copy's name does.)
      for (let n = 1; n < 100; n++) {
        try { raw.renameSheet(i, n === 1 ? name : `${name} (${n})`); break; } catch { /* the next number */ }
      }
    }
    const state = props.get(`${uid}/state`);
    if (['visible', 'hidden', 'veryHidden'].includes(state) && (state !== 'visible') !== (wb.sheets[i].state !== 'visible')) {
      try { raw.setSheetHidden(i, state !== 'visible'); } catch { /* (the last sheet shown stays shown) */ }
    }
    const color = props.get(`${uid}/color`);
    if (props.has(`${uid}/color`) && checkColor(color) && isGrid(wb, i) && JSON.stringify(color) !== JSON.stringify(tabColorOf(wb, i))) raw.setTabColor(i, color);
  }
}

/** Sheet i's tab colour as setTabColor takes it ({ rgb } or { theme, tint }), or null. */
export function tabColorOf(wb, i) {
  if (!isGrid(wb, i)) return null;
  const root = frameOf(wb._part(i)).documentElement;
  const pr = [...root.childNodes].find((n) => n.nodeType === 1 && n.localName === 'sheetPr');
  const el = pr && [...pr.childNodes].find((n) => n.nodeType === 1 && n.localName === 'tabColor');
  if (!el) return null;
  if (el.getAttribute('rgb')) return { rgb: el.getAttribute('rgb') };
  return { theme: Number(el.getAttribute('theme') || 0), ...(el.getAttribute('tint') ? { tint: Number(el.getAttribute('tint')) } : {}) };
}

/** The whole room made this workbook's, after its bytes were opened: sheets, then lines, then every entry. */
export function replayRoom(st) {
  reconcileSheets(st);
  for (const op of part(st, 'lines').toArray()) applyOp(st, op);
  for (const [uid, data] of part(st, 'cells').entries()) applyEntries(st, uid, [...data.keys()]);
}
