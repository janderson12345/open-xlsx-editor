// This editor's own changes to a co-edited workbook, caught for the room (collab.js): the workbook's methods that
// insert and delete lines and change its sheets are wrapped, so each goes to the room as what it is (an insert of 3
// rows after row '12', a sheet renamed) and not as its effects; every other change is found by comparing the sheets
// with their shadow after each edit (collab-room.js diffSheet).
//
// Undo and Redo of a change that was a whole-workbook snapshot (rows inserted, sizes, sheets) would put back this
// editor's workbook as it was, and with it whatever others changed since. Here the snapshot's step is taken back
// piece by piece instead, from a journal of what was shared: the lines an insert made deleted again, a deleted line's
// cells written back, each cell's and line's value before. Others' changes since stay.
import { refreshAll, refreshSheet, uidAt, indexOfUid, seqOf, dataOf, diffSheet, unmodelled, tabColorOf, isGrid } from './collab-room.js';
import { contentsOf, sameContents, writeRows, writeCols, writeFormat } from './collab-state.js';
import { tokenOf } from './collab-seq.js';

export const WRAPPED = ['changeLines', 'addSheet', 'renameSheet', 'moveSheet', 'setSheetHidden', 'setTabColor', 'deleteSheet', 'snapshot', 'restore', '_part'];

/** Everything this editor changed since its last share, shared: piece by piece, or (a change the room doesn't hold
 *  that way) the whole workbook. */
export function flush(st) {
  if (st.quiet || !st.room || st.publishing) return;
  if (unmodelled(st)) { st.journal.push({ type: 'publish' }); st.host.publish(); return; }
  const recs = [];
  st.ydoc.transact(() => {
    for (let i = 0; i < st.wb.sheets.length; i++) {
      if (!st.wb._parsed.has(st.wb.sheets[i]) || !isGrid(st.wb, i)) continue;
      const rec = diffSheet(st, i);
      if (rec) recs.push(rec);
    }
  }, st.origin);
  st.journal.push(...recs);
}

/** Cell contents typed ([{ si, contents: [{ r, c }] }], api.onEdit's cells): those cells alone compared and shared. */
export function flushCells(st, list) {
  if (st.quiet || !st.room) return;
  const recs = [];
  st.ydoc.transact(() => {
    for (const { si, contents } of list) {
      if (!isGrid(st.wb, si)) continue;
      const uid = uidAt(st, si);
      const shadow = st.shadow.get(uid);
      if (!shadow) { const rec = diffSheet(st, si); if (rec) recs.push(rec); continue; }
      const { model } = st.wb._part(si);
      const R = seqOf(st, uid, 'row');
      const C = seqOf(st, uid, 'col');
      const rec = { type: 'state', uid, cells: [], rows: [], cols: [], format: null };
      for (const { r, c } of contents) {
        const key = `${r},${c}`;
        const cell = model.cells.get(r)?.get(c);
        const now = contentsOf(st.wb, si, model, r, c, cell);
        const was = shadow.cells.get(key);
        const before = was ? was.c : { v: null };
        if (now === null || before === null || sameContents(before, now)) continue;
        const id = `${R.tokAt(r)},${C.tokAt(c)}`;
        dataOf(st, uid).set(`v ${id}`, now.f != null ? { f: now.f, ...(now.legacy ? { legacy: true } : {}), k: { ...st.known } } : { v: now.v });
        rec.cells.push({ id, kind: 'v', before, after: now });
        if ((now.v == null && now.f == null) && !(cell?.s)) shadow.cells.delete(key);
        else shadow.cells.set(key, { c: now, s: cell?.s || 0 });
      }
      if (rec.cells.length) recs.push(rec);
    }
  }, st.origin);
  st.journal.push(...recs);
}

/** The whole workbook shared once the change being made is done (a microtask: one command can make several). */
const publishSoon = (st) => {
  st.journal.push({ type: 'publish' });
  if (st.publishPending) return;
  st.publishPending = true;
  queueMicrotask(() => { st.publishPending = false; st.host.publish(); });
};

/** An insert or delete of lines made here (`forms`, as changeLines took them) recorded and put in the room. */
function commit(st, op, forms, step) {
  st.lamport = op.ts;
  st.applied.add(op.id);
  st.known[st.cid] = op.seq;
  st.order.push({ op, forms });
  st.ydoc.transact(() => st.room.get('lines').push([op]), st.origin);
  st.journal.push({ type: 'op', uid: op.sheet, axis: op.axis, kind: op.kind, ...step });
  st.lastOp = op;
  refreshAll(st);
}

/** What a delete of `count` lines from `at` takes away, from the shadow: each cell ({ d, o, c, s }: its line's offset
 *  in the delete, its place along the line, contents, format) and each line's attributes ({ d, attrs }). */
function lostLines(st, uid, axis, at, count) {
  const sh = st.shadow.get(uid);
  if (!sh) return { cells: [], lines: [] };
  const cells = [];
  for (const [key, v] of sh.cells) {
    const [r, c] = key.split(',').map(Number);
    const [d, o] = axis === 'row' ? [r - at, c] : [c - at, r];
    if (d >= 0 && d < count) cells.push({ d, o, c: v.c, s: v.s });
  }
  const lines = [];
  for (const [n, attrs] of sh[axis === 'row' ? 'rows' : 'cols']) if (n >= at && n < at + count) lines.push({ d: n - at, attrs });
  return { cells, lines };
}

/** The capturing versions of the workbook's methods, for `st.wb`; `st.raw` keeps the originals. */
export function captureOps(st) {
  const raw = st.raw;
  const nextId = () => `${st.cid}.${++st.seqNo}`;
  const cap = {
    changeLines(i, axis, at, n, band = null, opts = {}) {
      if (band) {
        // (Cells shifted within some columns or rows: positions, not lines, move; the whole workbook is shared. The
        // editor's Insert and Delete Cells call rowcol.js directly, and go as the cells that moved: docs/COLLAB.md.)
        raw.changeLines(i, axis, at, n, band, opts);
        publishSoon(st);
        return;
      }
      flush(st);
      const uid = uidAt(st, i);
      const seq = seqOf(st, uid, axis);
      const anchor = at > 0 ? seq.tokAt(at - 1) : null;
      const op = { id: nextId(), cid: st.cid, seq: st.seqNo, ts: st.lamport + 1, sheet: uid, axis };
      if (n > 0) Object.assign(op, { kind: 'ins', after: anchor, n });
      else Object.assign(op, { kind: 'del', runs: seq.span(at, -n) });
      const lost = n < 0 ? lostLines(st, uid, axis, at, -n) : null;
      raw.changeLines(i, axis, at, n, band, opts);
      if (n > 0) seq.insert(op); else seq.remove(op.runs);
      commit(st, op, [{ axis, at, n }], { anchor, lost, n: Math.abs(n), runs: n > 0 ? [{ id: op.id, start: 0, len: n }] : op.runs });
    },
    /** Lines `runs` (by identity) of sheet i deleted, wherever they are now: an insert's Undo. */
    removeRuns(i, axis, runs) {
      flush(st);
      const uid = uidAt(st, i);
      const seq = seqOf(st, uid, axis);
      const op = { id: nextId(), cid: st.cid, seq: st.seqNo, ts: st.lamport + 1, sheet: uid, axis, kind: 'del', runs };
      const forms = seq.remove(runs).map(({ pos, n }) => ({ axis, at: pos, n: -n }));
      for (const f of forms) raw.changeLines(i, axis, f.at, f.n);
      commit(st, op, forms, { anchor: null, lost: null, n: forms.reduce((t, f) => t - f.n, 0), runs });
    },
    addSheet(at = st.wb.sheets.length, name = null) {
      flush(st);
      const k = raw.addSheet(at, name);
      const uid = nextId();
      st.uids.set(st.wb.sheets[k].sheetId, uid);
      st.ydoc.transact(() => {
        dataOf(st, uid);
        placeUid(st, uid, k);
        const props = st.room.get('props');
        props.set(`${uid}/name`, st.wb.sheets[k].name);
        props.set(`${uid}/state`, st.wb.sheets[k].state);
      }, st.origin);
      st.journal.push({ type: 'sheet', kind: 'add', uid });
      refreshAll(st);
      return k;
    },
    deleteSheet(i) {
      flush(st);
      const uid = uidAt(st, i);
      raw.deleteSheet(i);
      st.ydoc.transact(() => {
        st.room.get('props').set(`${uid}/gone`, true);
        removeUid(st, uid);
      }, st.origin);
      // (A deleted sheet's cells and parts can't be put back piece by piece: its Undo shares the whole workbook.)
      st.journal.push({ type: 'sheet', kind: 'del', uid });
      refreshAll(st);
    },
    renameSheet(i, name) {
      flush(st);
      const before = st.wb.sheets[i].name;
      raw.renameSheet(i, name);
      sheetProp(st, uidAt(st, i), 'name', name, { kind: 'rename', before, after: name });
    },
    moveSheet(from, to) {
      flush(st);
      const uid = uidAt(st, from);
      raw.moveSheet(from, to);
      st.ydoc.transact(() => { removeUid(st, uid); placeUid(st, uid, to); }, st.origin);
      st.journal.push({ type: 'sheet', kind: 'move', uid, from, to });
      refreshAll(st);
    },
    setSheetHidden(i, hidden) {
      flush(st);
      const before = st.wb.sheets[i].state;
      raw.setSheetHidden(i, hidden);
      sheetProp(st, uidAt(st, i), 'state', st.wb.sheets[i].state, { kind: 'state', before, after: st.wb.sheets[i].state });
    },
    setTabColor(i, color) {
      flush(st);
      const before = tabColorOf(st.wb, i);
      raw.setTabColor(i, color);
      sheetProp(st, uidAt(st, i), 'color', tabColorOf(st.wb, i), { kind: 'color', before, after: tabColorOf(st.wb, i) });
    },
    snapshot() {
      flush(st);
      const snap = raw.snapshot();
      // (A command's before and after snapshots are taken in one go: paired, the journal between them is its step.)
      const meta = { j: st.journal.length, tick: st.tick, prev: null, next: null };
      const last = st.lastSnap;
      if (last && last.tick === st.tick && !last.next && !last.prev) { last.next = meta; meta.prev = last; }
      st.snapMeta.set(snap, meta);
      st.lastSnap = meta;
      if (!st.tickPending) { st.tickPending = true; queueMicrotask(() => { st.tick++; st.tickPending = false; }); }
      return snap;
    },
    restore(snap) {
      flush(st);
      const meta = st.snapMeta.get(snap);
      const steps = meta?.next ? st.journal.slice(meta.j, meta.next.j) : meta?.prev ? st.journal.slice(meta.prev.j, meta.j) : null;
      const start = meta?.next ? meta.j : meta?.prev ? meta.prev.j : -1;
      if (steps && start >= st.since && steps.every(invertible)) {
        if (meta.next) for (const s of [...steps].reverse()) undoStep(st, cap, s);
        else for (const s of steps) redoStep(st, cap, s);
        return;
      }
      raw.restore(snap);
      // (A command's own rollback with nothing shared since: the shadow taken again. Otherwise the whole workbook.)
      if (meta && !meta.next && !meta.prev && st.journal.length === meta.j) refreshAll(st);
      else publishSoon(st);
    },
  };
  return cap;
}

const invertible = (s) => s.type === 'state' || s.type === 'op' || (s.type === 'sheet' && s.kind !== 'del');

/** Sheet `uid`'s room entry in the sheet list placed where index k of this workbook's sheets is (after the sheet
 *  before it there, or first). */
function placeUid(st, uid, k) {
  const arr = st.room.get('sheets');
  const prev = k > 0 ? uidAt(st, k - 1) : null;
  const list = arr.toArray();
  arr.insert(prev ? list.indexOf(prev) + 1 : 0, [uid]);
}
function removeUid(st, uid) {
  const arr = st.room.get('sheets');
  for (let i = arr.length - 1; i >= 0; i--) if (arr.get(i) === uid) arr.delete(i, 1);
}
function sheetProp(st, uid, key, value, step) {
  st.ydoc.transact(() => st.room.get('props').set(`${uid}/${key}`, value), st.origin);
  st.journal.push({ type: 'sheet', uid, ...step });
  refreshAll(st);
}

/** A step's cells, lines and defaults written back as they were (`side` 'before') or again ('after'). */
function writeState(st, s, side) {
  const si = indexOfUid(st, s.uid);
  if (si < 0) return;
  const R = seqOf(st, s.uid, 'row');
  const C = seqOf(st, s.uid, 'col');
  const contents = [];
  const styles = [];
  for (const x of s.cells) {
    const [rt, ct] = x.id.split(',');
    const r = R.posOf(rt);
    const c = C.posOf(ct);
    if (!r?.alive || !c?.alive) continue;
    const v = x[side];
    if (x.kind === 's') styles.push({ r: r.pos, c: c.pos, s: v });
    else contents.push(v.f != null ? { r: r.pos, c: c.pos, formula: v.f, ...(v.legacy ? { legacy: true } : {}) } : { r: r.pos, c: c.pos, value: v.v ?? null });
  }
  const lines = (list, seq) => list.map((x) => ({ p: seq.posOf(x.tok), attrs: x[side] })).filter((x) => x.p?.alive);
  writeRows(st.wb, si, lines(s.rows, R).map(({ p, attrs }) => ({ r: p.pos, attrs })));
  writeCols(st.wb, si, lines(s.cols, C).map(({ p, attrs }) => ({ c: p.pos, attrs })));
  if (s.format) writeFormat(st.wb, si, s.format[side]);
  if (styles.length) st.wb.setStyles(si, styles);
  for (const ch of contents) { try { st.wb.apply(si, [ch]); } catch { /* (now part of an array: left) */ } }
}

/** An op step's lines taken away again, by identity: the ones it made (an insert's) or put back (a delete's Undo). */
function removeLines(st, cap, s) {
  const si = indexOfUid(st, s.uid);
  if (si < 0) return;
  const seq = seqOf(st, s.uid, s.axis);
  const first = s.runs.map((r) => seq.posOf(tokenOf(r.id, r.start))).find((p) => p?.alive);
  s.anchor = first && first.pos > 0 ? seq.tokAt(first.pos - 1) : null;
  cap.removeRuns(si, s.axis, s.runs);
  st.journal.pop();
}

/** An op step's lines made again after its anchor: an insert's Redo, or a delete's Undo with what it took away. */
function insertLines(st, cap, s, lost) {
  const si = indexOfUid(st, s.uid);
  if (si < 0) return;
  const seq = seqOf(st, s.uid, s.axis);
  const a = s.anchor == null ? null : seq.posOf(s.anchor);
  const at = a == null ? 0 : a.alive ? a.pos + 1 : a.pos;
  cap.changeLines(si, s.axis, at, s.n);
  st.journal.pop();
  s.runs = [{ id: st.lastOp.id, start: 0, len: s.n }];
  if (!lost) return;
  const rows = s.axis === 'row';
  const contents = lost.cells.filter((x) => x.c && (x.c.f != null || x.c.v != null)).map((x) => {
    const [r, c] = rows ? [at + x.d, x.o] : [x.o, at + x.d];
    return x.c.f != null ? { r, c, formula: x.c.f, ...(x.c.legacy ? { legacy: true } : {}) } : { r, c, value: x.c.v };
  });
  const styles = lost.cells.filter((x) => x.s).map((x) => (rows ? { r: at + x.d, c: x.o, s: x.s } : { r: x.o, c: at + x.d, s: x.s }));
  if (rows) writeRows(st.wb, si, lost.lines.map((x) => ({ r: at + x.d, attrs: x.attrs })));
  else writeCols(st.wb, si, lost.lines.map((x) => ({ c: at + x.d, attrs: x.attrs })));
  if (styles.length) st.wb.setStyles(si, styles);
  for (const ch of contents) { try { st.wb.apply(si, [ch]); } catch { /* (now part of an array: left) */ } }
}

function undoStep(st, cap, s) {
  if (s.type === 'state') writeState(st, s, 'before');
  else if (s.type === 'op') { if (s.kind === 'ins') removeLines(st, cap, s); else insertLines(st, cap, s, s.lost); }
  else sheetStep(st, cap, s, 'before');
}
function redoStep(st, cap, s) {
  if (s.type === 'state') writeState(st, s, 'after');
  else if (s.type === 'op') { if (s.kind === 'ins') insertLines(st, cap, s, null); else removeLines(st, cap, s); }
  else sheetStep(st, cap, s, 'after');
}

/** A sheet step taken back or made again with the capturing methods (so it reaches the room); their own journal
 *  entries dropped, the step keeping its place. */
function sheetStep(st, cap, s, side) {
  const i = indexOfUid(st, s.uid);
  const n = st.journal.length;
  if (s.kind === 'add') {
    if (side === 'before' && i >= 0) cap.deleteSheet(i);
    else if (side === 'after') { const k = cap.addSheet(s.at ?? st.wb.sheets.length, s.name ?? null); s.uid = uidAt(st, k); }
    if (side === 'before') { s.at = i; s.name = st.wb.sheets[i]?.name ?? null; }
  } else if (i >= 0) {
    if (s.kind === 'rename') cap.renameSheet(i, s[side]);
    else if (s.kind === 'state') cap.setSheetHidden(i, s[side] !== 'visible');
    else if (s.kind === 'color') cap.setTabColor(i, s[side]);
    else if (s.kind === 'move') cap.moveSheet(i, side === 'before' ? s.from : s.to);
  }
  st.journal.length = n;
}

/** The capturing methods put on `st.wb` (each passes straight through while the room's changes are made here, or
 *  inside another), and the originals kept as `st.raw`. */
export function wrapWorkbook(st) {
  const wb = st.wb;
  st.raw = Object.fromEntries(WRAPPED.map((k) => [k, wb[k].bind(wb)]));
  const cap = captureOps(st);
  for (const k of WRAPPED) {
    if (k === '_part') continue;
    wb[k] = (...args) => {
      if (st.quiet || st.depth || !st.room) return st.raw[k](...args);
      st.depth++;
      try { return cap[k](...args); } finally { st.depth--; }
    };
  }
  // (A sheet read for the first time: its shadow taken before anything changes it.)
  wb._part = (i) => {
    const had = wb._parsed.has(wb.sheets[i]);
    const p = st.raw._part(i);
    if (!had && !st.digesting && st.room && uidAt(st, i) != null && isGrid(wb, i)) refreshSheet(st, i);
    return p;
  };
  st.cap = cap;
}

/** The workbook's own methods back. */
export function unwrapWorkbook(wb) {
  for (const k of WRAPPED) delete wb[k];
}
