// Recalculation: which formulas a change reaches, the order to work them out in (each after the formulas it reads),
// and their results written back into the cells. Formulas Excel saved are trusted until something they read changes.
import { parseFormula } from './parse.js';
import { evaluate, result, pick, dynamicResult } from './evaluate.js';
import { place, release, spillOf, touched } from './spill.js';
import { Unsupported, E, cellValue } from './values.js';
import { FUNCTIONS } from './functions.js';
import { readTables, tableArea } from '../ooxml/tables.js';
import { runWithCircles } from './iterate.js';
import { cellInfo } from './cellinfo.js';
import { pivotData } from '../ooxml/pivotdata.js';
import { calcMode, compatVersion } from '../ooxml/calcmode.js';
import { contentsHidden } from '../ooxml/protection.js';

const key = (si, r, c) => `${si}:${r}:${c}`;
/** References wider than this many columns are checked against every change instead of being indexed by column. */
const WIDE = 64;
const overlaps = (a, b) => a.si === b.si && a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;
/** The cells a formula fills: an array formula's range, a dynamic formula's spill, or its own cell. */
const spanOf = (e) => e.array || e.spill || { r0: e.r, c0: e.c, r1: e.r, c1: e.c };

/** The calculation state of a workbook. `book` is the Workbook (workbook.js): its sheets' models, sheetIndexOf(name),
 *  definedName(name, si), date1904 and writeResult(si, r, c, value). */
export class Calc {
  constructor(book) {
    this.book = book;
    /** Formula cells by key: { si, r, c, text, tree, deps, volatile, array, unsupported }. */
    this.entries = null;
    /** Some formula can't be worked out here (unknown function, a link to another workbook): Excel should recalculate
     *  the workbook when it opens it. */
    this.incomplete = false;
    this._names = new Map();
    /** The workbook's tables, for table references (read at the first that needs them). */
    this._tables = null;
  }

  get tables() {
    if (!this._tables) this._tables = readTables(this.book);
    return this._tables;
  }

  /** Read every worksheet's formulas (done once, at the first change). */
  build() {
    if (this.entries) return;
    this.entries = new Map();
    this.byCol = new Map(); // si -> col -> Set of entries
    this.wide = new Map(); // si -> Set of entries
    this.at = new Map(); // key of every cell a formula fills (an array formula's whole range, a spill) -> entry
    this.spilledBy = new Map(); // key of a cell a dynamic formula spilled into (not its own) -> its entry (spill.js)
    this.blocked = new Set(); // dynamic formulas showing #SPILL!, waiting for room
    this.moved = []; // spill ranges that grew, shrank or went since the last pass: what reads them works out again
    this.book.sheets.forEach((info, si) => {
      if (info.kind !== 'worksheet' || !info.path) return;
      for (const [r, line] of this.book.sheet(si).cells) for (const [c, cell] of line) if (cell.f != null) this.add(si, r, c, cell);
    });
  }

  /** A defined name's formula as a tree (null if there is no such name or it doesn't parse). */
  nameTree(name, si) {
    const n = this.book.definedName(name, si);
    if (!n) return null;
    if (!this._names.has(n)) {
      let tree = null;
      try { tree = parseFormula(n.text); } catch { /* (kept as null: #NAME?) */ }
      this._names.set(n, tree);
    }
    return this._names.get(n);
  }

  /** What a formula reads, and whether it can be worked out here. */
  _collect(e, tree, si, seen = new Set()) {
    const walk = (n) => {
      if (!n) return;
      switch (n.k) {
        case 'ref': {
          if (n.book != null) { e.unsupported = true; return; }
          // (A 3D reference reads each sheet from its first to its last.)
          if (n.sheets) {
            const [s0, s1] = n.sheets.map((s) => this.book.sheetIndexOf(s));
            if (s0 >= 0 && s1 >= 0) for (let s = Math.min(s0, s1); s <= Math.max(s0, s1); s++) walk({ ...n, sheets: null, sheet: this.book.sheets[s].name });
            return;
          }
          const s = n.sheet == null ? si : this.book.sheetIndexOf(n.sheet);
          if (s >= 0) e.deps.push({ si: s, r0: Math.min(n.area.r0, n.area.r1), c0: Math.min(n.area.c0, n.area.c1), r1: Math.max(n.area.r0, n.area.r1), c1: Math.max(n.area.c0, n.area.c1) });
          return;
        }
        case 'struct': {
          const a = tableArea(this.tables, n.table, n.spec, { si, r: e.r, c: e.c });
          if (!a.error) e.deps.push({ si: a.si, r0: a.r0, c0: a.c0, r1: a.r1, c1: a.c1 });
          return;
        }
        case 'name': {
          const t = this.nameTree(n.name, si);
          if (t && !seen.has(n.name.toUpperCase())) { seen.add(n.name.toUpperCase()); walk(t); }
          return;
        }
        case 'fn':
          // (Not a function: a defined name holding a LAMBDA, called, =Twice(21); it reads what its formula reads.)
          if (!FUNCTIONS[n.name]) {
            const t = this.nameTree(n.name, si);
            if (!t) e.unsupported = true;
            else if (!seen.has(n.name.toUpperCase())) { seen.add(n.name.toUpperCase()); walk(t); }
          }
          else if (FUNCTIONS[n.name].volatile) e.volatile = true;
          n.args.forEach(walk);
          return;
        case 'bin': case 'isect': walk(n.a); walk(n.b); return;
        case 'range':
          walk(n.a);
          walk(n.b);
          // (The cells between its ends are read too: two references on one sheet give the range at once; one worked
          // out, as INDEX's, can be anywhere, so the formula is worked out at every change.)
          if (n.a.k === 'ref' && n.b.k === 'ref' && !n.a.sheets && !n.b.sheets && n.a.book == null && n.b.book == null) {
            const sa = n.a.sheet == null ? si : this.book.sheetIndexOf(n.a.sheet);
            const sb = n.b.sheet == null ? si : this.book.sheetIndexOf(n.b.sheet);
            const box = (x) => [Math.min(x.r0, x.r1), Math.min(x.c0, x.c1), Math.max(x.r0, x.r1), Math.max(x.c0, x.c1)];
            const [ar0, ac0, ar1, ac1] = box(n.a.area);
            const [br0, bc0, br1, bc1] = box(n.b.area);
            if (sa >= 0 && sa === sb) e.deps.push({ si: sa, r0: Math.min(ar0, br0), c0: Math.min(ac0, bc0), r1: Math.max(ar1, br1), c1: Math.max(ac1, bc1) });
          } else e.volatile = true;
          return;
        case 'neg': case 'pos': case 'pct': case 'paren': walk(n.a); return;
        case 'union': n.items.forEach(walk);
      }
    };
    walk(tree);
  }

  /** Take in a formula cell (from the file, or just entered). */
  add(si, r, c, cell) {
    if (!this.entries) return;
    this.remove(si, r, c);
    const e = { si, r, c, text: cell.f, tree: null, deps: [], volatile: false, unsupported: false, array: cell.array || null, dynamic: !!cell.dynamic, spill: null };
    if (cell.dataTable || cell.f === '') e.unsupported = true;
    else {
      try { e.tree = parseFormula(cell.f); } catch { e.unsupported = true; }
      if (e.tree) this._collect(e, e.tree, si);
    }
    if (e.unsupported) this.incomplete = true;
    this.entries.set(key(si, r, c), e);
    // (A dynamic formula from the file: the cells its saved spill fills are its, until it is worked out again.)
    if (e.dynamic && cell.spill && (cell.spill.r1 > r || cell.spill.c1 > c)) {
      e.spill = cell.spill;
      for (let rr = e.spill.r0; rr <= e.spill.r1; rr++) for (let cc = e.spill.c0; cc <= e.spill.c1; cc++) if (rr !== r || cc !== c) this.spilledBy.set(key(si, rr, cc), e);
    }
    const span = spanOf(e);
    for (let rr = span.r0; rr <= span.r1; rr++) for (let cc = span.c0; cc <= span.c1; cc++) this.at.set(key(si, rr, cc), e);
    for (const d of e.deps) {
      if (d.c1 - d.c0 >= WIDE) { if (!this.wide.has(d.si)) this.wide.set(d.si, new Set()); this.wide.get(d.si).add(e); continue; }
      if (!this.byCol.has(d.si)) this.byCol.set(d.si, new Map());
      const cols = this.byCol.get(d.si);
      for (let cc = d.c0; cc <= d.c1; cc++) { if (!cols.has(cc)) cols.set(cc, new Set()); cols.get(cc).add(e); }
    }
  }

  /** Forget a formula cell (it was cleared or given a value). */
  remove(si, r, c) {
    if (!this.entries) return;
    const e = this.entries.get(key(si, r, c));
    if (!e) return;
    this.entries.delete(key(si, r, c));
    const span = spanOf(e);
    for (let rr = span.r0; rr <= span.r1; rr++) for (let cc = span.c0; cc <= span.c1; cc++) if (this.at.get(key(si, rr, cc)) === e) this.at.delete(key(si, rr, cc));
    if (e.dynamic) release(this, e);
    for (const d of e.deps) {
      this.wide.get(d.si)?.delete(e);
      const cols = this.byCol.get(d.si);
      if (cols) for (let cc = d.c0; cc <= Math.min(d.c1, d.c0 + WIDE); cc++) cols.get(cc)?.delete(e);
    }
  }

  /** The formulas that read any cell of a rectangle { si, r0, c0, r1, c1 }. */
  *readers(rect) {
    const seen = new Set();
    const cols = this.byCol.get(rect.si);
    const lists = [this.wide.get(rect.si)];
    if (cols) for (let cc = rect.c0; cc <= rect.c1 && cc - rect.c0 < 16384; cc++) if (cols.has(cc)) lists.push(cols.get(cc));
    for (const list of lists) {
      if (!list) continue;
      for (const e of list) if (!seen.has(e) && (seen.add(e), e.deps.some((d) => overlaps(d, rect)))) yield e;
    }
  }

  /** The formulas among `dirty` that `e` reads. */
  _precedents(e, dirty) {
    const out = new Set();
    for (const d of e.deps) {
      const size = (d.r1 - d.r0 + 1) * (d.c1 - d.c0 + 1);
      if (size <= 4096) {
        for (let r = d.r0; r <= d.r1; r++) for (let c = d.c0; c <= d.c1; c++) { const p = this.at.get(key(d.si, r, c)); if (p && p !== e && dirty.has(p)) out.add(p); }
      } else {
        for (const p of dirty) if (p !== e && p.si === d.si && overlaps(d, { si: p.si, ...spanOf(p) })) out.add(p);
      }
    }
    return out;
  }

  /** Work out again every formula that the changed rectangles reach, and the volatile ones; `also`: formula cells
   *  to work out in any case (just entered). `only`: those alone (Manual calculation, calcmode.js); `sheet`: of what
   *  the rest reach, sheet `sheet`'s formulas alone (Calculate Sheet). Returns { circular: [entries] }. */
  recalc(changed, also = [], { only = false, sheet = null } = {}) {
    this.build();
    const dirty = new Set();
    const queue = [...changed];
    const take = (e) => {
      if (dirty.has(e)) return;
      dirty.add(e);
      queue.push({ si: e.si, ...spanOf(e) });
    };
    for (const a of also) { const e = this.entries.get(key(a.si, a.r, a.c)); if (e) take(e); }
    if (only) return this._settle(this._run(dirty));
    for (const e of touched(this, changed)) take(e);
    for (const e of this.entries.values()) if (e.volatile) take(e);
    // (With iterative calculation on, Excel goes round the circles again at every recalculation.)
    for (const e of this.cyclic || []) if (this.entries.get(key(e.si, e.r, e.c)) === e) take(e);
    while (queue.length) for (const e of this.readers(queue.pop())) take(e);
    if (sheet != null) for (const e of dirty) if (e.si !== sheet) dirty.delete(e);
    return this._settle(this._run(dirty));
  }

  /** Work out every formula in the workbook. */
  calculateAll() {
    this.build();
    return this._settle(this._run(new Set(this.entries.values())));
  }

  /** After a pass: while spills grew, shrank or went, what reads the cells they took or left is worked out again
   *  (a few rounds at most: a spill that reads what it spills over stops there). */
  _settle(out) {
    for (let round = 0; round < 8 && this.moved.length; round++) {
      const rects = this.moved.splice(0);
      const dirty = new Set();
      for (const a of rects) for (const e of this.readers(a)) dirty.add(e);
      if (!dirty.size) break;
      const more = this._run(dirty);
      out.circular.push(...more.circular);
    }
    this.moved.length = 0;
    return out;
  }

  _run(dirty) {
    // Depth first, each formula after those it reads (with an explicit stack: a column of formulas each reading the one
    // above is as deep as it is long).
    const state = new Map();
    const order = [];
    const circular = new Set();
    for (const start of dirty) {
      if (state.has(start)) continue;
      state.set(start, 1);
      const stack = [[start, this._precedents(start, dirty).values()]];
      while (stack.length) {
        const top = stack[stack.length - 1];
        const next = top[1].next();
        if (next.done) { state.set(top[0], 2); order.push(top[0]); stack.pop(); continue; }
        const p = next.value;
        if (state.get(p) === 1) { for (const [x] of stack.slice(stack.findIndex(([y]) => y === p))) circular.add(x); continue; }
        if (state.get(p) === 2) continue;
        state.set(p, 1);
        stack.push([p, this._precedents(p, dirty).values()]);
      }
    }
    // (A formula reading its own cell, =A1+1 in A1 or =SUM(A:A) there, is a circle of one.)
    for (const e of order) if (e.deps.some((d) => d.si === e.si && e.r >= d.r0 && e.r <= d.r1 && e.c >= d.c0 && e.c <= d.c1)) circular.add(e);
    // (Circles: left out, or gone round with iterative calculation on: iterate.js.)
    return { circular: runWithCircles(this, order, circular, dirty) };
  }

  /** A formula (its text as the file has it, without "=") worked out as if it were in cell (r, c) of sheet si, without
   *  being entered there (a conditional format's): its value, or null when it can't be worked out here. With `array`,
   *  a range or array as a whole (a validation list's cells). */
  formulaAt(si, r, c, text, array = false) {
    try {
      return result(parseFormula(text), this.context({ si, r, c }), array);
    } catch (err) {
      if (err instanceof Unsupported || err instanceof SyntaxError || /formula/i.test(String(err?.message))) return null;
      throw err;
    }
  }

  /** The context a formula is worked out in (evaluate.js). */
  context(e) {
    const book = this.book;
    const ctx = {
      here: { si: e.si, r: e.r, c: e.c },
      date1904: book.date1904,
      // (The text functions count code points from compatibility version 2 on, UTF-16 units before: calcmode.js.)
      codePoints: !!book.doc && compatVersion(book) >= 2,
      volatile: false,
      cellValue: (si, r, c) => cellValue(book.sheet(si).cells.get(r)?.get(c)),
      spillOf: (si, r, c) => spillOf(this, si, r, c),
      rowsOf: (si) => book.sheet(si).cells,
      extent: (si) => { const m = book.sheet(si); return { maxRow: m.maxRow, maxCol: m.maxCol }; },
      sheetIndex: (name) => book.sheetIndexOf(name),
      sheetCount: () => book.sheets.length,
      sheetName: (si) => book.sheets[si]?.name,
      // (CELL and INFO: functions-info.js.)
      cellInfo: (si, r, c) => cellInfo(book, si, r, c),
      pivotData: (area, name, pairs) => pivotData(book, area, name, pairs),
      bookName: () => book.fileName || null,
      calcMode: () => calcMode(book),
      tableArea: (table, spec) => tableArea(this.tables, table, spec, ctx.here),
      // (For SUBTOTAL: a cell's own formula, and whether its row is hidden.)
      formulaOf: (si, r, c) => book.sheet(si).cells.get(r)?.get(c)?.f || null,
      rowHidden: (si, r) => !!book.sheet(si).rows.get(r)?.hidden,
      // (FORMULATEXT: none from a hidden cell of a protected sheet.)
      formulaShown: (si, r, c) => (book._part && contentsHidden(book, si, r, c) ? null : book.sheet(si).cells.get(r)?.get(c)?.f || null),
      name: (name, mode) => {
        const t = this.nameTree(name, e.si);
        return t ? evaluate(t, ctx, mode) : E.name;
      },
      hasName: (name) => !!this.nameTree(name, e.si),
    };
    return ctx;
  }

  _evaluate(e) {
    if (!e.tree) return;
    const ctx = this.context(e);
    let v;
    try {
      // (A dynamic array formula: its result spilled from its cell, or one value; spill.js.)
      if (e.dynamic && !e.array) { place(this, e, dynamicResult(e.tree, ctx)); return; }
      v = result(e.tree, ctx, !!e.array);
    } catch (err) {
      if (!(err instanceof Unsupported)) throw err;
      e.unsupported = true;
      this.incomplete = true;
      // (A formula just entered has no saved value to keep: Excel will show what it makes of it.)
      if (cellValue(this.book.sheet(e.si).cells.get(e.r)?.get(e.c)) === null) this.book.writeResult(e.si, e.r, e.c, E.name);
      return;
    }
    if (!e.array) { this.book.writeResult(e.si, e.r, e.c, v); return; }
    const a = e.array;
    for (let r = a.r0; r <= a.r1; r++) for (let c = a.c0; c <= a.c1; c++) this.book.writeResult(e.si, r, c, pick(v, r - a.r0, c - a.c0));
  }
}
