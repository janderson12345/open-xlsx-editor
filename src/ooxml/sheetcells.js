// A sheet model's cells, and what both of its readers (sheet.js from the parsed XML, sheetscan.js from the part's
// bytes) make them with. `SheetCells` is the Map of rows the model has always had, with one addition: rows that hold
// nothing but formatted empty cells (`<c r="A5" s="9"/>`) are kept as bands, { r0, r1, runs }, a run being columns
// c0 to c1 with one format, and consecutive rows with the same runs one band. A sheet formatted down to its last row
// (a customer's: A1:J1046613, ten million such cells around a few thousand values) is then a handful of bands, not
// ten million cell objects. Asked for such a row (get), the Map makes it, as the cells it would have had; walked as
// a whole (for...of, keys, size's rows are counted without it), it makes them all. What only wants the cells that
// hold something walks `filled()` instead, and what only wants a format asks `blankStyle`.
import { shiftFormula } from '../calc/tokens.js';
import { parseRange } from './refs.js';

/** A cell as stored: `t` its type (n number, s string, b boolean, e error, str formula text, d ISO date), `v` its
 *  value (null for a cell that only carries a style), `s` its style (an index into cellXfs), `f` its formula text.
 *  The readers add, where they apply: `shared` (the index of the shared formula the cell belongs to; `f` is then its
 *  own text, worked out from the group's first cell), `array` (an array formula's range, on its first cell) and
 *  `dataTable`. */
export const makeCell = (t, v, s, f) => ({ t, v, s, f });
/** A text cell's rich-text runs (strings.js stringItemRuns), kept as `runs` for drawing it. */
export const withRuns = (cell, runs) => (runs ? Object.assign(cell, { runs }) : cell);

/** An outline level as Excel allows it, 0 to 7 (a file's larger or negative one kept in range: outline.js MAX_LEVEL). */
export const outlineLevel = (v) => Math.max(0, Math.min(7, Math.trunc(Number(v)) || 0));

/** A row's entry in the model's `rows` from its attributes' values (null: absent), or null when it says nothing:
 *  { ht (points), hidden, customHeight, s (its own format: `s` with customFormat, as Excel writes a formatted row),
 *  level and collapsed (its outline, outline.js), thickBot and thickTop (a thick bottom or top border in the row,
 *  for each of which Excel makes the row 0.75pt taller; rowheights.js) }. */
export function rowEntry(ht, hiddenAt, customFormat, sAt, levelAt, collapsedAt, thickBotAt, customHeight, thickTopAt = null) {
  const on = (v) => v === '1' || v === 'true';
  const hidden = on(hiddenAt);
  const own = on(customFormat) ? Number(sAt || 0) : 0;
  const level = outlineLevel(levelAt);
  const collapsed = on(collapsedAt);
  const thickBot = on(thickBotAt);
  const thickTop = on(thickTopAt);
  if (!(ht != null || hidden || own || level || collapsed || thickBot || thickTop)) return null;
  return { ht: ht == null ? null : Number(ht), hidden, customHeight: customHeight === '1', ...(own ? { s: own } : {}), ...(level ? { level } : {}), ...(collapsed ? { collapsed } : {}), ...(thickBot ? { thickBot } : {}), ...(thickTop ? { thickTop } : {}) };
}

/** What kind of formula a cell holds, from its `f` element's attributes (t, si, ref; null: absent). A shared formula
 *  is written out on its group's first cell only (the one with `ref`); the others read it as moved to where they are. */
export function formulaKind(cell, t, si, ref, r, c, shared, arrays, dynamic = false) {
  if (t === 'shared') {
    cell.shared = Number(si);
    if (cell.f !== '') shared.set(si, { r, c, text: cell.f });
    else if (shared.has(si)) {
      const first = shared.get(si);
      try { cell.f = shiftFormula(first.text, r - first.r, c - first.c); } catch { /* (left '': not worked out) */ }
    }
  } else if (t === 'array' && dynamic) {
    // (Excel 365's dynamic array formula, its cell metadata marking it: the range it spilled into, not a fixed array.)
    cell.dynamic = true;
    cell.spill = parseRange(ref) || { r0: r, c0: c, r1: r, c1: c };
  } else if (t === 'array') {
    cell.array = parseRange(ref) || { r0: r, c0: c, r1: r, c1: c };
    arrays.push(cell.array);
  } else if (t === 'dataTable') cell.dataTable = true;
}

/** A row of formatted empty cells from its runs [c0, c1, s, ...]: the Map the reader would have made of it. */
function lineOf(runs, n = runs.length) {
  const line = new Map();
  for (let k = 0; k < n; k += 3) {
    const s = runs[k + 2];
    for (let c = runs[k], c1 = runs[k + 1]; c <= c1; c++) line.set(c, { t: 'n', v: null, s, f: null });
  }
  return line;
}
export { lineOf as blankLine };

export class SheetCells extends Map {
  constructor() {
    super();
    /** Rows of formatted empty cells only, in row order, none of them in the Map itself: [{ r0, r1, runs }]. */
    this.bands = [];
  }

  /** The index in `bands` of the band row r is in, or -1. */
  _bandAt(r) {
    const b = this.bands;
    let lo = 0;
    let hi = b.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (r < b[mid].r0) hi = mid - 1;
      else if (r > b[mid].r1) lo = mid + 1;
      else return mid;
    }
    return -1;
  }

  /** Row r taken out of its band (index k); its runs back. */
  _cut(k, r) {
    const band = this.bands[k];
    if (band.r0 === band.r1) this.bands.splice(k, 1);
    else if (r === band.r0) band.r0++;
    else if (r === band.r1) band.r1--;
    else {
      this.bands.splice(k + 1, 0, { r0: r + 1, r1: band.r1, runs: band.runs });
      band.r1 = r - 1;
    }
    return band.runs;
  }

  /** A row of formatted empty cells added by a reader (rows come in order): `runs` [c0, c1, s, ...], `n` of its
   *  entries used. */
  addBlankRow(r, runs, n) {
    const last = this.bands[this.bands.length - 1];
    if (last && last.r1 === r - 1 && last.runs.length === n) {
      let k = 0;
      while (k < n && last.runs[k] === runs[k]) k++;
      if (k === n) { last.r1 = r; return; }
    }
    this.bands.push({ r0: r, r1: r, runs: runs.slice(0, n) });
  }

  /** How many formatted empty cells the bands stand for. */
  get bandCells() {
    let n = 0;
    for (const { r0, r1, runs } of this.bands) {
      let w = 0;
      for (let k = 0; k < runs.length; k += 3) w += runs[k + 1] - runs[k] + 1;
      n += w * (r1 - r0 + 1);
    }
    return n;
  }

  /** The format of the formatted empty cell (r, c) of a band, or -1 when (r, c) isn't one (a row of the Map itself,
   *  or no cell there): what drawing a row needs of it, without making the row. */
  blankStyle(r, c) {
    const k = this._bandAt(r);
    if (k < 0) return -1;
    const runs = this.bands[k].runs;
    for (let j = 0; j < runs.length; j += 3) if (c >= runs[j] && c <= runs[j + 1]) return runs[j + 2];
    return -1;
  }

  /** The runs [c0, c1, s, ...] of row r when it is a band's, else null. */
  blankRuns(r) {
    const k = this._bandAt(r);
    return k < 0 ? null : this.bands[k].runs;
  }

  /** The rows that may hold more than formatted empty cells, [r, line], without making the bands' rows: what
   *  looks for values, formulas, text or rich text walks these. */
  filled() { return super.entries(); }

  /** Every band's rows made (the Map then has every row, in row order, and no bands). */
  expand() {
    if (!this.bands.length) return this;
    const own = [...super.entries()].sort((a, b) => a[0] - b[0]);
    const bands = this.bands;
    this.bands = [];
    super.clear();
    let k = 0;
    for (const { r0, r1, runs } of bands) {
      while (k < own.length && own[k][0] < r0) { super.set(own[k][0], own[k][1]); k++; }
      for (let r = r0; r <= r1; r++) super.set(r, lineOf(runs));
    }
    while (k < own.length) { super.set(own[k][0], own[k][1]); k++; }
    return this;
  }

  get(r) {
    const line = super.get(r);
    if (line !== undefined || !this.bands.length) return line;
    const k = this._bandAt(r);
    if (k < 0) return undefined;
    const made = lineOf(this._cut(k, r));
    super.set(r, made);
    return made;
  }

  has(r) { return super.has(r) || (this.bands.length > 0 && this._bandAt(r) >= 0); }

  set(r, line) {
    if (this.bands && this.bands.length) {
      const k = this._bandAt(r);
      if (k >= 0) this._cut(k, r);
    }
    return super.set(r, line);
  }

  delete(r) {
    if (this.bands.length) {
      const k = this._bandAt(r);
      if (k >= 0) { this._cut(k, r); return true; }
    }
    return super.delete(r);
  }

  clear() {
    this.bands = [];
    super.clear();
  }

  get size() {
    let n = super.size;
    for (const { r0, r1 } of this.bands) n += r1 - r0 + 1;
    return n;
  }

  entries() { return this.expand() && super.entries(); }
  keys() { return this.expand() && super.keys(); }
  values() { return this.expand() && super.values(); }
  forEach(fn, self) { this.expand(); return super.forEach(fn, self); }
  [Symbol.iterator]() { return this.entries(); }
}

/** A parsed sheet's document outside its cells: a SheetPart's frame (workbook.js), or the document itself: what
 *  only reads the views, protection, validations or the extLst reads this, and a large sheet's tree isn't built for it. */
export const frameOf = (part) => part.frame || part.doc;

/** The rows of a model's cells that may hold more than formatted empty cells (SheetCells' `filled`; every row of a
 *  plain Map). */
export const filledRows = (cells) => (cells.filled ? cells.filled() : cells);

/** How many cells a model's cells are, formatted empty ones counted. */
export function cellCount(cells) {
  let n = cells.bandCells || 0;
  for (const [, line] of filledRows(cells)) n += line.size;
  return n;
}
