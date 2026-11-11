// Go To Special (Home > Find & Select > Go To Special, or F5 > Special...): the cells of a kind, found as Excel finds
// them (measured: tools/excel-compare/excel-special.ps1 on test/fixtures/special.xlsx, test/special-excel.json).
// - Where: one cell selected, the sheet's used range (Visible cells only and the rules' choices: the whole sheet);
//   more, only the cells selected.
// - A merged range is found whole when any cell of it is.
// - Row differences: in each row, the cells unlike the one in the active cell's column; Column differences, by
//   columns. A formula is unlike any constant, and like another formula only if the two are the same relative to
//   their cells (as R1C1 writes them); text is compared without regard to case.
// - Precedents and Dependents: of the selected cells, on this sheet; direct only or all levels.
// Excel lists the ranges found in an order of its own; here they go top to bottom, left to right (the same cells).
import { shiftFormula } from '../calc/tokens.js';
import { readNotes } from './notes.js';
import { MAX_ROWS, MAX_COLS } from './refs.js';

/** Excel's message when a choice finds nothing. */
export const NO_CELLS = 'No cells were found.';

/** Constants' and Formulas' types, as the box's check boxes (and Excel's XlSpecialCellsValue) number them. */
export const TYPES = { numbers: 1, text: 2, logicals: 4, errors: 16 };
const ALL_TYPES = 1 | 2 | 4 | 16;

const inside = (a, r, c) => r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1;
const clip = (a, b) => {
  const out = { r0: Math.max(a.r0, b.r0), c0: Math.max(a.c0, b.c0), r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1) };
  return out.r0 <= out.r1 && out.c0 <= out.c1 ? out : null;
};

/** The sheet's used range: its first to its last cell with anything (a value, a formula or a format) or a merge. */
export function usedRange(sheet) {
  let r0 = Infinity; let c0 = Infinity; let r1 = -1; let c1 = -1;
  for (const [r, line] of sheet.cells) for (const c of line.keys()) { r0 = Math.min(r0, r); r1 = Math.max(r1, r); c0 = Math.min(c0, c); c1 = Math.max(c1, c); }
  for (const m of sheet.merges) { r0 = Math.min(r0, m.r0); r1 = Math.max(r1, m.r1); c0 = Math.min(c0, m.c0); c1 = Math.max(c1, m.c1); }
  return r1 < 0 ? null : { r0, c0, r1, c1 };
}

/** Cells as ranges: each row's runs, a run joined to the one above when it spans the same columns. */
export function cellsToAreas(cells) {
  const byRow = new Map();
  for (const k of cells) {
    const [r, c] = k.split(',').map(Number);
    if (!byRow.has(r)) byRow.set(r, []);
    byRow.get(r).push(c);
  }
  const open = new Map(); // "c0,c1" -> the area reaching the row above
  const out = [];
  for (const r of [...byRow.keys()].sort((a, b) => a - b)) {
    const cols = byRow.get(r).sort((a, b) => a - b);
    const runs = [];
    for (const c of cols) { const last = runs[runs.length - 1]; if (last && last.c1 === c - 1) last.c1 = c; else runs.push({ c0: c, c1: c }); }
    for (const run of runs) {
      const k = `${run.c0},${run.c1}`;
      const above = open.get(k);
      if (above && above.r1 === r - 1) above.r1 = r;
      else { const a = { r0: r, c0: run.c0, r1: r, c1: run.c1 }; out.push(a); open.set(k, a); }
    }
  }
  return out.sort((a, b) => a.r0 - b.r0 || a.c0 - b.c0);
}

/** A cell's kind for Constants and Formulas (TYPES), or 0 when empty: a cell of an array formula is a formula. */
function typeOf(cell) {
  if (cell.t === 'e') return TYPES.errors;
  if (cell.t === 'b') return TYPES.logicals;
  if (cell.t === 's' || cell.t === 'str' || cell.t === 'inlineStr') return TYPES.text;
  return TYPES.numbers;
}
const isFormula = (sheet, cell, r, c) => (cell && cell.f != null) || sheet.arrays.some((a) => inside(a, r, c));
const hasValue = (cell) => !!cell && (cell.f != null || (cell.v != null && cell.v !== ''));

/** The rows (or columns) of first..last not hidden, as bands [a, b]. */
function shownBands(first, last, hidden) {
  const out = [];
  for (let i = first; i <= last; i++) {
    if (hidden(i)) continue;
    const band = out[out.length - 1];
    if (band && band[1] === i - 1) band[1] = i; else out.push([i, i]);
  }
  return out;
}
/** Hidden rows and columns of a sheet, as tests. */
function hiddenTests(sheet) {
  const rowHidden = (r) => !!sheet.rows.get(r)?.hidden;
  const hiddenCols = sheet.cols.filter((x) => x.hidden);
  return { rowHidden, colHidden: (c) => hiddenCols.some((x) => c >= x.min && c <= x.max) };
}

/** Visible cells only: each range of `scopes` less its hidden rows and columns. (Rows and columns past those the sheet
 *  says anything about are walked as one band, so a whole sheet is quick.) */
function visibleAreas(sheet, scopes) {
  const { rowHidden, colHidden } = hiddenTests(sheet);
  const lastRow = Math.max(-1, ...sheet.rows.keys());
  const lastCol = Math.max(-1, ...sheet.cols.map((x) => x.max));
  const bands = (a0, a1, last, hidden) => {
    const out = shownBands(a0, Math.min(a1, last + 1), hidden);
    const end = out[out.length - 1];
    if (a1 > last + 1) { if (end && end[1] === last + 1) end[1] = a1; else out.push([Math.max(a0, last + 1), a1]); }
    return out;
  };
  const out = [];
  for (const s of scopes) {
    const rows = bands(s.r0, s.r1, lastRow, rowHidden);
    const cols = bands(s.c0, s.c1, lastCol, colHidden);
    for (const [r0, r1] of rows) for (const [c0, c1] of cols) out.push({ r0, c0, r1, c1 });
  }
  return out;
}

/** Excel's current region round (r, c): grown while any cell next to it (diagonals too) has a value or formula. */
export function currentRegion(sheet, r, c) {
  const a = { r0: r, c0: c, r1: r, c1: c };
  const full = (rr, cc) => rr >= 0 && cc >= 0 && rr < MAX_ROWS && cc < MAX_COLS && hasValue(sheet.cells.get(rr)?.get(cc));
  const any = (r0, c0, r1, c1) => { for (let rr = r0; rr <= r1; rr++) for (let cc = c0; cc <= c1; cc++) if (full(rr, cc)) return true; return false; };
  for (let grew = true; grew;) {
    grew = false;
    if (a.r0 > 0 && any(a.r0 - 1, a.c0 - 1, a.r0 - 1, a.c1 + 1)) { a.r0--; grew = true; }
    if (any(a.r1 + 1, a.c0 - 1, a.r1 + 1, a.c1 + 1)) { a.r1++; grew = true; }
    if (a.c0 > 0 && any(a.r0, a.c0 - 1, a.r1, a.c0 - 1)) { a.c0--; grew = true; }
    if (any(a.r0, a.c1 + 1, a.r1, a.c1 + 1)) { a.c1++; grew = true; }
  }
  return a;
}

/** What row and column differences compare: a formula as it would be at (r, c) (moved there), a value by its type
 *  and value (text without regard to case), or empty. */
function likeness(sheet, cell, from, to) {
  if (cell && cell.f != null) return `f${shiftFormula(cell.f, to.r - from.r, to.c - from.c).toUpperCase()}`;
  if (!hasValue(cell)) return 'empty';
  const v = typeof cell.v === 'string' ? cell.v.toLowerCase() : String(cell.v);
  return `${typeOf(cell)}:${v}`;
}

/** Row ('row') or column differences over the selection's ranges, by the active cell's column (or row). */
function differences(sheet, areas, active, by) {
  const found = new Set();
  const at = (r, c) => sheet.cells.get(r)?.get(c);
  for (const a of areas) {
    for (let i = by === 'row' ? a.r0 : a.c0; i <= (by === 'row' ? a.r1 : a.c1); i++) {
      const base = by === 'row' ? { r: i, c: active.c } : { r: active.r, c: i };
      for (let j = by === 'row' ? a.c0 : a.r0; j <= (by === 'row' ? a.c1 : a.r1); j++) {
        const here = by === 'row' ? { r: i, c: j } : { r: j, c: i };
        if (here.r === base.r && here.c === base.c) continue;
        if (likeness(sheet, at(here.r, here.c), here, base) !== likeness(sheet, at(base.r, base.c), base, base)) found.add(`${here.r},${here.c}`);
      }
    }
  }
  return found;
}

/** Precedents ('precedents') or dependents of the cells in `areas` on sheet si, direct only or all levels. */
function links(wb, si, areas, which, all) {
  const calc = wb.calc;
  calc.build();
  const found = new Set();
  const add = (a) => { for (let r = a.r0; r <= a.r1; r++) for (let c = a.c0; c <= a.c1; c++) found.add(`${r},${c}`); };
  const seen = new Set();
  let next = areas.map((a) => ({ si, ...a }));
  while (next.length) {
    const now = next;
    next = [];
    for (const rect of now) {
      const step = which === 'precedents' ? precedentsOf(calc, rect) : [...calc.readers(rect)].map((e) => ({ si: e.si, ...(e.array || { r0: e.r, c0: e.c, r1: e.r, c1: e.c }) }));
      for (const d of step) {
        if (d.si !== si) continue;
        const k = `${d.r0},${d.c0},${d.r1},${d.c1}`;
        if (seen.has(k)) continue;
        seen.add(k);
        add(d);
        if (all) next.push(d);
      }
    }
  }
  return found;
}
/** The ranges the formulas in a rectangle read. */
function precedentsOf(calc, rect) {
  const out = [];
  for (const e of calc.entries.values()) if (e.si === rect.si && inside(rect, e.r, e.c)) out.push(...e.deps);
  return out;
}

/** The ranges of the rule blocks (conditional formats' or validations') covering the scopes, or only those of the
 *  block covering the active cell (`same`). */
function ruleAreas(blocks, scopes, active, same) {
  const chosen = same ? blocks.filter((b) => b.ranges.some((x) => inside(x, active.r, active.c))) : blocks;
  const out = new Map();
  for (const b of chosen) for (const x of b.ranges) for (const s of scopes) { const part = clip(x, s); if (part) out.set(`${part.r0},${part.c0},${part.r1},${part.c1}`, part); }
  return [...out.values()].sort((a, b) => a.r0 - b.r0 || a.c0 - b.c0);
}

/** The cell to make active in ranges found: the first not in a hidden row or column (top to bottom, left to right),
 *  or the first cell when all are hidden. (Excel makes its first range's top-left cell active, in an order of its own
 *  that can start in a hidden row; a cell out of sight is no use to Name Box or keyboard.) */
export function firstShown(sheet, areas) {
  const { rowHidden, colHidden } = hiddenTests(sheet);
  for (const a of areas) {
    let r = a.r0;
    while (r <= a.r1 && rowHidden(r)) r++;
    let c = a.c0;
    while (c <= a.c1 && colHidden(c)) c++;
    if (r <= a.r1 && c <= a.c1) return { r, c };
  }
  return { r: areas[0].r0, c: areas[0].c0 };
}

/**
 * The cells Go To Special finds on sheet si for a selection `sel` (its ranges `areas`, the active cell `active`):
 * { areas: [{ r0, c0, r1, c1 }] } or { error: NO_CELLS }. `choice` = { kind, types (Constants, Formulas: TYPES
 * summed), all (Precedents, Dependents: all levels) }; kinds: notes, constants, formulas, blanks, region, array,
 * rowDiff, colDiff, precedents, dependents, lastCell, visible, cfAll, cfSame, dvAll, dvSame.
 */
export function specialCells(wb, si, areas, active, choice) {
  const sheet = wb.sheet(si);
  const one = areas.length === 1 && areas[0].r0 === areas[0].r1 && areas[0].c0 === areas[0].c1;
  const used = usedRange(sheet);
  const sheetWide = [{ r0: 0, c0: 0, r1: MAX_ROWS - 1, c1: MAX_COLS - 1 }];
  const scopes = one ? (used ? [used] : []) : areas;
  const inScope = (r, c) => scopes.some((s) => inside(s, r, c));
  const done = (list) => (list.length ? { areas: list } : { error: NO_CELLS });
  const found = new Set();
  const { kind } = choice;
  if (kind === 'region') return done([currentRegion(sheet, active.r, active.c)]);
  if (kind === 'array') return done(sheet.arrays.filter((a) => inside(a, active.r, active.c)).map((a) => ({ r0: a.r0, c0: a.c0, r1: a.r1, c1: a.c1 })));
  if (kind === 'lastCell') return done(used ? [{ r0: used.r1, c0: used.c1, r1: used.r1, c1: used.c1 }] : [{ r0: 0, c0: 0, r1: 0, c1: 0 }]);
  if (kind === 'visible') return done(visibleAreas(sheet, one ? sheetWide : areas));
  if (kind === 'cfAll' || kind === 'cfSame') return done(ruleAreas(sheet.condFormats, one ? sheetWide : areas, active, kind === 'cfSame'));
  if (kind === 'dvAll' || kind === 'dvSame') return done(ruleAreas(sheet.validations, one ? sheetWide : areas, active, kind === 'dvSame'));
  if (kind === 'rowDiff' || kind === 'colDiff') for (const k of differences(sheet, areas, active, kind === 'rowDiff' ? 'row' : 'col')) found.add(k);
  else if (kind === 'precedents' || kind === 'dependents') for (const k of links(wb, si, areas, kind, !!choice.all)) found.add(k);
  else if (kind === 'notes') {
    for (const k of readNotes(wb, si).keys()) { const [r, c] = k.split(',').map(Number); if (inScope(r, c)) found.add(`${r},${c}`); }
  } else if (kind === 'constants' || kind === 'formulas') {
    const types = choice.types ?? ALL_TYPES;
    for (const [r, line] of sheet.cells) {
      for (const [c, cell] of line) {
        if (!inScope(r, c) || isFormula(sheet, cell, r, c) !== (kind === 'formulas')) continue;
        if (kind === 'constants' && !hasValue(cell)) continue;
        if (typeOf(cell) & types) found.add(`${r},${c}`);
      }
    }
  } else if (kind === 'blanks') {
    for (const s of scopes) {
      const a = used && clip(s, used);
      if (!a) continue;
      for (let r = a.r0; r <= a.r1; r++) for (let c = a.c0; c <= a.c1; c++) { const cell = sheet.cells.get(r)?.get(c); if (!hasValue(cell) && !isFormula(sheet, cell, r, c)) found.add(`${r},${c}`); }
    }
  } else throw new Error(`Go To Special: no choice ${kind}`);
  // (A merged range found whole when any of its cells is.)
  for (const m of sheet.merges) {
    let hit = false;
    for (let r = m.r0; r <= m.r1 && !hit; r++) for (let c = m.c0; c <= m.c1 && !hit; c++) hit = found.has(`${r},${c}`);
    if (hit) for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) found.add(`${r},${c}`);
  }
  return done(cellsToAreas(found));
}

