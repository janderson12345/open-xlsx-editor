// Sort as Excel does it (measured over COM: Worksheet.Sort, as its Sort box applies one, and Range.Sort with xlGuess,
// as its A to Z button guesses the headers; docs/SORT.md). Rows (or, left to right, columns) are reordered by their
// keys, ties kept in the order they were; each cell moves with its format (not its borders), a formula moved as a copy
// moves it; cells outside reading the sorted ones aren't changed. Returns a step as clipops.js does: { si, contents, styles, merges }.
import { shiftFormula } from '../calc/tokens.js';
import { MAX_ROWS, MAX_COLS } from './refs.js';

export const SORT_MESSAGES = {
  merged: 'To do this, all the merged cells need to be the same size.',
  array: "You can't change part of an array.",
  empty: "This can't be applied to the selected range. Select a single cell in a range and try again.",
};

// How text compares, as Excel was seen to compare it (United States settings): signs first, in this order (a
// non-breaking space as a space), then other signs by their code, then digits, then letters with their accents taken
// off (é as e, ñ as n, æ as ae, ß as ss); the case left aside unless asked for. Apostrophes, hyphens and en dashes are
// passed over: "co-op" falls between "coop" and "cop", after "coop".
const SIGNS = ' !"#$%&()*,./:;?@[\\]^_`{|}~’+<=>§°';
const SKIPPED = { "'": 1, '-': 2, '–': 3, ' ': 4 };
const EXPANDED = { 'æ': 'ae', 'Æ': 'AE', 'œ': 'oe', 'Œ': 'OE', 'ß': 'ss' };
const DIGITS = 1000;
const LETTERS = 2000;
const OTHERS = 3000;

/** A text's weights at each level: what it spells, its accents, its capitals, and the signs passed over. */
function textKey(text) {
  const base = [];
  const accent = [];
  const upper = [];
  const skipped = [];
  for (const raw of text) {
    if (SKIPPED[raw]) {
      skipped.push(base.length * 8 + SKIPPED[raw]);
      // (A non-breaking space spells a space; it only comes after one.)
      if (raw !== ' ') continue;
    }
    const ch = raw === ' ' ? ' ' : raw;
    for (const one of EXPANDED[ch] || ch) {
      const plain = one.normalize('NFD').replace(/[̀-ͯ]/g, '');
      const marks = one.normalize('NFD').slice(plain.length);
      const low = plain.toLowerCase();
      const sign = SIGNS.indexOf(low);
      let w;
      if (sign >= 0) w = sign;
      else if (low >= '0' && low <= '9') w = DIGITS + low.charCodeAt(0);
      else if (low >= 'a' && low <= 'z') w = LETTERS + low.charCodeAt(0);
      else if (low.toUpperCase() !== low) w = OTHERS + low.codePointAt(0);
      else w = SIGNS.length + low.codePointAt(0) / 0x110000;
      base.push(w);
      accent.push(marks ? marks.codePointAt(0) : 0);
      upper.push(plain !== low ? 1 : 0);
    }
  }
  return { base, accent, upper, skipped };
}
const lists = (a, b) => {
  for (let i = 0; i < Math.min(a.length, b.length); i++) if (a[i] !== b[i]) return a[i] < b[i] ? -1 : 1;
  return a.length - b.length;
};
/** Two texts as Excel orders them: below 0, 0 (the same: their order kept) or above. */
export function compareText(a, b, { matchCase = false } = {}) {
  const x = textKey(a);
  const y = textKey(b);
  return lists(x.base, y.base) || lists(x.accent, y.accent) || (matchCase ? lists(x.upper, y.upper) : 0) || lists(x.skipped, y.skipped);
}

const NUMBER_TEXT = /^\s*[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?\s*$/i;
/** A cell's sort value: { rank, n | text } (numbers 0, text 1, FALSE and TRUE 2, errors 3), or null for a blank. */
function sortValue(cell, textAsNumbers) {
  if (!cell || ((cell.f == null || cell.f === '') && cell.v == null)) return null;
  if (cell.v == null) return { rank: 0, n: 0 };
  if (cell.t === 'e') return { rank: 3 };
  if (typeof cell.v === 'boolean') return { rank: 2, n: cell.v ? 1 : 0 };
  if (typeof cell.v === 'number') return { rank: 0, n: cell.v };
  const text = String(cell.v);
  if (textAsNumbers && NUMBER_TEXT.test(text)) return { rank: 0, n: Number(text) };
  return { rank: 1, text };
}
/** Two sort values, ascending (blanks are put last by the caller, whichever way). */
function compareValues(a, b, matchCase) {
  if (a.rank !== b.rank) return a.rank - b.rank;
  if (a.rank === 1) return compareText(a.text, b.text, { matchCase });
  if (a.rank === 3) return 0;
  return a.n - b.n;
}

const filled = (cell) => cell && ((cell.f != null && cell.f !== '') || cell.v != null);

/** The range around `sel` that holds data, as Excel's CurrentRegion finds it: grown while any filled cell touches it,
 *  side or corner. */
export function currentRegion(sheet, sel) {
  let { r0, c0, r1, c1 } = sel;
  for (let grown = true; grown;) {
    grown = false;
    const a = { r0: Math.max(0, r0 - 1), c0: Math.max(0, c0 - 1), r1: Math.min(MAX_ROWS - 1, r1 + 1), c1: Math.min(MAX_COLS - 1, c1 + 1) };
    let n = { r0, c0, r1, c1 };
    for (let r = a.r0; r <= a.r1; r++) {
      const line = sheet.cells.get(r);
      if (!line) continue;
      const edge = r < r0 || r > r1;
      for (const [c, cell] of line) {
        if (c < a.c0 || c > a.c1 || (!edge && c >= c0 && c <= c1) || !filled(cell)) continue;
        n = { r0: Math.min(n.r0, r), c0: Math.min(n.c0, c), r1: Math.max(n.r1, r), c1: Math.max(n.c1, c) };
      }
    }
    if (n.r0 !== r0 || n.c0 !== c0 || n.r1 !== r1 || n.c1 !== c1) { ({ r0, c0, r1, c1 } = n); grown = true; }
  }
  return { r0, c0, r1, c1 };
}

/**
 * What Sort sorts for selection `sel`: one cell, the data around it (null when there is none: Excel's refusal);
 * whole rows or columns cut to the used range; any other range as it is.
 */
export function sortArea(sheet, sel) {
  if (sel.r0 === sel.r1 && sel.c0 === sel.c1) {
    const a = currentRegion(sheet, sel);
    const one = a.r0 === a.r1 && a.c0 === a.c1;
    return one && !filled(sheet.cells.get(a.r0)?.get(a.c0)) ? null : a;
  }
  return { r0: sel.r0, c0: sel.c0, r1: Math.min(sel.r1, Math.max(sel.r0, sheet.maxRow)), c1: Math.min(sel.c1, Math.max(sel.c0, sheet.maxCol)) };
}

/** Whether a selection of several cells has data beside it that a sort of it would leave behind (Excel's Sort
 *  Warning): the region around it is wider (for a sort top to bottom). */
export function dataBeside(sheet, sel) {
  if (sel.r0 === sel.r1 && sel.c0 === sel.c1) return null;
  const a = sortArea(sheet, sel);
  const around = currentRegion(sheet, a);
  return around.c0 < a.c0 || around.c1 > a.c1 ? { ...around, r0: a.r0, r1: a.r1 } : null;
}

/**
 * Whether Excel would take the first row of `area` as headers when not told (as its A to Z button and Sort box
 * guess): only when that row has fewer than two blank cells, and some column has text over a number, or a cell of
 * the row is formatted otherwise than the one below it (bold, fill, border, number format).
 */
export function guessHeader(wb, si, area) {
  if (area.r1 <= area.r0) return false;
  const sheet = wb.sheet(si);
  const at = (r, c) => sheet.cells.get(r)?.get(c);
  let blanks = 0;
  let sign = false;
  const xf = (cell) => wb.styles.cellXfs[cell?.s || 0] || {};
  for (let c = area.c0; c <= area.c1; c++) {
    const top = at(area.r0, c);
    const next = at(area.r0 + 1, c);
    const a = sortValue(top, false);
    const b = sortValue(next, false);
    if (!a) blanks++;
    if (a?.rank === 1 && b?.rank === 0) sign = true;
    const x = xf(top);
    const y = xf(next);
    if (x.fontId !== y.fontId || x.fillId !== y.fillId || x.borderId !== y.borderId || x.numFmtId !== y.numFmtId) sign = true;
  }
  return blanks < 2 && sign;
}

/** Sort A to Z (`desc` false) or Z to A on selection `sel`, as Excel's buttons do it: what sortArea finds (with the
 *  data beside it when `expand`), by the active cell's column, the first row taken as headers when guessHeader says
 *  so. { area, keys, header }, or { error } when there is nothing to sort. */
export function quickSort(wb, si, sel, desc, { expand = false } = {}) {
  const sheet = wb.sheet(si);
  const area = (expand && dataBeside(sheet, sel)) || sortArea(sheet, sel);
  if (!area) return { error: SORT_MESSAGES.empty };
  return { area, keys: [{ at: sel.ac, desc }], header: guessHeader(wb, si, area) };
}

/** The merged ranges in `area`: [] when there are none; their size when every cell of the area is in one of them,
 *  all of a size (then each sorts as one); null when Excel refuses. */
function mergedBlocks(sheet, area) {
  const inside = sheet.merges.filter((m) => m.r0 <= area.r1 && m.r1 >= area.r0 && m.c0 <= area.c1 && m.c1 >= area.c0);
  if (!inside.length) return { h: 1, w: 1 };
  const h = inside[0].r1 - inside[0].r0 + 1;
  const w = inside[0].c1 - inside[0].c0 + 1;
  const fits = inside.every((m) => m.r1 - m.r0 + 1 === h && m.c1 - m.c0 + 1 === w && m.r0 >= area.r0 && m.r1 <= area.r1 && m.c0 >= area.c0 && m.c1 <= area.c1);
  const cells = (area.r1 - area.r0 + 1) * (area.c1 - area.c0 + 1);
  return fits && inside.length * h * w === cells ? { h, w } : null;
}

/**
 * Sort `area` of sheet si. `keys`: [{ at (the column, or for `leftToRight` the row, by its index), desc }], the first
 * deciding first; `header`: its first row (column) stays; `matchCase`: capitals after small letters, else alike;
 * `textAsNumbers`: text that reads as a number sorts as one. A step ({ si, contents, styles, merges, area }), or
 * { error } with Excel's refusal.
 */
export function sortChanges(wb, si, area, { keys, header = false, matchCase = false, leftToRight = false, textAsNumbers = false }) {
  const sheet = wb.sheet(si);
  const meets = (a) => a.r0 <= area.r1 && a.r1 >= area.r0 && a.c0 <= area.c1 && a.c1 >= area.c0;
  if (sheet.arrays.some(meets)) return { error: SORT_MESSAGES.array };
  // (A dynamic array's spill too, any part of it, its formula's cell in or not: Excel refuses, measured.)
  for (const line of sheet.cells.values()) for (const cell of line.values()) if (cell.spill && meets(cell.spill)) return { error: SORT_MESSAGES.array };
  const block = mergedBlocks(sheet, area);
  if (!block) return { error: SORT_MESSAGES.merged };
  // (Lines are rows, or columns left to right; each unit is one line, or a merged range's height of them.)
  const size = leftToRight ? block.w : block.h;
  const [first, last] = leftToRight ? [area.c0, area.c1] : [area.r0, area.r1];
  const start = first + (header ? size : 0);
  const at = (line, across) => (leftToRight ? sheet.cells.get(across)?.get(line) : sheet.cells.get(line)?.get(across));
  const units = [];
  for (let line = start; line + size - 1 <= last; line += size) units.push({ line, values: keys.map((k) => sortValue(at(line, k.at), textAsNumbers)) });
  const order = units.map((u, i) => ({ ...u, i }));
  order.sort((p, q) => {
    for (let k = 0; k < keys.length; k++) {
      const a = p.values[k];
      const b = q.values[k];
      if (!a || !b) { if (a !== b) return a ? -1 : 1; continue; }
      const d = compareValues(a, b, matchCase);
      if (d) return keys[k].desc ? -d : d;
    }
    return p.i - q.i;
  });
  const out = { si, contents: [], styles: [], merges: sheet.merges, area: { ...area } };
  const [a0, a1] = leftToRight ? [area.r0, area.r1] : [area.c0, area.c1];
  // (A cell's format moves with it, but not its borders: Excel leaves those where they were, a bottom border seen to
  // stay under the same cell.)
  const memo = new Map();
  const restyle = (s, here) => {
    const want = wb.styles.cellXfs[here]?.borderId || 0;
    if ((wb.styles.cellXfs[s]?.borderId || 0) === want) return s;
    const key = `${s}:${want}`;
    if (!memo.has(key)) {
      const sides = wb.styles.borders[want];
      memo.set(key, wb.styleWith(s, { border: Object.fromEntries(['left', 'right', 'top', 'bottom'].map((k) => [k, sides[k]?.style ? sides[k] : null])) }));
    }
    return memo.get(key);
  };
  order.forEach((from, i) => {
    const to = units[i].line;
    if (from.line === to) return;
    for (let k = 0; k < size; k++) {
      for (let across = a0; across <= a1; across++) {
        const [r, c] = leftToRight ? [across, to + k] : [to + k, across];
        const [fr, fc] = leftToRight ? [across, from.line + k] : [from.line + k, across];
        const cell = sheet.cells.get(fr)?.get(fc);
        const now = sheet.cells.get(r)?.get(c);
        if (!cell && !now) continue;
        out.styles.push({ r, c, s: restyle(cell?.s || 0, now?.s || 0) });
        if (cell?.f != null && cell.f !== '') out.contents.push({ r, c, formula: shiftFormula(cell.f, r - fr, c - fc) });
        else out.contents.push({ r, c, value: cell?.t === 'e' ? { code: String(cell.v) } : cell?.v ?? null });
      }
    }
  });
  return out;
}
