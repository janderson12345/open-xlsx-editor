// Excel's data tools, as Excel does them (measured over COM: tools/excel-compare/excel-datatools.ps1,
// test/datatools-excel.json). Fill > Series: each row (or column) of the selection from its first cell on, by a step
// (linear), times a step (growth), or by days, weekdays, months or years from the first date (months and years from the
// start, the day kept or the month's last: Jan 31, Feb 28, Mar 31), up to a stop value; Trend fits a line (or a growth
// curve) to the values there and writes it over them. Remove Duplicates: rows alike in the columns chosen (text case
// aside, numbers by value) after the first go, the rows under moved up within the range. Text to Columns: each cell
// split at its delimiters (quotes keeping a piece whole), each piece read as typed (numbers, dates, TRUE, 5%). Each
// returns steps ({ si, contents, styles, area }) for the editor's history, as fillops.js's do.
import { shiftFormula } from '../calc/tokens.js';
import { parseInput, dateSerial, fullYear } from './input.js';
import { formatCodeOf } from './styles.js';
import { cellText } from './numfmt.js';

const cellAt = (sheet, r, c) => sheet.cells.get(r)?.get(c) || null;
const DAY_MS = 86400000;
const dateOf = (serial) => new Date(Date.UTC(1899, 11, 30) + Math.round(serial) * DAY_MS);
const serialOf = (d) => Math.round((d.getTime() - Date.UTC(1899, 11, 30)) / DAY_MS);
/** A date `n` months on from `start`, the day kept or the month's last (Excel's EDATE). */
function addMonths(serial, n) {
  const d = dateOf(serial);
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth() + n;
  const last = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
  return serialOf(new Date(Date.UTC(y, m, Math.min(d.getUTCDate(), last)))) + (serial - Math.floor(serial));
}
/** The weekday `k` weekdays on from `serial` (Saturday and Sunday skipped; a start on a weekend counts from it). */
function addWeekdays(serial, k) {
  let s = serial;
  let left = Math.abs(k);
  const step = k < 0 ? -1 : 1;
  while (left > 0) { s += step; const w = dateOf(s).getUTCDay(); if (w !== 0 && w !== 6) left--; }
  return s;
}
/** Least squares: a straight line (or, `growth`, y = b·m^x) through points (x, y). */
function fit(points, growth) {
  const n = points.length;
  const ys = points.map(([, y]) => (growth ? Math.log(y) : y));
  const mx = points.reduce((s, [x]) => s + x, 0) / n;
  const my = ys.reduce((s, y) => s + y, 0) / n;
  let sxy = 0;
  let sxx = 0;
  points.forEach(([x], i) => { sxy += (x - mx) * (ys[i] - my); sxx += (x - mx) ** 2; });
  const slope = sxx ? sxy / sxx : 0;
  const at = (x) => my + slope * (x - mx);
  return growth ? (x) => Math.exp(at(x)) : at;
}

/**
 * Fill > Series over `sel` of sheet si: `by` 'rows' (each row across) or 'columns' (each column down); `type`
 * 'linear', 'growth' or 'date' (`unit` 'day', 'weekday', 'month' or 'year'); `step`; `stop` (or null); `trend`.
 */
export function fillSeries(wb, si, sel, { by = 'columns', type = 'linear', unit = 'day', step = 1, stop = null, trend = false }) {
  const sheet = wb.sheet(si);
  const out = { si, contents: [], styles: [], area: { ...sel } };
  const lines = by === 'columns' ? [sel.c0, sel.c1] : [sel.r0, sel.r1];
  const len = by === 'columns' ? sel.r1 - sel.r0 + 1 : sel.c1 - sel.c0 + 1;
  for (let line = lines[0]; line <= lines[1]; line++) {
    const at = (k) => (by === 'columns' ? { r: sel.r0 + k, c: line } : { r: line, c: sel.c0 + k });
    const first = cellAt(sheet, at(0).r, at(0).c);
    if (typeof first?.v !== 'number') continue;
    const s = first.s || 0;
    let values;
    if (trend && type !== 'date') {
      const known = [];
      for (let k = 0; k < len; k++) { const x = cellAt(sheet, at(k).r, at(k).c); if (typeof x?.v === 'number') known.push([k + 1, x.v]); else break; }
      const f = fit(known.length > 1 ? known : [[1, first.v], [2, first.v + (type === 'growth' ? 0 : 1)]], type === 'growth');
      values = Array.from({ length: len }, (_, k) => f(k + 1));
    } else {
      values = [first.v];
      for (let k = 1; k < len; k++) {
        const v = type === 'growth' ? values[0] * step ** k
          : type === 'date' ? (unit === 'month' ? addMonths(values[0], step * k) : unit === 'year' ? addMonths(values[0], 12 * step * k) : unit === 'weekday' ? addWeekdays(values[0], step * k) : values[0] + step * k)
            : values[0] + step * k;
        // (Up to the stop value and no further: past it, the cells are left as they are.)
        if (stop != null && (step >= 0 || type === 'growth' ? v > stop : v < stop)) break;
        values.push(Number(v.toPrecision(15)));
      }
    }
    values.forEach((v, k) => {
      const p = at(k);
      if (k === 0 && !trend) return;
      out.contents.push({ ...p, value: v });
      out.styles.push({ ...p, s });
    });
  }
  return [out];
}

/** A value as Remove Duplicates compares it: numbers by value, text case aside. */
const key = (cell) => (cell?.v == null ? '' : typeof cell.v === 'number' ? `n:${cell.v}` : `s:${String(cell.v).toLowerCase()}`);

/**
 * Data > Remove Duplicates on `range` of sheet si, comparing the columns `cols` (absolute indexes), `header` its first
 * row a heading. Returns { steps, removed, kept } (Excel's message: "n duplicate values found and removed; k unique
 * values remain.").
 */
export function removeDuplicates(wb, si, range, cols, header = true) {
  const sheet = wb.sheet(si);
  const r0 = range.r0 + (header ? 1 : 0);
  const seen = new Set();
  const keep = [];
  for (let r = r0; r <= range.r1; r++) {
    const k = cols.map((c) => key(cellAt(sheet, r, c))).join('\u0000');
    if (seen.has(k)) continue;
    seen.add(k);
    keep.push(r);
  }
  const out = { si, contents: [], styles: [], area: { ...range } };
  for (let r = r0; r <= range.r1; r++) {
    const from = keep[r - r0];
    for (let c = range.c0; c <= range.c1; c++) {
      if (from === r) continue;
      const cell = from != null ? cellAt(sheet, from, c) : null;
      // (A row moved up: its formulas move with it, as cells shifted up do.)
      out.contents.push(cell?.f != null && cell.f !== '' ? { r, c, formula: shiftFormula(cell.f, r - from, 0) } : { r, c, value: cell ? (cell.t === 'e' ? { code: String(cell.v) } : cell.v ?? null) : null });
      out.styles.push({ r, c, s: cell?.s || 0 });
    }
  }
  return { steps: [out], removed: range.r1 - r0 + 1 - keep.length, kept: keep.length };
}

/** A line split at the delimiters (`set`: the characters), quotes keeping what they hold whole; with `consecutive`,
 *  delimiters together count once. */
export function splitLine(text, set, { consecutive = false, quote = '"' } = {}) {
  const out = [];
  let piece = '';
  let quoted = false;
  let last = -2;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quote && ch === quote) {
      if (quoted && text[i + 1] === quote) { piece += quote; i++; continue; }
      quoted = !quoted;
      continue;
    }
    if (!quoted && set.includes(ch)) {
      if (consecutive && last === i - 1) { last = i; continue; }
      out.push(piece);
      piece = '';
      last = i;
      continue;
    }
    piece += ch;
  }
  out.push(piece);
  return out;
}

const DATE_ORDERS = ['mdy', 'dmy', 'ymd', 'myd', 'dym', 'ydm'];

/** A piece of a Date column (`order`, as 'dmy'), as Excel reads it (excel-datatools.ps1): its three numbers, apart by
 *  / - or ., taken in that order, a two-digit year as typed ones are (03 is 2003, 99 is 1999); null when they aren't
 *  a date, the piece then read as typed (2026-03-04 in an MDY column: 4 March). */
function orderedDate(text, order, date1904) {
  const m = /^(\d{1,4})[/.-](\d{1,4})[/.-](\d{1,4})$/.exec(text.trim());
  if (!m) return null;
  const v = Object.fromEntries([...order].map((k, i) => [k, m[i + 1]]));
  return dateSerial(fullYear(v.y), Number(v.m), Number(v.d), date1904);
}

/** Column breaks to suggest for Fixed width, as Excel's wizard suggests some when it opens (its rule isn't reachable
 *  over COM, so this is the editor's own): a break at each place where every line has a space just before and some
 *  line has a character, after the first character of every line. */
export function guessBreaks(lines) {
  const texts = lines.filter((t) => t.trim());
  if (!texts.length) return [];
  const width = Math.max(...texts.map((t) => t.length));
  const blank = (k) => texts.every((t) => k >= t.length || t[k] === ' ');
  const out = [];
  for (let k = 1; k < width; k++) if (blank(k - 1) && !blank(k)) out.push(k);
  return out;
}

/**
 * Data > Text to Columns on the one-column `range` of sheet si: each cell split (`delimiters` { tab, semicolon,
 * comma, space, other }, `consecutive`, `quote`), the pieces written across from `dest` ({ r, c }: the range's own
 * first cell by default), each read as typed; an empty piece leaves its cell. Fixed width (`breaks`: the character
 * places columns start at, after the first's at 0), as Excel splits (excel-datatools.ps1): the characters between
 * two breaks, spaces at either end trimmed; a number's cell split from the text it shows. `formats`, one a column (the
 * wizard's Column data format): 'general' (as typed), 'text' (kept as text, in the @ format), 'skip' (Do not import
 * column: left out, the columns after it moved left) or a date's order ('mdy', 'dmy', 'ymd', 'myd', 'dym', 'ydm':
 * see orderedDate).
 */
export function textToColumns(wb, si, range, { delimiters = {}, consecutive = false, quote = '"', dest = null, breaks = null, formats = [] } = {}) {
  const sheet = wb.sheet(si);
  const set = [delimiters.tab ? '\t' : '', delimiters.semicolon ? ';' : '', delimiters.comma ? ',' : '', delimiters.space ? ' ' : '', delimiters.other || ''].join('');
  const to = dest || { r: range.r0, c: range.c0 };
  const out = { si, contents: [], styles: [], area: { r0: to.r, c0: to.c, r1: to.r + range.r1 - range.r0, c1: to.c } };
  for (let r = range.r0; r <= range.r1; r++) {
    const cell = cellAt(sheet, r, range.c0);
    if (cell?.v == null || (typeof cell.v !== 'string' && !(breaks && cell.t === 'n'))) continue;
    let pieces;
    if (breaks) {
      const text = typeof cell.v === 'string' ? cell.v : cellText(cell, Infinity, formatCodeOf(wb.styles, cell.s || 0), wb.date1904);
      const at = [0, ...[...new Set(breaks)].filter((b) => b > 0).sort((a, b) => a - b)];
      pieces = at.map((b, k) => text.slice(b, at[k + 1]).replace(/^ +| +$/g, ''));
    } else pieces = set ? splitLine(cell.v, set, { consecutive, quote }) : [cell.v];
    pieces.map((p, k) => [p, formats[k] || 'general']).filter(([, f]) => f !== 'skip').forEach(([p, f], k) => {
      const at = { r: to.r + r - range.r0, c: to.c + k };
      if (p === '') { if (k === 0 && at.r === r && at.c === range.c0) out.contents.push({ ...at, value: null }); return; }
      const serial = DATE_ORDERS.includes(f) ? orderedDate(p, f, wb.date1904) : null;
      const read = f === 'text' ? { value: p, format: '@' } : serial != null ? { value: serial, format: 'm/d/yyyy' } : parseInput(p, { date1904: wb.date1904 });
      out.contents.push({ ...at, value: read.value ?? p });
      out.styles.push({ ...at, s: read.format ? wb.styleWith(cell.s || 0, { numFmt: read.format }) : cell.s || 0 });
      out.area.c1 = Math.max(out.area.c1, at.c);
    });
  }
  return [out];
}
