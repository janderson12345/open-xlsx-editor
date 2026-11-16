// Fill and the fill handle, as Excel does them (measured over COM, docs/FILL.md). Fill Down, Right, Up and Left
// (Ctrl+D, Ctrl+R) copy the first row or column of the selection (its last for Up and Left) over the rest; the fill
// handle (Range.AutoFill) carries cells on: numbers as a series, text with a number counting on, day and month names
// and quarters round their lists, dates by days, months or years, formulas moved as a copy moves them, formats
// repeating. Both return steps as clipops.js does: [{ si, contents, styles, merges }].
import { shiftFormula } from '../calc/tokens.js';
import { formatCodeOf } from './styles.js';
import { isDateFormat, dateParts } from './format.js';
import { dateSerial } from './input.js';
import { DATE_FILLS, fillDates } from './filldates.js';
import { currentRegion } from './sortops.js';

const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
/** Excel's built-in lists: short and long day and month names. */
const LISTS = [DAYS.map((d) => d.slice(0, 3)), DAYS, MONTHS.map((m) => m.slice(0, 3)), MONTHS];
/** The lists a workbook's fills go round: Excel's built-in ones, then the custom lists given (`wb.fillLists`, as
 *  Excel's Edit Custom Lists: each a list of words, matched whatever their capitals). */
const listsOf = (wb) => [...LISTS, ...(Array.isArray(wb?.fillLists) ? wb.fillLists.filter((l) => Array.isArray(l) && l.length > 1).map((l) => l.map(String)) : [])];
/** A number as Excel keeps one (15 significant digits: 0.1 + 0.2 is 0.3). */
const exact = (n) => Number(n.toPrecision(15));

/** What a cell holds, as the fill handle sees it. */
function kindOf(wb, cell) {
  if (!cell || (cell.f == null && cell.v == null)) return { kind: 'blank' };
  if (cell.f != null && cell.f !== '') return { kind: 'formula' };
  if (typeof cell.v === 'number') {
    const code = formatCodeOf(wb.styles, cell.s || 0);
    if (!isDateFormat(code)) return { kind: 'number', n: cell.v };
    // (A time alone, its format without days or years, steps by hours; a date by days.)
    const bare = code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
    return { kind: /[dy]/i.test(bare) ? 'date' : 'time', n: cell.v };
  }
  if (typeof cell.v !== 'string') return { kind: 'text' };
  const t = cell.v;
  const lists = listsOf(wb);
  for (let li = 0; li < lists.length; li++) {
    const i = lists[li].findIndex((x) => x.toLowerCase() === t.toLowerCase());
    if (i >= 0) return { kind: 'list', list: li, i, look: t === t.toUpperCase() ? 'upper' : t === t.toLowerCase() ? 'lower' : 'title' };
  }
  // (Q1 to Q4, Qtr1, Quarter 1...: round the four quarters.)
  let m = /^(q|qtr|quarter)(\s*)([1-4])$/i.exec(t);
  if (m) return { kind: 'quarter', key: `${m[1]}${m[2]}`.toLowerCase(), before: `${m[1]}${m[2]}`, n: Number(m[3]) };
  m = /^(.*?\D|)(\d+)$/.exec(t);
  if (m) return { kind: 'counted', key: `${m[1]}|`, before: m[1], after: '', n: Number(m[2]), width: m[2].length };
  m = /^(\d+)(\D.*)$/.exec(t);
  if (m) return { kind: 'counted', key: `|${m[2]}`, before: '', after: m[2], n: Number(m[1]), width: m[1].length };
  return { kind: 'text' };
}

/** The months between two serial dates when one is the other some whole months on (EDATE's sense), else null. */
function monthsApart(a, b, date1904) {
  const p = dateParts(a, date1904);
  const q = dateParts(b, date1904);
  const months = (q.y - p.y) * 12 + (q.m - p.m);
  return months && addMonths(a, months, date1904) === Math.floor(b) ? months : null;
}
/** A serial date some months on, the day kept or the month's last (EDATE). */
function addMonths(serial, months, date1904) {
  const p = dateParts(serial, date1904);
  const total = p.y * 12 + (p.m - 1) + months;
  const y = Math.floor(total / 12);
  const m = (total % 12) + 1;
  const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return dateSerial(y, m, Math.min(p.d, last), date1904);
}

/** A straight line through values at 0, 1, ... (least squares, as Excel's trend for three or more): [start, step]. */
function line(values) {
  const n = values.length;
  const mx = (n - 1) / 2;
  const my = values.reduce((a, b) => a + b, 0) / n;
  let sxy = 0;
  let sxx = 0;
  values.forEach((v, i) => { sxy += (i - mx) * (v - my); sxx += (i - mx) ** 2; });
  const step = sxy / sxx;
  return [my - step * mx, step];
}

/** Runs of a line's cells the fill carries on together: next to each other and of one kind (numbers, the same text
 *  with a number, one list...). Each is { start, items, at(i) } where at(i) gives the run's i-th value on (i < 0
 *  before it), as a cell content: { value } or null to copy the cell itself. */
function runsOf(kinds, { series, date1904, lists = LISTS }) {
  const runs = [];
  const same = (a, b) => a.kind === b.kind && ['number', 'date', 'time', 'list', 'quarter', 'counted'].includes(a.kind) && a.key === b.key && a.list === b.list;
  kinds.forEach((k, i) => {
    const last = runs[runs.length - 1];
    if (last && same(last.items[0], k)) last.items.push(k);
    else runs.push({ start: i, items: [k] });
  });
  const alone = kinds.length === 1;
  for (const run of runs) {
    const { items } = run;
    const first = items[0];
    const nums = items.map((x) => x.n);
    if (first.kind === 'number' || first.kind === 'date' || first.kind === 'time') {
      if (items.length === 1) {
        // (One number alone is copied unless a series is asked for; among other cells it counts on by one. A date
        // goes on by a day, a time by an hour.)
        const step = first.kind === 'time' ? 1 / 24 : 1;
        run.at = alone && first.kind === 'number' && !series ? () => null : (i) => ({ value: exact(first.n + step * i) });
      } else if (first.kind === 'date' && items.length === 2 && monthsApart(nums[0], nums[1], date1904)) {
        const months = monthsApart(nums[0], nums[1], date1904);
        run.at = (i) => ({ value: addMonths(nums[0], months * i, date1904) });
      } else {
        const [a, step] = items.length === 2 ? [nums[0], nums[1] - nums[0]] : line(nums);
        run.at = (i) => (i < items.length && i >= 0 && items.length > 2 ? null : { value: exact(a + step * i) });
      }
    } else if (first.kind === 'list' || first.kind === 'quarter') {
      const size = first.kind === 'list' ? lists[first.list].length : 4;
      const step = items.length > 1 ? ((first.kind === 'list' ? items[1].i - first.i : items[1].n - first.n) + size) % size : 1;
      run.at = (i) => {
        const k = ((((first.kind === 'list' ? first.i : first.n - 1) + step * i) % size) + size) % size;
        if (first.kind === 'quarter') return { value: `${first.before}${k + 1}` };
        const word = lists[first.list][k];
        return { value: first.look === 'upper' ? word.toUpperCase() : first.look === 'lower' ? word.toLowerCase() : word };
      };
    } else if (first.kind === 'counted') {
      const step = items.length > 1 ? items[1].n - first.n : 1;
      run.at = (i) => {
        const n = first.n + step * i;
        // (Counting below 0 keeps the number's size and the sign goes, as Excel's "Item 1" filled upwards does.)
        return { value: `${first.before}${String(Math.abs(n)).padStart(first.width, '0')}${first.after}` };
      };
    } else run.at = () => null;
  }
  return runs;
}

/** Cell (r, c) of sheet si as a content change that would put it at (r + dr, c + dc): its value, or its formula moved. */
function copyOf(sheet, r, c, dr, dc, to) {
  const cell = sheet.cells.get(r)?.get(c);
  if (cell?.f != null && cell.f !== '') return { ...to, formula: shiftFormula(cell.f, dr, dc) };
  return { ...to, value: cell?.t === 'e' ? { code: String(cell.v) } : cell?.v ?? null };
}

/**
 * Fill Down ('down', Ctrl+D), Right ('right', Ctrl+R), Up or Left on `sel` of sheet si: its first row (column; for Up
 * and Left its last) copied over the rest, formulas moved and formats with them. A selection one row high (one column
 * wide) takes the row above (the column to the left), as Excel's Ctrl+D does.
 */
export function fillCopy(wb, si, sel, dir) {
  const sheet = wb.sheet(si);
  const rows = dir === 'down' || dir === 'up';
  let { r0, r1, c0, c1 } = sel;
  if (dir === 'down' && r0 === r1 && r0 > 0) r0 -= 1;
  if (dir === 'right' && c0 === c1 && c0 > 0) c0 -= 1;
  const out = { si, contents: [], styles: [], merges: sheet.merges, area: { r0, c0, r1, c1 } };
  const src = dir === 'down' ? r0 : dir === 'up' ? r1 : dir === 'right' ? c0 : c1;
  for (let r = r0; r <= r1; r++) {
    for (let c = c0; c <= c1; c++) {
      if ((rows ? r : c) === src) continue;
      const [fr, fc] = rows ? [src, c] : [r, src];
      out.contents.push(copyOf(sheet, fr, fc, r - fr, c - fc, { r, c }));
      out.styles.push({ r, c, s: sheet.cells.get(fr)?.get(fc)?.s || 0 });
    }
  }
  return [out];
}

/** Which way `to` extends `from` (a fill handle's drag): 'down', 'up', 'right' or 'left', or null when it doesn't. */
export function fillWay(from, to) {
  if (to.c0 === from.c0 && to.c1 === from.c1) return to.r1 > from.r1 && to.r0 === from.r0 ? 'down' : to.r0 < from.r0 && to.r1 === from.r1 ? 'up' : null;
  if (to.r0 === from.r0 && to.r1 === from.r1) return to.c1 > from.c1 && to.c0 === from.c0 ? 'right' : to.c0 < from.c0 && to.c1 === from.c1 ? 'left' : null;
  return null;
}

/**
 * Where a double click on the fill handle fills `sel` down to, as Excel fills (its window, test/fill-dblclick-excel.json,
 * 13 of 13): to the last row of the region round the selection (sortops.js currentRegion: grown while a filled cell
 * touches it, side or corner, as Ctrl+* finds it; so data beside it on the left or the right, or further over through
 * cells that touch), stopping before a cell already filled in the selection's own columns. Null when the region goes
 * no lower than the selection.
 */
export function doubleClickFill(sheet, sel) {
  const has = (r, c) => sheet.cells.get(r)?.get(c)?.v != null && sheet.cells.get(r).get(c).v !== '';
  let r1 = currentRegion(sheet, { r0: sel.r0, c0: sel.c0, r1: sel.r1, c1: sel.c1 }).r1;
  for (let r = sel.r1 + 1; r <= r1; r++) {
    for (let c = sel.c0; c <= sel.c1; c++) if (has(r, c)) { r1 = r - 1; break; }
  }
  return r1 <= sel.r1 ? null : { r0: sel.r0, c0: sel.c0, r1, c1: sel.c1 };
}

/**
 * The fill handle: `from` (the selection) carried on over `to` (the selection dragged out down, up, right or left),
 * as Excel's AutoFill. `type`: 'default' (Excel's guess), 'series' (Fill Series: one number counts on), 'copy' (Copy
 * Cells), 'formats' (Fill Formatting Only), 'values' (Fill Without Formatting), or 'days', 'weekdays', 'months',
 * 'years' (Fill Days ... Fill Years: filldates.js). Each column (row) on its own; the
 * copy's formats repeat over the new cells.
 */
export function autoFill(wb, si, from, to, type = 'default') {
  const way = fillWay(from, to);
  if (!way) return null;
  const sheet = wb.sheet(si);
  const rows = way === 'down' || way === 'up';
  const out = { si, contents: [], styles: [], merges: sheet.merges, area: { ...to } };
  const [a0, a1] = rows ? [from.r0, from.r1] : [from.c0, from.c1];
  const n = a1 - a0 + 1;
  const [b0, b1] = rows ? [to.r0, to.r1] : [to.c0, to.c1];
  const [l0, l1] = rows ? [from.c0, from.c1] : [from.r0, from.r1];
  for (let line = l0; line <= l1; line++) {
    const at = (i) => (rows ? [i, line] : [line, i]);
    const cells = Array.from({ length: n }, (_, k) => sheet.cells.get(at(a0 + k)[0])?.get(at(a0 + k)[1]));
    // (Fill Days ... Years: the dates going on in that unit, the cells after them from the new values.)
    // (Downwards and rightwards: Excel's measured ones; upwards and leftwards Excel's guess, as before.)
    const dated = DATE_FILLS.includes(type) && (way === 'down' || way === 'right') ? fillDates(cells.map((cell) => (typeof cell?.v === 'number' && cell.t !== 's' ? cell.v : null)), type, b1 - a1, wb.date1904) : null;
    const runs = type === 'copy' || type === 'formats' || dated ? null : runsOf(cells.map((cell) => kindOf(wb, cell)), { series: type === 'series', date1904: wb.date1904, lists: listsOf(wb) });
    for (let i = b0; i <= b1; i++) {
      if (i >= a0 && i <= a1) continue;
      const k = i - a0;
      const cycle = Math.floor(k / n);
      const j = k - cycle * n;
      const [r, c] = at(i);
      const [fr, fc] = at(a0 + j);
      if (type !== 'values') out.styles.push({ r, c, s: cells[j]?.s || 0 });
      if (type === 'formats') continue;
      if (dated) {
        // (The k-th cell after the dates takes the k-th value.)
        out.contents.push({ r, c, value: dated[i - a1 - 1] });
        continue;
      }
      const run = runs?.find((x) => j >= x.start && j < x.start + x.items.length);
      const next = run && run.at(cycle * run.items.length + (j - run.start));
      out.contents.push(next ? { r, c, ...next } : copyOf(sheet, fr, fc, r - fr, c - fc, { r, c }));
    }
  }
  return [out];
}

/**
 * Fill Down, Right, Up or Left over several areas ([{ r0, c0, r1, c1 }]), as Excel did it (test/multirange-excel.json):
 * each column (Down, Up) or row (Right, Left) filled over the cells selected in it, from its first one (the topmost
 * for Down): A2:B4 and A7:B9 down fill A7:B9 from A2:B2 too, A2:B4 and B3:C5 fill B3:B5 from B2 and C4:C5 from C3. A
 * line with one cell selected in it fills from the cell before, as one area does. One step, as fillCopy's.
 */
export function fillCopyAreas(wb, si, areas, dir) {
  if (areas.length < 2) return fillCopy(wb, si, areas[0], dir);
  const sheet = wb.sheet(si);
  const rows = dir === 'down' || dir === 'up';
  const back = dir === 'up' || dir === 'left';
  // (Each line's cells selected, along the way the fill goes: { line: Set of positions }.)
  const lines = new Map();
  for (const a of areas) {
    for (let r = a.r0; r <= a.r1; r++) {
      for (let c = a.c0; c <= a.c1; c++) {
        const [line, at] = rows ? [c, r] : [r, c];
        if (!lines.has(line)) lines.set(line, new Set());
        lines.get(line).add(at);
      }
    }
  }
  const out = { si, contents: [], styles: [], merges: sheet.merges };
  for (const [line, set] of lines) {
    const list = [...set].sort((x, y) => x - y);
    let src = back ? list.at(-1) : list[0];
    if (list.length === 1) { if (back || src === 0) continue; src -= 1; }
    for (const at of list) {
      if (at === src) continue;
      const [r, c] = rows ? [at, line] : [line, at];
      const [fr, fc] = rows ? [src, line] : [line, src];
      out.contents.push(copyOf(sheet, fr, fc, r - fr, c - fc, { r, c }));
      out.styles.push({ r, c, s: sheet.cells.get(fr)?.get(fc)?.s || 0 });
    }
  }
  return [out];
}
