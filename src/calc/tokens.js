// Formulas as tokens: what the parser reads, and what moves a formula to another cell (a shared formula's other cells,
// a formula copied, rows inserted, a sheet renamed). The file's form and the typed one are in fileform.js.
import { colIndex, colName, MAX_ROWS, MAX_COLS } from '../ooxml/refs.js';
import { PREFIXED, fileName } from './prefixed.js';

/** Error values as Excel writes them. */
export const ERROR_CODES = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!',
  '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!'];

export { PREFIXED, fileName };

export class FormulaError extends Error {
  constructor(message, pos) {
    super(message);
    this.pos = pos;
  }
}

// What ends an unquoted sheet name.
const STOP = String.raw`\s!'"(),;:{}\[\]=<>+\-*/^&%#$`;
const RE_BOOK = /\[(\d+)\]/y;
const RE_SHEET = new RegExp(`([^${STOP}0-9.][^${STOP}]*)!`, 'y');
// (A 3D reference's sheets, Jan:Apr!, unless the first reads as a cell: S1:S4!A1 is S1 to S4!A1, as Excel reads it.)
const RE_SHEETS = new RegExp(`([^${STOP}0-9.][^${STOP}]*):([^${STOP}0-9.][^${STOP}]*)!`, 'y');
const CELL = '(\\$?)([A-Za-z]{1,3})(\\$?)([0-9]{1,7})';
const RE_AREA = new RegExp(`${CELL}(?::${CELL})?`, 'y');
const RE_COLS = /(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/y;
const RE_ROWS = /(\$?)([0-9]{1,7}):(\$?)([0-9]{1,7})/y;
const RE_NUMBER = /(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?/y;
const RE_IDENT = /[A-Za-z_\\][A-Za-z0-9_.?\\]*/y;
// (What can't follow a reference: "LOG10(" is a function and "A1B" a name.)
const RE_AFTER_REF = /[A-Za-z0-9_.?\\(!]/;

const at = (re, text, i) => { re.lastIndex = i; return re.exec(text); };
/** A quoted sheet name at i ('Q1 ''24'!A1): its text as written, '' doubled, and its length with the quotes and the
 *  "!"; null if there is none. A loop rather than a regular expression, so a long run of quotes can't make it slow
 *  (code scanning, js/polynomial-redos). */
function quotedSheetAt(text, i) {
  if (text[i] !== "'") return null;
  let j = i + 1;
  while (j < text.length) {
    if (text[j] === "'") {
      if (text[j + 1] !== "'") break;
      j += 2;
    } else j++;
  }
  if (j >= text.length || j === i + 1 || text[j + 1] !== '!') return null;
  return { raw: text.slice(i + 1, j), length: j + 2 - i };
}
/** Where the brackets opened at i close (a table reference's), nested ones and ' escapes passed over; -1 if never. */
function bracketsEnd(text, i) {
  let depth = 0;
  for (let j = i; j < text.length; j++) {
    if (text[j] === "'") { j++; continue; }
    if (text[j] === '[') depth++;
    else if (text[j] === ']' && --depth === 0) return j + 1;
  }
  return -1;
}
const okRow = (r) => r >= 0 && r < MAX_ROWS;
const okCol = (c) => c >= 0 && c < MAX_COLS;

/** A reference starting at text[i] (with its sheet, `Sheet1!` or `'My sheet'!`, and an external book's `[1]`), or null. */
function refAt(text, i) {
  let j = i;
  let sheet = null;
  let book = null;
  let m = at(RE_BOOK, text, j);
  if (m) { book = m[1]; j += m[0].length; }
  const quoted = quotedSheetAt(text, j);
  if (quoted) {
    sheet = quoted.raw.replace(/''/g, "'");
    j += quoted.length;
    const inner = /^\[(\d+)\](.*)$/.exec(sheet);
    if (inner) { book = inner[1]; sheet = inner[2]; }
  } else if ((m = at(RE_SHEETS, text, j)) && !/^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$/.test(m[1])) {
    sheet = `${m[1]}:${m[2]}`;
    j += m[0].length;
  } else if ((m = at(RE_SHEET, text, j))) {
    sheet = m[1];
    j += m[0].length;
  }
  // (Sheet names can't hold ":": one in the name, quoted or not, is a 3D reference's first and last sheets.)
  const sheets = sheet != null && sheet.includes(':') ? sheet.split(':') : null;
  if (book != null && sheet == null) return null;
  const prefix = text.slice(i, j);
  if (sheet != null && text.startsWith('#REF!', j)) return { end: j + 5, sheet, sheets, book, prefix, area: null };
  let area = null;
  let end = j;
  if ((m = at(RE_AREA, text, j))) {
    const r0 = Number(m[4]) - 1;
    const c0 = colIndex(m[2]);
    const two = m[6] != null;
    const r1 = two ? Number(m[8]) - 1 : r0;
    const c1 = two ? colIndex(m[6]) : c0;
    if (okRow(r0) && okRow(r1) && okCol(c0) && okCol(c1)) {
      area = { kind: two ? 'area' : 'cell', r0, c0, r1, c1, ac0: !!m[1], ar0: !!m[3], ac1: two ? !!m[5] : !!m[1], ar1: two ? !!m[7] : !!m[3] };
      end = j + m[0].length;
    }
  }
  if (!area && (m = at(RE_COLS, text, j))) {
    const c0 = colIndex(m[2]);
    const c1 = colIndex(m[4]);
    if (okCol(c0) && okCol(c1)) {
      area = { kind: 'cols', r0: 0, c0, r1: MAX_ROWS - 1, c1, ac0: !!m[1], ac1: !!m[3], ar0: true, ar1: true };
      end = j + m[0].length;
    }
  }
  if (!area && (m = at(RE_ROWS, text, j))) {
    const r0 = Number(m[2]) - 1;
    const r1 = Number(m[4]) - 1;
    if (okRow(r0) && okRow(r1)) {
      area = { kind: 'rows', r0, c0: 0, r1, c1: MAX_COLS - 1, ar0: !!m[1], ar1: !!m[3], ac0: true, ac1: true };
      end = j + m[0].length;
    }
  }
  if (!area || (end < text.length && RE_AFTER_REF.test(text[end]))) return null;
  return { end, sheet, sheets, book, prefix, area };
}

/** A number written in a formula, as Excel keeps it: to 15 significant digits, the rest cut off (=1234567890123456
 *  is 1234567890123450). */
function literal(s) {
  const m = /^(\d*)(?:\.(\d*))?(?:[eE]([+-]?\d+))?$/.exec(s);
  const all = (m[1] + (m[2] || '')).replace(/^0+/, '');
  // (Trailing zeros counted by a loop: /0+$/ on a long run of zeros followed by another digit is slow, code scanning.)
  let end = all.length;
  while (end > 0 && all[end - 1] === '0') end--;
  if (end <= 15) return Number(s);
  // (Significant digits counted from the first non-zero one; the point stays where it was.)
  const lead = (m[1] + (m[2] || '')).length - all.length;
  const digits = (m[1] + (m[2] || '')).slice(0, lead + 15);
  const intLen = m[1].length;
  const kept = digits.length > intLen ? `${digits.slice(0, intLen)}.${digits.slice(intLen)}` : digits.padEnd(intLen, '0');
  return Number(`${kept || '0'}${m[3] ? `e${m[3]}` : ''}`);
}

// (What an intersection's space comes between: a reference, name or ")" before it, a reference, name, function or
// "(" after.)
const ENDS = new Set(['ref', 'name', 'struct', ')']);
const STARTS = new Set(['ref', 'name', 'struct', 'func', '(']);

/** A formula (without its "=") as tokens { type, text, start, end, ... }: num, str, bool, err (with `value`), ref (with
 *  `sheet`, `book`, `prefix` and `area`, or area null for Sheet1!#REF!), func (a name followed by "("), name, struct (a
 *  table reference, Fruit[Qty] or [@Qty]: `table` (null when unnamed) and `spec`, the brackets' text), op, punct.
 *  Spaces are skipped, but for Excel's intersection operator: spaces between two references (or names, or a ")"
 *  and a reference) are an op " " (=SUM(A1:C3 B2:D4)). */
export function tokenize(text) {
  const out = [];
  const n = text.length;
  let i = 0;
  let gap = -1;
  const push = (type, end, extra = {}) => {
    const prev = out[out.length - 1];
    if (gap >= 0 && prev && ENDS.has(prev.type === 'punct' ? prev.text : prev.type) && STARTS.has(type === 'punct' ? text[i] : type)) {
      out.push({ type: 'op', text: ' ', start: gap, end: gap + 1 });
    }
    gap = -1;
    out.push({ type, text: text.slice(i, end), start: i, end, ...extra });
    i = end;
  };
  while (i < n) {
    const ch = text[i];
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t') { if (gap < 0) gap = i; i++; continue; }
    if (ch === '"') {
      let j = i + 1;
      let s = '';
      for (;;) {
        if (j >= n) throw new FormulaError('Text is missing its closing quote', i);
        if (text[j] === '"') {
          if (text[j + 1] === '"') { s += '"'; j += 2; continue; }
          break;
        }
        s += text[j++];
      }
      push('str', j + 1, { value: s });
      continue;
    }
    if (ch === '#') {
      const code = ERROR_CODES.find((e) => text.slice(i, i + e.length).toUpperCase() === e);
      if (!code) throw new FormulaError('Not an error value', i);
      push('err', i + code.length, { value: code });
      continue;
    }
    if (ch === '[') {
      const end = bracketsEnd(text, i);
      if (end < 0) throw new FormulaError('A table reference is missing its closing bracket', i);
      push('struct', end, { table: null, spec: text.slice(i + 1, end - 1) });
      continue;
    }
    const ref = refAt(text, i);
    // (A cell followed by # is the range its formula spills into: D1#.)
    if (ref && text[ref.end] === '#' && ref.area?.kind === 'cell' && !ERROR_CODES.some((e) => text.slice(ref.end, ref.end + e.length).toUpperCase() === e)) {
      push('ref', ref.end + 1, { ...ref, end: ref.end + 1, spill: true });
      continue;
    }
    if (ref) { push('ref', ref.end, ref); continue; }
    let m = at(RE_NUMBER, text, i);
    if (m) { push('num', i + m[0].length, { value: literal(m[0]) }); continue; }
    if ((m = at(RE_IDENT, text, i))) {
      let j = i + m[0].length;
      while (text[j] === ' ') j++;
      const upper = m[0].toUpperCase();
      if (text[i + m[0].length] === '[') {
        const end = bracketsEnd(text, i + m[0].length);
        if (end < 0) throw new FormulaError('A table reference is missing its closing bracket', i);
        push('struct', end, { table: m[0], spec: text.slice(i + m[0].length + 1, end - 1) });
      } else if (text[j] === '(') push('func', i + m[0].length);
      else if (upper === 'TRUE' || upper === 'FALSE') push('bool', i + m[0].length, { value: upper === 'TRUE' });
      else push('name', i + m[0].length);
      continue;
    }
    const two = text.slice(i, i + 2);
    if (two === '<>' || two === '<=' || two === '>=') { push('op', i + 2); continue; }
    if ('+-*/^&=<>%:@'.includes(ch)) { push('op', i + 1); continue; }
    if ('(),;{}'.includes(ch)) { push('punct', i + 1); continue; }
    throw new FormulaError(`Unexpected "${ch}"`, i);
  }
  return out;
}

/** An area as A1 text: "$B$2", "A1:C3", "A:A", "2:5". */
export function areaText(a) {
  const cell = (c, r, ac, ar) => `${ac ? '$' : ''}${colName(c)}${ar ? '$' : ''}${r + 1}`;
  if (a.kind === 'cols') return `${a.ac0 ? '$' : ''}${colName(a.c0)}:${a.ac1 ? '$' : ''}${colName(a.c1)}`;
  if (a.kind === 'rows') return `${a.ar0 ? '$' : ''}${a.r0 + 1}:${a.ar1 ? '$' : ''}${a.r1 + 1}`;
  const first = cell(a.c0, a.r0, a.ac0, a.ar0);
  return a.kind === 'cell' ? first : `${first}:${cell(a.c1, a.r1, a.ac1, a.ar1)}`;
}

/** An area moved by (dr, dc): its relative rows and columns move, its $ ones stay. Null if it would leave the sheet. */
export function shiftArea(a, dr, dc) {
  const out = { ...a };
  if (a.kind !== 'cols') {
    if (!a.ar0) out.r0 += dr;
    if (!a.ar1) out.r1 += dr;
  }
  if (a.kind !== 'rows') {
    if (!a.ac0) out.c0 += dc;
    if (!a.ac1) out.c1 += dc;
  }
  return okRow(out.r0) && okRow(out.r1) && okCol(out.c0) && okCol(out.c1) ? out : null;
}

/** Rewrite the references of a formula, keeping everything else as written. */
export function mapRefs(text, fn) {
  let out = '';
  let last = 0;
  for (const t of tokenize(text)) {
    const next = fn(t);
    if (next == null) continue;
    out += text.slice(last, t.start) + next;
    last = t.end;
  }
  return out + text.slice(last);
}

/** A formula as it reads from a cell (dr, dc) away: a shared formula's other cells, or a formula copied. References that
 *  would leave the sheet become #REF!, as in Excel. */
export function shiftFormula(text, dr, dc) {
  if (!dr && !dc) return text;
  return mapRefs(text, (t) => {
    if (t.type !== 'ref' || !t.area) return null;
    const moved = shiftArea(t.area, dr, dc);
    return t.prefix + (moved ? areaText(moved) : '#REF!');
  });
}

/**
 * A formula copied with Paste Special's Transpose, as Excel rewrites it: `from` the cells copied ({ sheet, r0, c0, r1,
 * c1 }), `src` the formula's cell there and `dst` where it goes ({ r, c }), `to` where `from`'s top-left cell goes. A
 * cell of the copy that a reference names, $ or not, becomes the cell it went to (=$A$1+B1 in B2, A1:C3 to E1, is
 * =$E$1+E2 in F2); a reference outside the copy keeps its $ parts, and its relative parts are the same distance away
 * across instead of down (=D9 in B2 is =M4 in F2). Off the sheet, #REF!.
 */
export function transposeFormula(text, { here, from, src, dst, to }) {
  const same = (a, b) => String(a).toLowerCase() === String(b).toLowerCase();
  const corner = (r, c, absR, absC, own) => {
    if (own && r >= from.r0 && r <= from.r1 && c >= from.c0 && c <= from.c1) return [to.r + (c - from.c0), to.c + (r - from.r0)];
    return [absR ? r : dst.r + (c - src.c), absC ? c : dst.c + (r - src.r)];
  };
  try {
    return mapRefs(text, (t) => {
      if (t.type !== 'ref' || t.book != null || !t.area || t.area.kind === 'rows' || t.area.kind === 'cols') return null;
      const a = t.area;
      const own = same(t.sheet ?? here, from.sheet);
      const [r0, c0] = corner(a.r0, a.c0, a.ar0, a.ac0, own);
      const [r1, c1] = corner(a.r1, a.c1, a.ar1, a.ac1, own);
      const next = { ...a, r0: Math.min(r0, r1), r1: Math.max(r0, r1), c0: Math.min(c0, c1), c1: Math.max(c0, c1) };
      if (!(okRow(next.r0) && okRow(next.r1) && okCol(next.c0) && okCol(next.c1))) return `${t.prefix}#REF!`;
      return t.prefix + areaText(next);
    });
  } catch {
    return text;
  }
}

/** A row or column index after `n` lines were inserted at `at` (n > 0) or deleted from `at` (n < 0: at to at - n - 1);
 *  null when deleted. */
export function shiftLine(i, at, n) {
  if (n > 0) return i >= at ? i + n : i;
  const last = at - n - 1;
  return i < at ? i : i > last ? i + n : null;
}

/** A span of lines [a, b] after an insert or delete (see shiftLine), as Excel changes a range: an insert at or before
 *  its end moves the end (inside, the range grows); a delete shrinks it, and takes it away when every line goes. Null
 *  then. */
export function shiftSpan(a, b, at, n) {
  if (n > 0) return [shiftLine(a, at, n), shiftLine(b, at, n)];
  const last = at - n - 1;
  if (a >= at && b <= last) return null;
  return [a < at ? a : a > last ? a + n : at, b > last ? b + n : b < at ? b : at - 1];
}

/** An area ({ r0, c0, r1, c1 }) after cells of band `band` were shifted (see shiftRefs): undefined when it stays as
 *  it was, null when all of it was deleted, the moved, grown or shrunk area otherwise. Whole rows and columns stay. */
export function bandArea(a, axis, at, n, [b0, b1]) {
  if (a.kind === 'rows' || a.kind === 'cols') return undefined;
  const rows = axis === 'row';
  const [p0, p1] = rows ? [a.c0, a.c1] : [a.r0, a.r1];
  const [s0, s1] = rows ? [a.r0, a.r1] : [a.c0, a.c1];
  let next;
  if (p0 >= b0 && p1 <= b1) {
    const span = shiftSpan(s0, s1, at, n);
    if (!span) return null;
    if (!(rows ? okRow(span[1]) : okCol(span[1]))) return null;
    next = rows ? { ...a, r0: span[0], r1: span[1] } : { ...a, c0: span[0], c1: span[1] };
  } else if (n < 0 && s0 >= at && s1 <= at - n - 1 && b0 <= p1 && b1 >= p0) {
    // (The deleted cells take one end off a range they cross; in its middle they leave it as it was.)
    let q0 = p0;
    let q1 = p1;
    if (b0 <= p0) q0 = b1 + 1; else if (b1 >= p1) q1 = b0 - 1; else return undefined;
    next = rows ? { ...a, c0: q0, c1: q1 } : { ...a, r0: q0, r1: q1 };
  } else return undefined;
  if (next.r0 === a.r0 && next.r1 === a.r1 && next.c0 === a.c0 && next.c1 === a.c1) return undefined;
  return next;
}

/**
 * A formula after rows or columns of sheet `sheet` were inserted or deleted (`axis` 'row' or 'col'; `at`, `n` as in
 * shiftLine), as Excel rewrites it: references move, $ or not; a range grows or shrinks; a reference whose cells were
 * all deleted becomes #REF! (its sheet kept: Sheet1!#REF!). `here`: the formula's own sheet (null for a defined name,
 * whose references name their sheet). Text that doesn't read as a formula comes back as it was.
 *
 * With `band` ([first, last] columns for axis 'row', rows for 'col'), only the cells of those columns (or rows) were
 * shifted (Excel's Insert and Delete with Shift cells down, right, up or left): a reference moves as above when its
 * columns (rows) are within the band and stays otherwise, whole rows and columns included; but a range all of whose
 * shifted-axis lines were deleted, with the band covering one end of it, loses that end (B3:D4 is D3:D4 once B3:C4 is
 * deleted upwards).
 */
export function shiftRefs(text, { here, sheet, axis, at, n, band = null }) {
  const same = (a, b) => a != null && b != null && String(a).toLowerCase() === String(b).toLowerCase();
  const rows = axis === 'row';
  try {
    return mapRefs(text, (t) => {
      if (t.type !== 'ref' || t.book != null || !t.area || !same(t.sheet ?? here, sheet)) return null;
      const a = t.area;
      if (rows ? a.kind === 'cols' : a.kind === 'rows') return null;
      if (band) {
        const next = bandArea(a, axis, at, n, band);
        if (next === undefined) return null;
        if (next === null) return `${t.prefix}#REF!`;
        return t.prefix + areaText(next);
      }
      const span = rows ? shiftSpan(a.r0, a.r1, at, n) : shiftSpan(a.c0, a.c1, at, n);
      if (!span) return `${t.prefix}#REF!`;
      const next = rows ? { ...a, r0: span[0], r1: span[1] } : { ...a, c0: span[0], c1: span[1] };
      if ((rows ? !okRow(next.r1) : !okCol(next.c1))) return `${t.prefix}#REF!`;
      if (next.r0 === a.r0 && next.r1 === a.r1 && next.c0 === a.c0 && next.c1 === a.c1) return null;
      return t.prefix + areaText(next);
    });
  } catch {
    return text;
  }
}

/**
 * A formula after cells were cut and pasted (Excel's move), as Excel rewrites it. `here`: the sheet the formula was on;
 * `home`: the sheet it is on now (moved with the cells, or `here`); `from`: { sheet, r0, c0, r1, c1 }, the cells cut;
 * `to`: { sheet, dr, dc }, where they went. A reference wholly inside the cut cells moves with them, $ or not; a
 * reference to a single cell that was pasted over becomes #REF!; other references stay, with their sheet named when
 * the formula has moved to another sheet. Text that doesn't read as a formula comes back as it was.
 * A range only partly cut becomes the cells of it left behind and those of it moved (on the same sheet), when they
 * make a rectangle, else stays (measured over COM, B2:C3 cut: =SUM(B2:B5) is =SUM(B3:B5) for B3, =SUM(B4:B5) for
 * another sheet, as it was for E5; A1:A3 with A1 moved onto A3, A2:A3).
 */
export function moveRefs(text, { here, home = here, from, to }) {
  const same = (a, b) => String(a).toLowerCase() === String(b).toLowerCase();
  const over = { r0: from.r0 + to.dr, c0: from.c0 + to.dc, r1: from.r1 + to.dr, c1: from.c1 + to.dc };
  const inside = (a, r) => a.r0 >= r.r0 && a.r1 <= r.r1 && a.c0 >= r.c0 && a.c1 <= r.c1;
  try {
    return mapRefs(text, (t) => {
      if (t.type !== 'ref' || t.book != null || !t.area) return null;
      const sheet = t.sheet ?? here;
      const body = text.slice(t.start + t.prefix.length, t.end);
      const named = (name, area) => (same(name, home) && t.sheet == null ? '' : sheetPrefix(name)) + area;
      if (same(sheet, from.sheet) && t.area.kind !== 'cols' && t.area.kind !== 'rows' && inside(t.area, from)) {
        const a = { ...t.area, r0: t.area.r0 + to.dr, r1: t.area.r1 + to.dr, c0: t.area.c0 + to.dc, c1: t.area.c1 + to.dc };
        return named(to.sheet, areaText(a));
      }
      if (same(sheet, from.sheet) && t.area.kind === 'area' && overlap(t.area, from)) {
        const kept = partlyMoved(t.area, from, same(from.sheet, to.sheet) ? to : null);
        if (kept) return (t.sheet == null && !same(here, home) ? sheetPrefix(here) : text.slice(t.start, t.start + t.prefix.length)) + areaText({ ...t.area, ...kept });
      }
      if (same(sheet, to.sheet) && t.area.kind === 'cell' && inside(t.area, over)) return `${text.slice(t.start, t.start + t.prefix.length)}#REF!`;
      if (t.sheet == null && !same(here, home)) return sheetPrefix(here) + body;
      return null;
    });
  } catch {
    return text;
  }
}

const overlap = (a, b) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;

/** The cells of rectangles (overlaps counted once). */
function unionArea(rects) {
  const rs = [...new Set(rects.flatMap((x) => [x.r0, x.r1 + 1]))].sort((a, b) => a - b);
  const cs = [...new Set(rects.flatMap((x) => [x.c0, x.c1 + 1]))].sort((a, b) => a - b);
  let n = 0;
  for (let i = 0; i + 1 < rs.length; i++) {
    for (let j = 0; j + 1 < cs.length; j++) {
      if (rects.some((x) => rs[i] >= x.r0 && rs[i] <= x.r1 && cs[j] >= x.c0 && cs[j] <= x.c1)) n += (rs[i + 1] - rs[i]) * (cs[j + 1] - cs[j]);
    }
  }
  return n;
}

/** Range `a`, partly inside the cut cells `cut`, after the cut: the cells of it left behind and those of it moved by
 *  `to` (null: to another sheet), as { r0, c0, r1, c1 } when they make one rectangle other than `a`, else null. */
function partlyMoved(a, cut, to) {
  const k = { r0: Math.max(a.r0, cut.r0), c0: Math.max(a.c0, cut.c0), r1: Math.min(a.r1, cut.r1), c1: Math.min(a.c1, cut.c1) };
  const left = [];
  if (a.r0 < k.r0) left.push({ r0: a.r0, c0: a.c0, r1: k.r0 - 1, c1: a.c1 });
  if (k.r1 < a.r1) left.push({ r0: k.r1 + 1, c0: a.c0, r1: a.r1, c1: a.c1 });
  if (a.c0 < k.c0) left.push({ r0: k.r0, c0: a.c0, r1: k.r1, c1: k.c0 - 1 });
  if (k.c1 < a.c1) left.push({ r0: k.r0, c0: k.c1 + 1, r1: k.r1, c1: a.c1 });
  const pieces = to ? [...left, { r0: k.r0 + to.dr, c0: k.c0 + to.dc, r1: k.r1 + to.dr, c1: k.c1 + to.dc }] : left;
  if (!pieces.length) return null;
  const box = { r0: Math.min(...pieces.map((x) => x.r0)), c0: Math.min(...pieces.map((x) => x.c0)), r1: Math.max(...pieces.map((x) => x.r1)), c1: Math.max(...pieces.map((x) => x.c1)) };
  if (box.r0 < 0 || box.c0 < 0 || unionArea(pieces) !== (box.r1 - box.r0 + 1) * (box.c1 - box.c0 + 1)) return null;
  return box.r0 === a.r0 && box.c0 === a.c0 && box.r1 === a.r1 && box.c1 === a.c1 ? null : box;
}

/** A sheet name as a formula writes it before a reference: plain (Data!) when it reads as a name, quoted otherwise
 *  ('My Data'!, 'A1'!, '2024'!, 'R1C1'!, 'TRUE'!; XFE1, past the last column, stays plain, as Excel has it), an
 *  apostrophe doubled. */
export function sheetPrefix(name) {
  const cell = /^([A-Za-z]{1,3})([0-9]+)$/.exec(name);
  const plain = /^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u.test(name) && !(cell && okCol(colIndex(cell[1])) && okRow(Number(cell[2]) - 1)) &&
    !/^(R[0-9]*|C[0-9]*|R[0-9]*C[0-9]*)$/i.test(name) && !/^(TRUE|FALSE)$/i.test(name);
  return plain ? `${name}!` : `'${name.replace(/'/g, "''")}'!`;
}

/** A 3D reference's sheets as a formula writes them: Jan:Apr!, quoted as one when either name needs it ('Feb:My Mar'!). */
function spanPrefix(a, b) {
  const plain = sheetPrefix(a) === `${a}!` && sheetPrefix(b) === `${b}!`;
  return plain ? `${a}:${b}!` : `'${`${a}:${b}`.replace(/'/g, "''")}'!`;
}

/** A formula with each 3D reference's sheets ([first, last] as written) passed through `fn`, which gives a new pair,
 *  one sheet's name (a plain reference), false (#REF!) or null (as it was). Text that isn't a formula is kept. */
export function spanSheetRefs(text, fn) {
  try {
    return mapRefs(text, (t) => {
      if (t.type !== 'ref' || t.book != null || !t.sheets) return null;
      const next = fn(t.sheets);
      if (next == null) return null;
      if (next === false) return '#REF!';
      const body = text.slice(t.start + t.prefix.length, t.end);
      return (typeof next === 'string' ? sheetPrefix(next) : spanPrefix(next[0], next[1])) + body;
    });
  } catch {
    return text;
  }
}

/** A formula with its references to sheet `from` (case ignored, this workbook's only) pointing at sheet `to`, or, with
 *  `to` null, become #REF! (a deleted sheet: Excel writes the whole reference as #REF!). Text that doesn't read as a
 *  formula comes back as it was. */
export function renameSheetRefs(text, from, to) {
  const was = from.toLowerCase();
  try {
    return mapRefs(text, (t) => {
      // (A 3D reference: an end renamed is renamed in it; one deleted is left as it is, worked out as #REF!.)
      if (t.type === 'ref' && t.book == null && t.sheets) {
        const ends = t.sheets.map((s) => (s.toLowerCase() === was ? to : s));
        if (to == null || ends.every((s, k) => s === t.sheets[k])) return null;
        return spanPrefix(ends[0], ends[1]) + text.slice(t.start + t.prefix.length, t.end);
      }
      if (t.type !== 'ref' || t.book != null || t.sheet == null || t.sheet.toLowerCase() !== was) return null;
      return to == null ? '#REF!' : sheetPrefix(to) + text.slice(t.start + t.prefix.length, t.end);
    });
  } catch {
    return text;
  }
}
