// Formulas as tokens: what the parser reads, and what moves a formula to another cell (a shared formula's other cells)
// or turns it between the file's form (_xlfn.CONCAT) and what a person types and sees (CONCAT).
import { colIndex, colName, MAX_ROWS, MAX_COLS } from '../ooxml/refs.js';

/** Error values as Excel writes them. */
export const ERROR_CODES = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!',
  '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!'];

/** Functions newer than Excel 2007, which the file names with a prefix (`_xlfn.CONCAT`) so that older versions show
 *  #NAME? instead of misreading them. Only the ones the engine knows: others typed are kept as typed. */
export const PREFIXED = new Set(['CONCAT', 'TEXTJOIN', 'IFS', 'SWITCH', 'MAXIFS', 'MINIFS', 'XLOOKUP', 'XMATCH', 'IFNA',
  'DAYS', 'XOR', 'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P', 'CEILING.MATH', 'FLOOR.MATH', 'ISOWEEKNUM', 'PERCENTILE.INC', 'PERCENTILE.EXC',
  'QUARTILE.INC', 'QUARTILE.EXC', 'RANK.EQ', 'RANK.AVG', 'MODE.SNGL', 'TEXTBEFORE', 'TEXTAFTER', 'NUMBERVALUE', 'UNICODE', 'UNICHAR',
  'SEQUENCE', 'SORTBY', 'UNIQUE', 'LET', 'LAMBDA', 'ANCHORARRAY', 'SINGLE', 'SORT', 'FILTER',
  'TAKE', 'DROP', 'VSTACK', 'HSTACK', 'CHOOSECOLS', 'CHOOSEROWS', 'TOCOL', 'TOROW', 'WRAPROWS', 'WRAPCOLS', 'EXPAND', 'TEXTSPLIT', 'ISOMITTED', 'MAP',
  'REDUCE', 'SCAN', 'BYROW', 'BYCOL', 'MAKEARRAY', 'VALUETOTEXT', 'ARRAYTOTEXT', 'NORM.DIST', 'NORM.S.DIST', 'NORM.INV', 'NORM.S.INV', 'EXPON.DIST',
  'POISSON.DIST', 'BINOM.DIST', 'COVARIANCE.P', 'COVARIANCE.S', 'FORECAST.LINEAR', 'SEC', 'CSC', 'COT', 'ARABIC', 'BASE', 'DECIMAL', 'ISFORMULA',
  'FORMULATEXT', 'SHEET', 'SHEETS', 'NETWORKDAYS.INTL', 'WORKDAY.INTL']);
/** Of those, the ones the file names with a second prefix (`_xlfn._xlws.SORT`), as Excel 365 writes them. */
const WORKSHEET_PREFIXED = new Set(['SORT', 'FILTER']);
/** A newer function's name as the file has it. */
export const fileName = (name) => (WORKSHEET_PREFIXED.has(name) ? `_xlfn._xlws.${name}` : PREFIXED.has(name) ? `_xlfn.${name}` : name);

export class FormulaError extends Error {
  constructor(message, pos) {
    super(message);
    this.pos = pos;
  }
}

// What ends an unquoted sheet name.
const STOP = String.raw`\s!'"(),;:{}\[\]=<>+\-*/^&%#$`;
const RE_BOOK = /\[(\d+)\]/y;
const RE_QSHEET = /'((?:[^']|'')+)'!/y;
const RE_SHEET = new RegExp(`([^${STOP}0-9.][^${STOP}]*)!`, 'y');
const CELL = '(\\$?)([A-Za-z]{1,3})(\\$?)([0-9]{1,7})';
const RE_AREA = new RegExp(`${CELL}(?::${CELL})?`, 'y');
const RE_COLS = /(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/y;
const RE_ROWS = /(\$?)([0-9]{1,7}):(\$?)([0-9]{1,7})/y;
const RE_NUMBER = /(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?/y;
const RE_IDENT = /[A-Za-z_\\][A-Za-z0-9_.?\\]*/y;
// (What can't follow a reference: "LOG10(" is a function and "A1B" a name.)
const RE_AFTER_REF = /[A-Za-z0-9_.?\\(!]/;

const at = (re, text, i) => { re.lastIndex = i; return re.exec(text); };
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
  if ((m = at(RE_QSHEET, text, j))) {
    sheet = m[1].replace(/''/g, "'");
    j += m[0].length;
    const inner = /^\[(\d+)\](.*)$/.exec(sheet);
    if (inner) { book = inner[1]; sheet = inner[2]; }
  } else if ((m = at(RE_SHEET, text, j))) {
    sheet = m[1];
    j += m[0].length;
  }
  if (book != null && sheet == null) return null;
  const prefix = text.slice(i, j);
  if (sheet != null && text.startsWith('#REF!', j)) return { end: j + 5, sheet, book, prefix, area: null };
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
  return { end, sheet, book, prefix, area };
}

/** A number written in a formula, as Excel keeps it: to 15 significant digits, the rest cut off (=1234567890123456
 *  is 1234567890123450). */
function literal(s) {
  const m = /^(\d*)(?:\.(\d*))?(?:[eE]([+-]?\d+))?$/.exec(s);
  const all = (m[1] + (m[2] || '')).replace(/^0+/, '');
  if (all.replace(/0+$/, '').length <= 15) return Number(s);
  // (Significant digits counted from the first non-zero one; the point stays where it was.)
  const lead = (m[1] + (m[2] || '')).length - all.length;
  const digits = (m[1] + (m[2] || '')).slice(0, lead + 15);
  const intLen = m[1].length;
  const kept = digits.length > intLen ? `${digits.slice(0, intLen)}.${digits.slice(intLen)}` : digits.padEnd(intLen, '0');
  return Number(`${kept || '0'}${m[3] ? `e${m[3]}` : ''}`);
}

/** A formula (without its "=") as tokens { type, text, start, end, ... }: num, str, bool, err (with `value`), ref (with
 *  `sheet`, `book`, `prefix` and `area`, or area null for Sheet1!#REF!), func (a name followed by "("), name, struct (a
 *  table reference, Fruit[Qty] or [@Qty]: `table` (null when unnamed) and `spec`, the brackets' text), op, punct.
 *  Spaces are skipped (Excel's intersection operator, a space between references, isn't read). */
export function tokenize(text) {
  const out = [];
  const n = text.length;
  let i = 0;
  const push = (type, end, extra = {}) => { out.push({ type, text: text.slice(i, end), start: i, end, ...extra }); i = end; };
  while (i < n) {
    const ch = text[i];
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t') { i++; continue; }
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
function mapRefs(text, fn) {
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
      if (same(sheet, to.sheet) && t.area.kind === 'cell' && inside(t.area, over)) return '#REF!';
      if (t.sheet == null && !same(here, home)) return sheetPrefix(here) + body;
      return null;
    });
  } catch {
    return text;
  }
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

/** A formula with its references to sheet `from` (case ignored, this workbook's only) pointing at sheet `to`, or, with
 *  `to` null, become #REF! (a deleted sheet: Excel writes the whole reference as #REF!). Text that doesn't read as a
 *  formula comes back as it was. */
export function renameSheetRefs(text, from, to) {
  const was = from.toLowerCase();
  try {
    return mapRefs(text, (t) => {
      if (t.type !== 'ref' || t.book != null || t.sheet == null || t.sheet.toLowerCase() !== was) return null;
      return to == null ? '#REF!' : sheetPrefix(to) + text.slice(t.start + t.prefix.length, t.end);
    });
  } catch {
    return text;
  }
}

/** Where the bracket opened at token k (a "(" punct, or the one after a func) closes: the index of its ")". */
function closing(toks, k) {
  let depth = 0;
  for (let j = k; j < toks.length; j++) {
    if (toks[j].type === 'punct' && toks[j].text === '(') depth++;
    else if (toks[j].type === 'punct' && toks[j].text === ')' && --depth === 0) return j;
  }
  return toks.length - 1;
}
/** The text with some tokens' spans replaced ([{ from, to (token indexes, inclusive), text }]). */
function splice(text, toks, edits) {
  let out = '';
  let at = 0;
  for (const e of edits.sort((a, b) => toks[a.from].start - toks[b.from].start)) {
    out += text.slice(at, toks[e.from].start) + e.text;
    at = toks[e.to].end;
  }
  return out + text.slice(at);
}

/** A typed formula's dynamic-array parts in the file's form, as Excel 365 writes them: D1# as _xlfn.ANCHORARRAY(D1),
 *  @x as _xlfn.SINGLE(x), and LET's and LAMBDA's names as _xlpm.name. */
function dynamicFileForm(typed) {
  const toks = tokenize(typed);
  const edits = [];
  const params = new Map(); // token index of a name -> its file form
  toks.forEach((t, k) => {
    const name = t.type === 'func' ? t.text.toUpperCase().replace(/^(_XL(FN|WS)\.)+/, '') : null;
    if (name !== 'LET' && name !== 'LAMBDA') return;
    // (Its arguments, split at its own commas: LET's names are every other one but the last, LAMBDA's all but the last.)
    const end = closing(toks, k + 1);
    const args = [[]];
    let depth = 0;
    for (let j = k + 2; j < end; j++) {
      const x = toks[j];
      if (x.type === 'punct' && (x.text === '(' || x.text === '{')) depth++;
      if (x.type === 'punct' && (x.text === ')' || x.text === '}')) depth--;
      if (depth === 0 && x.type === 'punct' && x.text === ',') args.push([]);
      else args[args.length - 1].push(j);
    }
    const names = args.slice(0, -1).filter((_, a) => name === 'LAMBDA' || a % 2 === 0)
      .map((a) => (a.length === 1 && toks[a[0]].type === 'name' ? toks[a[0]].text.replace(/^_xlpm\./i, '').toLowerCase() : null)).filter(Boolean);
    for (let j = k + 2; j < end; j++) if (toks[j].type === 'name' && names.includes(toks[j].text.replace(/^_xlpm\./i, '').toLowerCase())) params.set(j, `_xlpm.${toks[j].text.replace(/^_xlpm\./i, '')}`);
  });
  for (const [k, text] of params) edits.push({ from: k, to: k, text });
  toks.forEach((t, k) => {
    if (t.type === 'ref' && t.spill) edits.push({ from: k, to: k, text: `_xlfn.ANCHORARRAY(${t.text.slice(0, -1)})` });
    if (t.type === 'op' && t.text === '@' && toks[k + 1]) {
      const next = toks[k + 1];
      const end = next.type === 'func' ? closing(toks, k + 2) : next.type === 'punct' && next.text === '(' ? closing(toks, k + 1) : k + 1;
      // (The operand as typed, inner names and # converted too.)
      const inner = dynamicFileForm(typed.slice(next.start, toks[end].end));
      edits.push({ from: k, to: end, text: `_xlfn.SINGLE(${inner})` });
    }
  });
  // (An @ that covers tokens another edit also covers: the @'s operand was converted whole.)
  const kept = edits.filter((e) => !edits.some((o) => o !== e && o.from < e.from && o.to >= e.to));
  return kept.length ? splice(typed, toks, kept) : typed;
}

/** The file's dynamic-array parts as Excel shows them: D1#, @x, and LET's and LAMBDA's names without _xlpm. */
function dynamicShownForm(text) {
  const toks = tokenize(text);
  const edits = [];
  toks.forEach((t, k) => {
    const name = t.type === 'func' ? t.text.toUpperCase().replace(/^(_XL(FN|WS)\.)+/, '') : null;
    if (name === 'ANCHORARRAY' && toks[k + 2]?.type === 'ref' && toks[k + 3]?.text === ')') edits.push({ from: k, to: k + 3, text: `${toks[k + 2].text}#` });
    else if (name === 'SINGLE') {
      const end = closing(toks, k + 1);
      edits.push({ from: k, to: end, text: `@${dynamicShownForm(text.slice(toks[k + 2].start, toks[end - 1].end))}` });
    } else if (t.type === 'name' && /^_xlpm\./i.test(t.text)) edits.push({ from: k, to: k, text: t.text.replace(/^_xlpm\./i, '') });
  });
  const kept = edits.filter((e) => !edits.some((o) => o !== e && o.from <= e.from && o.to >= e.to && (o.from !== e.from || o.to !== e.to)));
  return kept.length ? splice(text, toks, kept) : text;
}

/** A formula from the file as Excel shows it: without the _xlfn. and _xlws. prefixes of newer functions, and with
 *  Excel 365's D1#, @ and LET's names as typed. */
export function displayFormula(text) {
  try {
    const shown = /_xlfn\.(ANCHORARRAY|SINGLE)|_xlpm\./i.test(text) ? dynamicShownForm(text) : text;
    return mapRefs(shown, (t) => (t.type === 'func' && /^_xl(fn|ws)\./i.test(t.text) ? t.text.replace(/^(_xl(fn|ws)\.)+/i, '') : null));
  } catch {
    return text;
  }
}

/** A typed formula (without "=") in the file's form, as Excel takes an entry: function names, references and TRUE and
 *  FALSE in capitals, newer functions prefixed, missing closing brackets added. Throws a FormulaError when it doesn't
 *  read as a formula. */
export function normalizeFormula(entry) {
  const typed = /[#@]|\b(LET|LAMBDA)\s*\(/i.test(entry) ? dynamicFileForm(entry) : entry;
  const toks = tokenize(typed);
  let depth = 0;
  for (const t of toks) {
    if (t.type === 'punct' && t.text === '(') depth++;
    if (t.type === 'func') depth++;
    if (t.type === 'punct' && t.text === ')') depth--;
  }
  // (A function token doesn't include its "(", which comes as its own token: count it once.)
  depth -= toks.filter((t) => t.type === 'func').length;
  let text = mapRefs(typed, (t) => {
    if (t.type === 'func') return fileName(t.text.toUpperCase().replace(/^(_XL(FN|WS)\.)+/, ''));
    if (t.type === 'bool') return t.text.toUpperCase();
    if (t.type === 'ref' && t.area) return t.prefix + areaText(t.area);
    if (t.type === 'err') return t.value;
    return null;
  });
  if (depth > 0) text += ')'.repeat(depth);
  // (A formula that is all @, =@A1:A3, is saved as Excel 2019 wrote it, A1:A3: its implicit intersection is the old one.)
  if (/^_xlfn\.SINGLE\(/i.test(text)) {
    const t = tokenize(text);
    if (closing(t, 1) === t.length - 1) text = text.slice(t[2].start, t[t.length - 1].start);
  }
  return text;
}
