// Excel's background error checking, the green triangle in a cell's top-left corner, for the rules it has on by
// default (File > Options > Formulas; measured over COM with a hidden Excel's Range.Errors on ~60 cells):
// 'evalError'          a formula whose result is an error, but #N/A from a formula calling NA() (=NA(), =NA()+1);
// 'twoDigitTextYear'   text that reads as a date with a two-digit year: '1/2/24, 'Jan 2, 24, '2-Jan-24;
// 'numberStoredAsText' text that reads as a number: '123, '12.5, '1,234, '1e3, ' 7, '50%, '(5), '-3, '0012 (not
//                      dates or times, 'TRUE, or a formula's text);
// 'formula'            inconsistent: the cells either side (left and right, or above and below) hold the same
//                      formula as each other, moved, and this one differs;
// 'formulaRange'       omits adjacent cells: a number (not text, TRUE or a formula) just before or after a range of
//                      one row or column the formula reads, in that row or column;
// 'unlockedFormula'    a formula in a cell whose format isn't locked.
// A sheet's ignoredErrors (Excel's Ignore Error) turn a rule off for their cells. The names are ignoredError's.
import { kid, kids, attr } from './xml.js';
import { textToNumber } from '../calc/values.js';
import { parseFormula } from '../calc/parse.js';
import { shiftFormula } from '../calc/tokens.js';
import { parseRange } from './refs.js';

/** What each error indicator says, as Excel's error button does. */
export const ERROR_TEXT = {
  evalError: "This cell's formula results in an error.",
  twoDigitTextYear: 'This cell contains a date string represented with only 2 digits for the year.',
  numberStoredAsText: 'The number in this cell is formatted as text or preceded by an apostrophe.',
  formula: 'The formula in this cell differs from the formulas in this area of the spreadsheet.',
  formulaRange: 'The formula in this cell refers to a range that has additional numbers adjacent to it.',
  unlockedFormula: 'This cell contains a formula and is not locked to protect it from being changed inadvertently.',
};

export const ERROR_RULES = ['evalError', 'twoDigitTextYear', 'numberStoredAsText', 'formula', 'formulaRange', 'unlockedFormula'];

const MONTH = '(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\\.?';
const TWO_DIGIT_DATES = [
  /^\d{1,2}[/-]\d{1,2}[/-]\d{2}$/,
  new RegExp(`^${MONTH}\\s+\\d{1,2},?\\s+\\d{2}$`, 'i'),
  new RegExp(`^\\d{1,2}[-\\s]${MONTH}[-\\s]\\d{2}$`, 'i'),
];
const isText = (cell) => cell && cell.f == null && (cell.t === 's' || cell.t === 'str' || cell.t === 'inlineStr') && typeof cell.v === 'string';
const isNumber = (cell) => cell && cell.f == null && cell.t === 'n' && cell.v != null;

/** The rules a text constant breaks: a two-digit-year date, or a number stored as text. */
function textRule(v) {
  const s = v.trim();
  if (TWO_DIGIT_DATES.some((re) => re.test(s))) return 'twoDigitTextYear';
  // (Dates and times read as numbers too; Excel doesn't call them numbers stored as text.)
  if (/[/:]|^\d{4}-\d/.test(s) || !/\d/.test(s)) return null;
  return textToNumber(v) != null ? 'numberStoredAsText' : null;
}

/** The one-row or one-column ranges a formula reads (other sheets' by their index), each { si, r0, c0, r1, c1 }. */
function rangesOf(f, si, wb) {
  let tree;
  try { tree = parseFormula(f); } catch { return []; }
  const out = [];
  const walk = (n) => {
    if (!n || typeof n !== 'object') return;
    if (n.k === 'ref' && n.book == null) {
      const a = n.area;
      const r0 = Math.min(a.r0, a.r1); const r1 = Math.max(a.r0, a.r1); const c0 = Math.min(a.c0, a.c1); const c1 = Math.max(a.c0, a.c1);
      const s = n.sheet == null ? si : wb.sheetIndexOf(n.sheet);
      if (s >= 0 && (r0 === r1) !== (c0 === c1) && r1 - r0 < 1048575 && c1 - c0 < 16383) out.push({ si: s, r0, c0, r1, c1 });
      return;
    }
    for (const v of Object.values(n)) if (Array.isArray(v)) v.forEach(walk); else if (v && typeof v === 'object' && v.k) walk(v);
  };
  walk(tree);
  return out;
}

/** Whether a formula leaves out a number just before or after a one-row or one-column range it reads. */
function omitsCells(wb, si, f) {
  for (const g of rangesOf(f, si, wb)) {
    const cells = wb.sheet(g.si).cells;
    const at = (r, c) => (r >= 0 && c >= 0 ? cells.get(r)?.get(c) : null);
    const ends = g.c0 === g.c1 ? [at(g.r0 - 1, g.c0), at(g.r1 + 1, g.c0)] : [at(g.r0, g.c0 - 1), at(g.r0, g.c1 + 1)];
    if (ends.some(isNumber)) return true;
  }
  return false;
}

/** Whether the cells either side of (r, c) hold one formula between them, moved, and this cell another. */
function inconsistent(cells, r, c, f) {
  const at = (rr, cc) => (rr >= 0 && cc >= 0 ? cells.get(rr)?.get(cc) : null);
  for (const [dr, dc] of [[1, 0], [0, 1]]) {
    const a = at(r - dr, c - dc);
    const b = at(r + dr, c + dc);
    if (!a?.f || !b?.f || a.array || b.array) continue;
    const moved = (text, n) => { try { return shiftFormula(text, dr * n, dc * n); } catch { return null; } };
    if (moved(a.f, 2) === b.f && moved(a.f, 1) !== f) return true;
  }
  return false;
}

/** The sheet's ignoredErrors: rule => [ranges]. */
function ignored(wb, si) {
  const doc = wb._part(si)?.doc;
  const out = new Map();
  if (!doc) return out;
  const root = doc.documentElement;
  for (const el of kids(kid(root, 'ignoredErrors', root.namespaceURI), 'ignoredError', root.namespaceURI)) {
    const ranges = (attr(el, 'sqref') || '').split(/\s+/).filter(Boolean).map(parseRange).filter(Boolean);
    for (const rule of ERROR_RULES.concat('emptyCellReference', 'calculatedColumn', 'listDataValidation')) {
      if (attr(el, rule) === '1') { if (!out.has(rule)) out.set(rule, []); out.get(rule).push(...ranges); }
    }
  }
  return out;
}

/** Sheet si's cells with an error indicator: Map "r,c" => [rules]. A sheet of more than `limit` cells isn't checked
 *  (Excel checks in the background). */
export function errorFlags(wb, si, limit = 200000) {
  const locked = (s) => wb.styles?.cellXfs?.[s]?.protection?.locked !== false;
  const out = new Map();
  const sheet = wb.sheet(si);
  if (!sheet) return out;
  let n = 0;
  for (const line of sheet.cells.values()) n += line.size;
  if (n > limit) return out;
  const skip = ignored(wb, si);
  const off = (rule, r, c) => skip.get(rule)?.some((g) => r >= g.r0 && r <= g.r1 && c >= g.c0 && c <= g.c1);
  for (const [r, line] of sheet.cells) {
    for (const [c, cell] of line) {
      const rules = [];
      if (cell.f != null) {
        if (cell.t === 'e' && !(cell.v === '#N/A' && /\bNA\(/i.test(cell.f))) rules.push('evalError');
        if (!cell.array && inconsistent(sheet.cells, r, c, cell.f)) rules.push('formula');
        if (omitsCells(wb, si, cell.f)) rules.push('formulaRange');
        if (!locked(cell.s || 0)) rules.push('unlockedFormula');
      } else if (isText(cell)) {
        const rule = textRule(cell.v);
        if (rule) rules.push(rule);
      }
      const kept = rules.filter((rule) => !off(rule, r, c));
      if (kept.length) out.set(`${r},${c}`, kept);
    }
  }
  return out;
}
