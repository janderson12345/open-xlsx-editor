// Find and Replace as Excel does them (measured over COM, docs/FIND.md). A cell is searched in the text of its formula
// bar (Look in: Formulas; a formula with its "=", a percentage as 50%, a date as 1/1/2024) or in the text it shows
// (Values). Search goes from the cell after the active one, by rows or by columns, round to the start; * and ? are
// wildcards, ~ takes the next character as it is. A replacement is entered as if typed (a number becomes a number,
// "=..." a formula); in a formula it is plain text, so A1 to B9 also makes A10 B90, as in Excel.
import { formatValue, isDateFormat } from './format.js';
import { formatGeneral, cellDisplay } from './numfmt.js';
import { formatCodeOf } from './styles.js';
import { displayFormula } from '../calc/tokens.js';
import { parseInput } from './input.js';

/** A pattern as a regular expression: Excel's wildcards, case ignored unless `matchCase`, the whole text when
 *  `whole`. Null for an empty pattern. */
export function pattern(what, { matchCase = false, whole = false } = {}, global = false) {
  if (!what) return null;
  let src = '';
  for (let i = 0; i < what.length; i++) {
    const ch = what[i];
    if (ch === '~' && i + 1 < what.length && '*?~'.includes(what[i + 1])) { src += `\\${what[++i]}`; continue; }
    if (ch === '*') src += '[\\s\\S]*';
    else if (ch === '?') src += '[\\s\\S]';
    else src += ch.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');
  }
  return new RegExp(whole ? `^(?:${src})$` : src, `${matchCase ? '' : 'i'}${global ? 'g' : ''}`);
}

const PERCENT = /%(?=(?:[^"]*"[^"]*")*[^"]*$)/;

/** A cell's text as Excel's formula bar has it, which Look in Formulas searches. */
export function formulaText(cell, code = 'General', date1904 = false) {
  if (!cell) return '';
  if (cell.f) return `=${displayFormula(cell.f)}`;
  if (cell.v == null) return '';
  if (cell.t === 'b') return cell.v ? 'TRUE' : 'FALSE';
  if (cell.t !== 'n') return String(cell.v);
  if (code && code !== 'General' && isDateFormat(code)) {
    const frac = cell.v % 1;
    const date = formatValue(Math.floor(cell.v), 'm/d/yyyy', { date1904 }).text;
    return frac ? `${date} ${formatValue(cell.v, 'h:mm:ss AM/PM', { date1904 }).text}` : date;
  }
  if (code && PERCENT.test(code)) return `${formatGeneral(Math.round(cell.v * 100 * 1e10) / 1e10)}%`;
  return String(cell.v);
}

/** The text of cell (r, c) of sheet si that a search looks at. */
export function searchText(wb, si, cell, lookIn = 'formulas') {
  const code = formatCodeOf(wb.styles, cell?.s || 0);
  return lookIn === 'values' ? cellDisplay(cell, Infinity, code, wb.date1904).text : formulaText(cell, code, wb.date1904);
}

/** The sheets a search goes through: the one shown, or every worksheet that can be shown from it on (Within:
 *  Workbook). */
const sheetsFrom = (wb, si, within) => (within === 'workbook'
  ? [...wb.sheets.keys()].filter((k) => wb.canShow(k)).sort((a, b) => ((a - si + wb.sheets.length) % wb.sheets.length) - ((b - si + wb.sheets.length) % wb.sheets.length))
  : [si]);

/** The cells of sheet si that match, in search order (rows, or columns with `byColumns`). */
function matchesIn(wb, si, re, opts) {
  const out = [];
  for (const [r, line] of wb.sheet(si).cells) {
    for (const [c, cell] of line) if (re.test(searchText(wb, si, cell, opts.lookIn))) out.push({ si, r, c });
  }
  return out.sort(opts.byColumns ? (a, b) => a.c - b.c || a.r - b.r : (a, b) => a.r - b.r || a.c - b.c);
}

/**
 * Every matching cell, from the sheet `si` on (opts: { matchCase, whole, lookIn: 'formulas' | 'values', byColumns,
 * within: 'sheet' | 'workbook' }): [{ si, r, c }].
 */
export function findAll(wb, si, what, opts = {}) {
  const re = pattern(what, opts);
  if (!re) return [];
  return sheetsFrom(wb, si, opts.within).flatMap((k) => matchesIn(wb, k, re, opts));
}

/** The next match after cell `from` ({ si, r, c }), or before it (dir -1), round to the start; null when nothing
 *  matches. (A search from a cell finds that cell last, as Excel's does.) */
export function findNext(wb, from, what, opts = {}, dir = 1) {
  const all = findAll(wb, from.si, what, opts);
  if (!all.length) return null;
  const key = (x) => {
    const s = (x.si - from.si + wb.sheets.length) % wb.sheets.length;
    return opts.byColumns ? [s, x.c, x.r] : [s, x.r, x.c];
  };
  const cmp = (a, b) => { const ka = key(a); const kb = key(b); return ka[0] - kb[0] || ka[1] - kb[1] || ka[2] - kb[2]; };
  if (dir > 0) return all.find((x) => cmp(x, from) > 0) || all[0];
  return [...all].reverse().find((x) => cmp(x, from) < 0) || all.at(-1);
}

/** The text after a replacement: the whole text when `whole`, else every match in it. Null when nothing matches. */
export function replaced(text, what, withText, opts = {}) {
  const re = pattern(what, opts, true);
  if (!re || !re.test(text)) return null;
  re.lastIndex = 0;
  return opts.whole ? withText : text.replace(re, () => withText);
}

/**
 * The changes Replace makes to cells [{ si, r, c }] (Look in is always Formulas, as in Excel's Replace): grouped by
 * sheet for the editor's undo history, [{ si, contents: [{ r, c, value } or { r, c, formula }], styles: [], merges }],
 * with `count` and `refused` (a replacement that isn't a formula Excel would take).
 */
export function replaceChanges(wb, cells, what, withText, opts = {}) {
  const bySheet = new Map();
  let count = 0;
  let refused = 0;
  for (const { si, r, c } of cells) {
    const cell = wb.sheet(si).cells.get(r)?.get(c);
    const next = replaced(searchText(wb, si, cell, 'formulas'), what, withText, opts);
    if (next == null) continue;
    const p = parseInput(next);
    // (A date's text changed would be text, not a date: dates typed aren't read as dates yet, so it is left alone.)
    const date = cell.t === 'n' && !cell.f && isDateFormat(formatCodeOf(wb.styles, cell.s || 0));
    if (p.error || (date && typeof p.value !== 'number')) { refused++; continue; }
    if (!bySheet.has(si)) bySheet.set(si, { si, contents: [], styles: [], merges: wb.sheet(si).merges });
    bySheet.get(si).contents.push(p.formula != null ? { r, c, formula: p.formula } : { r, c, value: p.value });
    count++;
  }
  return Object.assign([...bySheet.values()], { count, refused });
}
