// Data > Advanced (Advanced Filter), as Excel does it (measured over COM: tools/excel-compare/excel-advfilter.ps1,
// test/advfilter-excel.json). The criteria range's first row names columns of the list (case aside); each row below is
// one way to pass (rows are Or), its cells all to hold (a row is And); a blank cell holds for anything. A cell's text:
// =x the value x itself, x alone a value beginning with x (text; a number: equal), or >, >=, <, <=, <> before a value;
// * and ? are wildcards. Unique records only drops a row the same as one before it. In place, the rows not passing are
// hidden; copied, the heading and the rows passing go to the cells given.
import { shownText } from './filterops.js';
import { setRows } from './sizes.js';
import { readSheet } from './sheet.js';
import { cellTest } from '../calc/criteria.js';

const blank = (cell) => !cell || cell.v == null || cell.v === '';
/**
 * The rows of `list` ({ r0, c0, r1, c1 }, its heading first) on sheet si that pass `criteria` (a range, on any sheet:
 * { si, r0, c0, r1, c1 }, or null for every row), without later duplicates when `unique`. Returns { rows: [r], error }.
 */
export function advancedRows(wb, si, list, criteria, unique = false) {
  const sheet = wb.sheet(si);
  const text = (sh, r, c) => shownText(wb, sh.cells.get(r)?.get(c));
  const heads = new Map();
  for (let c = list.c0; c <= list.c1; c++) heads.set(text(sheet, list.r0, c).trim().toLowerCase(), c);
  let ways = null;
  if (criteria) {
    const cs = wb.sheet(criteria.si ?? si);
    const cols = [];
    for (let c = criteria.c0; c <= criteria.c1; c++) {
      const name = text(cs, criteria.r0, c).trim().toLowerCase();
      if (!heads.has(name)) return { rows: [], error: "The criteria's headings must be headings of the list." };
      cols.push([c, heads.get(name)]);
    }
    ways = [];
    for (let r = criteria.r0 + 1; r <= criteria.r1; r++) {
      const tests = cols.filter(([c]) => !blank(cs.cells.get(r)?.get(c))).map(([c, lc]) => [lc, cellTest(text(cs, r, c))]);
      ways.push(tests);
    }
  }
  const rows = [];
  const seen = new Set();
  for (let r = list.r0 + 1; r <= list.r1; r++) {
    const value = (c) => { const x = sheet.cells.get(r)?.get(c); return blank(x) ? null : x.v; };
    const ok = !ways || ways.some((tests) => tests.every(([c, t]) => t(value(c), text(sheet, r, c))));
    if (!ok) continue;
    if (unique) {
      const key = [];
      for (let c = list.c0; c <= list.c1; c++) key.push(text(sheet, r, c).toLowerCase());
      const k = key.join('\u0000');
      if (seen.has(k)) continue;
      seen.add(k);
    }
    rows.push(r);
  }
  return { rows, error: null };
}

/** Data > Advanced on sheet i of `wb` (Workbook.advancedFilter): the rows passing, in place or copied to `copyTo`
 *  ({ r, c }); Excel's names for the list, the criteria and the copy. Returns null, or Excel's words. */
export function applyAdvancedFilter(wb, i, { list, criteria = null, copyTo = null, unique = false }) {
  const got = advancedRows(wb, i, list, criteria, unique);
  if (got.error) return got.error;
  const part = wb._part(i);
  if (!copyTo) {
    const keep = new Set(got.rows);
    for (let r = list.r0 + 1; r <= list.r1;) {
      const on = !keep.has(r);
      let end = r;
      while (end + 1 <= list.r1 && !keep.has(end + 1) === on) end++;
      setRows(part.doc, r, end, { hidden: on });
      r = end + 1;
    }
    wb._parsed.set(wb.sheets[i], { doc: part.doc, model: readSheet(part.doc, wb.sharedStrings) });
    wb.dirty.add(wb.sheets[i].path);
  } else {
    const sheet = wb.sheet(i);
    const contents = [];
    const styles = [];
    [list.r0, ...got.rows].forEach((r, k) => {
      for (let c = list.c0; c <= list.c1; c++) {
        const cell = sheet.cells.get(r)?.get(c);
        const at = { r: copyTo.r + k, c: copyTo.c + c - list.c0 };
        contents.push({ ...at, value: cell?.v ?? null });
        if (cell?.s) styles.push({ ...at, s: cell.s });
      }
    });
    for (const x of contents) wb.setCell(i, { r: x.r, c: x.c }, x.value);
    if (styles.length) wb.setStyles(i, styles);
    wb._filterName(i, { r0: copyTo.r, c0: copyTo.c, r1: copyTo.r, c1: copyTo.c + list.c1 - list.c0 }, '_xlnm.Extract', false);
  }
  wb._filterName(i, list);
  if (criteria) wb._filterName(i, criteria, '_xlnm.Criteria', false);
  wb.edited = true;
  wb.resetCalc();
  return null;
}

