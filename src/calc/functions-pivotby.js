// Excel 365's PIVOTBY: values grouped by the rows of one or more row fields and across the rows of one or more column
// fields, each group's values given to a function. As Excel works it out, measured with a hidden Excel on 76 formulas
// (spilled answers read cell by cell): field values across in as many header rows as there are column fields, a
// subtotal row or column for each group of the levels the total depth reaches, "Grand Total" with them (else
// "Total"); the column fields' header "Item, Size" in one cell; sorted by fields, or by a value column's totals;
// PERCENTOF and two-argument LAMBDAs of the column's total, the row's, the grand total or a parent's (relative_to).
// The helpers it shares with GROUPBY are in functions-groupby.js.
import { Arr, E, matrix } from './values.js';
import { def, given } from './functions.js';
import { ARRAY, whole, keyOf, functionsOf, aggregate, headersOf, kept, tree, sortKeys, depthOf, sorter } from './functions-groupby.js';

/** The tree of groups as lines in order: { path, rows }, a subtotal line after (or, negative depth, before) its
 *  group's for the levels the depth reaches, and the grand total's ({ path: [], grand }) at the end or start. */
function lines(groups, k, depth, all) {
  const out = [];
  const walk = (list, path, d) => {
    for (const g of list) {
      const here = [...path, g.value];
      if (d === k - 1) { out.push({ path: here, rows: g.rows }); continue; }
      const sub = { path: here, rows: g.rows };
      const subtotal = Math.abs(depth) > d + 1;
      if (subtotal && depth < 0) out.push(sub);
      walk(g.kids, here, d + 1);
      if (subtotal && depth > 0) out.push(sub);
    }
  };
  walk(groups, [], 0);
  const grand = { path: [], rows: all, grand: true };
  if (depth < 0) out.unshift(grand);
  else if (depth > 0) out.push(grand);
  return out;
}

/** PIVOTBY(row_fields, col_fields, values, function, [field_headers], [row_total_depth], [row_sort_order],
 *  [col_total_depth], [col_sort_order], [filter_array], [relative_to]). */
function pivotBy(a, ctx) {
  let rf = matrix(a[0], ctx);
  let cf = matrix(a[1], ctx);
  let valueRows = matrix(a[2], ctx);
  if (rf.length !== valueRows.length || cf.length !== valueRows.length) throw E.value;
  const fns = functionsOf(a[3]);
  const kr = rf[0].length;
  const kc = cf[0].length;
  const m = valueRows[0].length;
  const nf = fns.list.length;
  const h = headersOf(a, 4, valueRows, ctx);
  const made = (n, name) => Array.from({ length: n }, (_, i) => `${name} ${i + 1}`);
  let heads = null;
  if (h === 1 || h === 3) { heads = { rows: rf[0], cols: cf[0], values: valueRows[0] }; rf = rf.slice(1); cf = cf.slice(1); valueRows = valueRows.slice(1); }
  if (h === 2) heads = { rows: made(kr, 'Row Field'), cols: made(kc, 'Column Field'), values: made(m, 'Value') };
  const rowSpec = sortKeys(a, 6, kr, m * nf, ctx);
  const colSpec = sortKeys(a, 8, kc, m * nf, ctx);
  const rowTotal = depthOf(a, 5, kr, rowSpec, ctx);
  const colTotal = depthOf(a, 7, kc, colSpec, ctx);
  const keep = kept(a, 9, rf.length, ctx);
  const relative = given(a[10]) ? whole(a[10], ctx) : 0;
  if (relative < 0 || relative > 4) throw E.value;
  const rows = rf.map((_, i) => i).filter((i) => keep[i]);
  const valuesOf = (list, c) => list.map((i) => valueRows[i][c]);
  const allOf = (c) => valuesOf(rows, c);
  // (Sorted by a value column: the answer of the group's rows, across all the columns or down all the rows.)
  const answer = (list, j) => aggregate(fns.list[j % nf], valuesOf(list, Math.floor(j / nf)), allOf(Math.floor(j / nf)), ctx);
  // (Rows sort by value at every level; columns only at the last and where there are subtotals, as Excel does.)
  const rowTree = tree(rf, rows, kr, sorter(rowSpec, answer, () => true));
  const colTree = tree(cf, rows, kc, sorter(colSpec, answer, (d) => d === kc - 1 || Math.abs(colTotal) > d + 1));
  const rowLines = lines(rowTree, kr, rowTotal, rows);
  const colLines = lines(colTree, kc, colTotal, rows);
  for (const line of colLines) line.set = new Set(line.rows);
  // (The rows of a parent group, for relative_to 3 and 4: the path without its last field.)
  const under = new Map();
  const rowsUnder = (fields, path) => {
    const key = `${fields === rf ? 'r' : 'c'}${path.map(keyOf).join('\u0000')}`;
    if (!under.has(key)) under.set(key, rows.filter((i) => path.every((v, j) => keyOf(fields[i][j]) === keyOf(v))));
    return under.get(key);
  };
  const both = (list, other) => { const set = new Set(other); return list.filter((i) => set.has(i)); };
  const relativeTo = (line) => {
    if (relative === 0) return line.col.rows;
    if (relative === 1) return line.row.rows;
    if (relative === 2) return rows;
    if (relative === 3) return both(line.row.rows, rowsUnder(cf, line.col.path.slice(0, -1)));
    return both(line.col.rows, rowsUnder(rf, line.row.path.slice(0, -1)));
  };
  const cols = [];
  for (const col of colLines) for (let c = 0; c < m; c++) for (const fn of fns.list) cols.push({ col, c, fn });
  // (Several value columns, one function: the total and subtotal columns are left empty, as Excel leaves them.)
  const totalsEmpty = m > 1 && !fns.many;
  const cell = (row, x) => {
    if (totalsEmpty && x.col.path.length < kc) return '';
    const group = row.rows.filter((i) => x.col.set.has(i));
    if (!group.length) return '';
    return aggregate(x.fn, valuesOf(group, x.c), valuesOf(relativeTo({ row, col: x.col }), x.c), ctx);
  };
  const blank = (n) => Array(n).fill('');
  const out = [];
  if (heads && h !== 1) out.push([...blank(kr), heads.cols.map(String).join(', '), ...blank(cols.length - 1)]);
  const grandName = (depth) => (Math.abs(depth) > 1 ? 'Grand Total' : 'Total');
  for (let j = 0; j < kc; j++) {
    out.push([...blank(kr), ...cols.map((x) => (x.col.grand ? (j === 0 ? grandName(colTotal) : '') : j < x.col.path.length ? x.col.path[j] : ''))]);
  }
  if (fns.many) out.push([...blank(kr), ...cols.map((x) => fns.names[fns.list.indexOf(x.fn)])]);
  if (heads && h !== 1) out.push([...heads.rows, ...cols.map((x) => heads.values[x.c])]);
  for (const row of rowLines) {
    const label = row.grand ? [grandName(rowTotal), ...blank(kr - 1)] : [...row.path, ...blank(kr - row.path.length)];
    out.push([...label, ...cols.map((x) => cell(row, x))]);
  }
  return new Arr(out);
}
def('PIVOTBY', (a, ctx) => pivotBy(a, ctx), ARRAY);
