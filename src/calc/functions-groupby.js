// Excel 365's GROUPBY and PIVOTBY (values grouped by the rows of one or more fields, and across the rows of another,
// each group's values given to a function: SUM, AVERAGE, PERCENTOF, a LAMBDA, or several side by side), PERCENTOF,
// and RANDARRAY. As Excel works them out, measured with a hidden Excel (spilled answers read cell by cell) and the
// formula probe: headers found when the first value is text over a number, shown with 3, made up with 2 ("Row Field
// 1", "Value 1"); totals at the end, at the start when negative, "Grand Total" with subtotals; sorted by the fields,
// or by a value column; an empty cell of the answer is "" (not a blank). Registered into functions.js's table.
import { Area, Arr, E, XlError, MISSING, Lambda, Lambdas, Builtin, Unsupported, matrix, scalar, toBool, compare } from './values.js';
import { def, given, num, numbers, cellsIn } from './functions.js';

const ARRAY = { array: true };
const whole = (v, ctx) => Math.trunc(num(v, ctx));
const column = (rows, c) => rows.map((row) => row[c]);
const vec = (values) => new Arr(values.map((v) => [v]));
/** Field values in Excel's order for them: numbers, text, TRUE and FALSE, then a blank last (its own group, not FALSE). */
const RANK = (v) => (v == null ? 4 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3);
const order = (x, y) => RANK(x) - RANK(y) || (x == null || y == null ? 0 : compare(x, y));
/** A field value as a group's key: text case aside, a blank its own. */
const keyOf = (v) => (v == null ? 'blank' : typeof v === 'string' ? `s${v.toUpperCase()}` : `${typeof v}${v}`);

/** The function argument as a list: one LAMBDA, or HSTACK's; with the name each column's header row shows. */
function functionsOf(v) {
  const list = v instanceof Lambdas ? v.list : [v];
  if (!list.every((f) => f instanceof Lambda)) throw E.value;
  return { list, names: list.map((f) => (f instanceof Builtin ? f.name : 'LAMBDA')), many: list.length > 1 };
}
/** A group's values given to a function (PERCENTOF and a two-argument LAMBDA also get all the values). */
function aggregate(fn, group, all, ctx) {
  const two = fn instanceof Builtin ? fn.name === 'PERCENTOF' : fn.params?.length === 2;
  let v = ctx.callLambda(fn, two ? [vec(group), vec(all)] : [vec(group)]);
  if (v instanceof Arr || (v && typeof v === 'object' && 'si' in v)) {
    const m = matrix(v, ctx);
    v = m.length === 1 && m[0].length === 1 ? m[0][0] : E.calc;
  }
  return v == null || v === MISSING ? 0 : v;
}

/** field_headers: 0 none, 1 there but not shown, 2 none but made up, 3 there and shown; left out: there when the
 *  first value is text and the one below a number. */
function headersOf(a, k, valueRows, ctx) {
  if (!given(a[k])) return typeof valueRows[0]?.[0] === 'string' && typeof valueRows[1]?.[0] === 'number' ? 1 : 0;
  const h = whole(a[k], ctx);
  if (h < 0 || h > 3) throw E.value;
  return h;
}

/** The rows kept: filter_array's TRUE ones (as long as the rows), all when it is left out. */
function kept(a, k, n, ctx) {
  if (!given(a[k])) return Array(n).fill(true);
  const f = matrix(a[k], ctx).flat();
  if (f.length !== n) throw E.value;
  return f.map((x) => { if (x instanceof XlError) throw x; try { return toBool(x); } catch { return false; } });
}

/** Groups of rows by their field values, in a tree: { key, value, rows, kids } at each level, sorted. */
function tree(fields, rows, depth, sortBy) {
  const level = (list, d) => {
    if (d === depth) return [];
    const byKey = new Map();
    for (const i of list) {
      const v = fields[i][d];
      const key = keyOf(v);
      if (!byKey.has(key)) byKey.set(key, { value: v, rows: [] });
      byKey.get(key).rows.push(i);
    }
    const groups = [...byKey.values()].map((g) => ({ ...g, kids: level(g.rows, d + 1) }));
    groups.sort((x, y) => sortBy(x, y, d));
    return groups;
  };
  return level(rows, 0);
}

/** GROUPBY(row_fields, values, function, [field_headers], [total_depth], [sort_order], [filter_array],
 *  [field_relationship]). */
function groupBy(a, ctx) {
  const fieldRows = matrix(a[0], ctx);
  let valueRows = matrix(a[1], ctx);
  if (fieldRows.length !== valueRows.length) throw E.value;
  const fns = functionsOf(a[2]);
  const k = fieldRows[0].length;
  const m = valueRows[0].length;
  const h = headersOf(a, 3, valueRows, ctx);
  let fields = fieldRows;
  let heads = null;
  if (h === 1 || h === 3) { heads = [...fields[0], ...valueRows[0]]; fields = fields.slice(1); valueRows = valueRows.slice(1); }
  if (h === 2) heads = [...Array.from({ length: k }, (_, i) => `Row Field ${i + 1}`), ...Array.from({ length: m }, (_, i) => `Value ${i + 1}`)];
  const total = given(a[4]) ? whole(a[4], ctx) : 1;
  if (Math.abs(total) > Math.max(1, k)) throw E.value;
  // (One column to sort by; an array of them, {-1,3}, is #VALUE! here as it was in Excel.)
  if (given(a[5]) && matrix(a[5], ctx).flat().length > 1) throw E.value;
  const sort = given(a[5]) ? scalar(a[5], ctx) : 1;
  if (typeof sort !== 'number' || !Number.isInteger(sort) || sort === 0 || Math.abs(sort) > k + m * fns.list.length) throw E.value;
  const keep = kept(a, 6, fields.length, ctx);
  const relation = given(a[7]) ? whole(a[7], ctx) : 0;
  if (relation !== 0 && relation !== 1) throw E.value;
  if (relation === 1 && Math.abs(total) > 1) throw E.value;
  const rows = fields.map((_, i) => i).filter((i) => keep[i]);
  const all = (c) => rows.map((i) => valueRows[i][c]);
  // (Each group's answers: for each value column, for each function.)
  const answers = (list) => {
    const out = [];
    for (let c = 0; c < m; c++) for (const fn of fns.list) out.push(aggregate(fn, list.map((i) => valueRows[i][c]), all(c), ctx));
    return out;
  };
  const sortBy = (x, y, d) => {
    if (Math.abs(sort) <= k) {
      const sign = Math.abs(sort) === d + 1 && sort < 0 ? -1 : 1;
      return sign * order(x.value, y.value);
    }
    const col = Math.abs(sort) - k - 1;
    const r = compare(answers(x.rows)[col], answers(y.rows)[col]);
    return (sort < 0 ? -r : r) || order(x.value, y.value);
  };
  const groups = tree(fields, rows, k, sortBy);
  const out = [];
  if (fns.many) out.push([...Array(k).fill(''), ...Array.from({ length: m }, () => fns.names).flat()]);
  if (heads && h !== 1) out.push([...heads.slice(0, k), ...heads.slice(k).flatMap((name) => fns.list.map(() => name))]);
  const blank = (n) => Array(n).fill('');
  const body = [];
  const walk = (list, path, d) => {
    for (const g of list) {
      const here = [...path, g.value];
      if (d === k - 1) { body.push([...here, ...answers(g.rows)]); continue; }
      const sub = [...here, ...blank(k - here.length), ...answers(g.rows)];
      // (Subtotals for the levels total_depth reaches: before the group's rows when it is negative.)
      const subtotal = Math.abs(total) > d + 1;
      if (subtotal && total < 0) body.push(sub);
      walk(g.kids, here, d + 1);
      if (subtotal && total > 0) body.push(sub);
    }
  };
  walk(groups, [], 0);
  const grand = [Math.abs(total) > 1 ? 'Grand Total' : 'Total', ...blank(k - 1), ...answers(rows)];
  if (total < 0) out.push(grand);
  out.push(...body);
  if (total > 0) out.push(grand);
  return new Arr(out);
}
def('GROUPBY', (a, ctx) => groupBy(a, ctx), ARRAY);

/** PIVOTBY(row_fields, col_fields, values, function, [field_headers], [row_total_depth], [row_sort_order],
 *  [col_total_depth], [col_sort_order], [filter_array], [relative_to]): one row field and one column field here. */
function pivotBy(a, ctx) {
  const rowFields = matrix(a[0], ctx);
  const colFields = matrix(a[1], ctx);
  let valueRows = matrix(a[2], ctx);
  if (rowFields.length !== valueRows.length || colFields.length !== valueRows.length) throw E.value;
  if (rowFields[0].length !== 1 || colFields[0].length !== 1) throw new Unsupported('PIVOTBY of more than one field');
  const fns = functionsOf(a[3]);
  const m = valueRows[0].length;
  const h = headersOf(a, 4, valueRows, ctx);
  let rf = column(rowFields, 0);
  let cf = column(colFields, 0);
  let heads = null;
  if (h === 1 || h === 3) { heads = { row: rf[0], col: cf[0], values: valueRows[0] }; rf = rf.slice(1); cf = cf.slice(1); valueRows = valueRows.slice(1); }
  if (h === 2) heads = { row: 'Row Field 1', col: 'Col Field 1', values: valueRows[0].map((_, i) => `Value ${i + 1}`) };
  const depth = (k) => { const d = given(a[k]) ? whole(a[k], ctx) : 1; if (Math.abs(d) > 1) throw E.value; return d; };
  const rowTotal = depth(5);
  const colTotal = depth(7);
  const sortArg = (k) => { const s = given(a[k]) ? whole(a[k], ctx) : 1; if (Math.abs(s) !== 1) throw E.value; return s; };
  const rowSort = sortArg(6);
  const colSort = sortArg(8);
  const keep = kept(a, 9, rf.length, ctx);
  const rows = rf.map((_, i) => i).filter((i) => keep[i]);
  const distinct = (vals, sign) => {
    const seen = new Map();
    for (const i of rows) { const v = vals[i]; if (!seen.has(keyOf(v))) seen.set(keyOf(v), v); }
    return [...seen.values()].sort((x, y) => sign * order(x, y));
  };
  const same = (x, y) => keyOf(x) === keyOf(y);
  const rowKeys = distinct(rf, rowSort);
  const colKeys = distinct(cf, colSort);
  // (Several value columns: the row totals are left empty, as Excel leaves them.)
  const rowTotalsEmpty = m > 1 && !fns.many;
  const cell = (rk, ck, c, fn) => {
    const inRow = rows.filter((i) => (rk === undefined || same(rf[i], rk)));
    const inCol = rows.filter((i) => (ck === undefined || same(cf[i], ck)));
    const group = inRow.filter((i) => inCol.includes(i));
    if (!group.length) return '';
    // (PERCENTOF: of the column's total, as relative_to's default has it.)
    return aggregate(fn, group.map((i) => valueRows[i][c]), inCol.map((i) => valueRows[i][c]), ctx);
  };
  const colsOut = [];
  for (const ck of colKeys) for (let c = 0; c < m; c++) for (const fn of fns.list) colsOut.push({ ck, c, fn, total: false });
  const totals = [];
  if (colTotal !== 0) for (let c = 0; c < m; c++) for (const fn of fns.list) totals.push({ ck: undefined, c, fn, total: true });
  const cols = colTotal < 0 ? [...totals, ...colsOut] : [...colsOut, ...totals];
  const out = [];
  if (heads && h !== 1) out.push(['', heads.col, ...Array(cols.length - 1).fill('')]);
  out.push(['', ...cols.map((x) => (x.total ? 'Total' : x.ck))]);
  if (fns.many) out.push(['', ...cols.map((x) => fns.names[fns.list.indexOf(x.fn)])]);
  if (heads && h !== 1) out.push([heads.row, ...cols.map((x) => heads.values[x.c])]);
  const line = (rk, label) => [label, ...cols.map((x) => (x.total && rowTotalsEmpty ? '' : cell(rk, x.ck, x.c, x.fn)))];
  const body = rowKeys.map((rk) => line(rk, rk));
  const grand = line(undefined, 'Total');
  if (rowTotal < 0) out.push(grand);
  out.push(...body);
  if (rowTotal > 0) out.push(grand);
  return new Arr(out);
}
def('PIVOTBY', (a, ctx) => pivotBy(a, ctx), ARRAY);

/** PERCENTOF(data_subset, data_all): the subset's sum over the whole's. */
def('PERCENTOF', (a, ctx) => {
  const part = numbers([a[0]], ctx).reduce((s, x) => s + x, 0);
  const whole2 = numbers([a[1]], ctx).reduce((s, x) => s + x, 0);
  if (whole2 === 0) throw E.div0;
  return part / whole2;
}, ARRAY);

/** RANDARRAY([rows], [columns], [min], [max], [whole_number]): 0 rows or columns #CALC!, fewer #VALUE!, min above max
 *  #VALUE!, whole numbers between bounds that aren't whole #VALUE!. */
def('RANDARRAY', (a, ctx) => {
  const rows = given(a[0]) ? whole(a[0], ctx) : 1;
  const cols = given(a[1]) ? whole(a[1], ctx) : 1;
  const min = given(a[2]) ? num(a[2], ctx) : 0;
  const max = given(a[3]) ? num(a[3], ctx) : 1;
  const ints = given(a[4]) ? toBool(scalar(a[4], ctx)) : false;
  if (rows < 0 || cols < 0 || min > max) throw E.value;
  if (rows === 0 || cols === 0) throw E.calc;
  if (ints && (!Number.isInteger(min) || !Number.isInteger(max))) throw E.value;
  const one = () => (ints ? min + Math.floor(Math.random() * (max - min + 1)) : min + Math.random() * (max - min));
  return new Arr(Array.from({ length: rows }, () => Array.from({ length: cols }, one)));
}, { volatile: true });

/** TRIMRANGE(range, [trim_rows], [trim_cols]): the range without its empty rows and columns at the edges (0 none, 1
 *  leading, 2 trailing, 3 both, the default), still a reference (=SUM(TRIMRANGE(A:A)) sums what A holds). A cell
 *  holding "" isn't empty. Nothing left: #REF!. An array comes back as it is. The trim references are this too:
 *  A1:.E8 is _xlfn._TRO_TRAILING(A1:E8) in the file, A1.:E8 _TRO_LEADING, A1.:.E8 _TRO_ALL. */
function trimRange(v, rowsHow, colsHow, ctx) {
  if (v instanceof XlError) return v;
  if (!(v instanceof Area)) return v;
  let r0 = Infinity; let r1 = -1; let c0 = Infinity; let c1 = -1;
  for (const [r, c, x] of cellsIn(v, ctx)) {
    if (x == null) continue;
    r0 = Math.min(r0, r); r1 = Math.max(r1, r); c0 = Math.min(c0, c); c1 = Math.max(c1, c);
  }
  if (r1 < 0) throw E.ref;
  const lead = (how) => how === 1 || how === 3;
  const trail = (how) => how === 2 || how === 3;
  return new Area(v.si, lead(rowsHow) ? r0 : v.r0, lead(colsHow) ? c0 : v.c0, trail(rowsHow) ? r1 : v.r1, trail(colsHow) ? c1 : v.c1);
}
const how = (a, k, ctx) => { if (!given(a[k])) return 3; const n = whole(a[k], ctx); if (n < 0 || n > 3) throw E.value; return n; };
def('TRIMRANGE', (a, ctx) => trimRange(a[0], how(a, 1, ctx), how(a, 2, ctx), ctx), ARRAY);
def('_TRO_TRAILING', (a, ctx) => trimRange(a[0], 2, 2, ctx), ARRAY);
def('_TRO_LEADING', (a, ctx) => trimRange(a[0], 1, 1, ctx), ARRAY);
def('_TRO_ALL', (a, ctx) => trimRange(a[0], 3, 3, ctx), ARRAY);
