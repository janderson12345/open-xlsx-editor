// Pivot tables: read from their parts, and refreshed from their source as Excel's Refresh does (docs/PIVOT.md,
// measured with tools/excel-compare/excel-pivot.ps1). A refresh reads the source range again, rebuilds the cache (each
// field's shared items in the order the values first come in the data, every record), keeps each field's items in the
// order they were shown with new ones after them, and lays the pivot out again in the compact form Excel 365 uses:
// "Row Labels", nested row items in one column with each group's subtotal on its own row, the column field's items
// across, the grand totals. The cells, the location and the row and column items are rewritten; everything else in the
// parts stays as it was.
//
// Refused (PIVOT_MESSAGES.unsupported, the pivot left as it is): report filters, hidden items, calculated fields or
// items, grouped fields, a source heading renamed, the tabular and outline forms, grand totals turned off, more than one column field, several
// values with a column field, values on rows.
import { kids, kid, attr } from './xml.js';
import { parseRange, rangeName } from './refs.js';
import { readTables } from './tables.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const PIVOT_MESSAGES = {
  unsupported: "This PivotTable can't be refreshed here yet",
  fieldName: 'The PivotTable field name is not valid. To create a PivotTable report, you must use data that is organized as a list with labeled columns. If you are changing the name of a PivotTable field, you must type a new name for the field.',
  source: 'Reference isn’t valid.',
};

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const flag = (el, name, dflt = true) => { const v = attr(el, name); return v == null ? dflt : v === '1' || v === 'true'; };

/** A cache item's value: a string, number, boolean, error text or null (missing). */
function itemValue(el) {
  const v = attr(el, 'v');
  switch (el.localName) {
    case 'n': return Number(v);
    case 'b': return v === '1' || v === 'true';
    case 'm': return null;
    case 'e': return { error: v };
    default: return v ?? '';
  }
}
/** The key two values are the same item by: Excel's pivot items match text case aside. */
const keyOf = (v) => (v == null ? 'm' : typeof v === 'string' ? `s${v.toLowerCase()}` : typeof v === 'number' ? `n${v}` : typeof v === 'boolean' ? `b${v}` : `e${v.error}`);
/** The XML of a value in the cache: <s>, <n>, <b>, <m/>, <e>. */
const itemXml = (v) => (v == null ? '<m/>' : typeof v === 'string' ? `<s v="${esc(v)}"/>` : typeof v === 'number' ? `<n v="${v}"/>` : typeof v === 'boolean' ? `<b v="${v ? 1 : 0}"/>` : `<e v="${esc(v.error)}"/>`);

/** Every pivot table of the workbook: [{ si, path, name, cachePath, recordsPath }]. */
export function readPivots(wb) {
  const out = [];
  wb.sheets.forEach((info, si) => {
    if (!info.path || info.kind !== 'worksheet') return;
    for (const rel of wb.pkg.readRels(info.path).filter((r) => r.type === `${REL}/pivotTable`)) {
      const cacheRel = wb.pkg.readRels(rel.target).find((r) => r.type === `${REL}/pivotCacheDefinition`);
      if (!cacheRel) continue;
      const recordsRel = wb.pkg.readRels(cacheRel.target).find((r) => r.type === `${REL}/pivotCacheRecords`);
      const doc = wb.pkg.xml(rel.target);
      out.push({ si, path: rel.target, name: attr(doc.documentElement, 'name'), cachePath: cacheRel.target, recordsPath: recordsRel?.target ?? null });
    }
  });
  return out;
}

/** The pivot table whose location takes in cell (r, c) of sheet si, or null. */
export function pivotAt(wb, si, r, c) {
  for (const p of readPivots(wb)) {
    if (p.si !== si) continue;
    const a = parseRange(attr(kid(wb.pkg.xml(p.path).documentElement, 'location'), 'ref') || '');
    if (a && r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1) return p;
  }
  return null;
}

/** The table definition, as read: what a refresh needs, or { unsupported: reason }. */
function readDefinition(doc) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const loc = kid(root, 'location', ns);
  const fields = kids(kid(root, 'pivotFields', ns), 'pivotField', ns).map((f) => {
    const items = kid(f, 'items', ns);
    return {
      axis: attr(f, 'axis'),
      sortType: attr(f, 'sortType') || 'manual',
      subtotal: flag(f, 'defaultSubtotal'),
      compact: flag(f, 'compact'),
      outline: flag(f, 'outline'),
      items: items ? kids(items, 'item', ns).map((i) => ({ x: attr(i, 'x') == null ? null : Number(attr(i, 'x')), t: attr(i, 't'), h: flag(i, 'h', false), f: flag(i, 'f', false) })) : null,
    };
  });
  const list = (name) => kids(kid(root, name, ns), 'field', ns).map((f) => Number(attr(f, 'x')));
  const def = {
    location: parseRange(attr(loc, 'ref') || ''),
    fields,
    rowFields: list('rowFields'),
    colFields: list('colFields'),
    dataFields: kids(kid(root, 'dataFields', ns), 'dataField', ns).map((d) => ({ name: attr(d, 'name'), fld: Number(attr(d, 'fld')), subtotal: attr(d, 'subtotal') || 'sum', showAs: attr(d, 'showDataAs') })),
    captions: { row: attr(root, 'rowHeaderCaption') || 'Row Labels', col: attr(root, 'colHeaderCaption') || 'Column Labels', grand: attr(root, 'grandTotalCaption') || 'Grand Total', missing: attr(root, 'missingCaption') },
  };
  const why = (() => {
    if (!def.location) return 'its location';
    if (kid(root, 'pageFields', ns)) return 'report filters';
    if (kid(root, 'calculatedItems', ns)) return 'calculated items';
    if (!flag(root, 'compact') || !flag(root, 'outline') || attr(root, 'gridDropZones') === '1') return 'its layout';
    if (!flag(root, 'rowGrandTotals') || !flag(root, 'colGrandTotals')) return 'grand totals turned off';
    if (fields.some((f) => !f.compact || !f.outline)) return 'its layout';
    if (fields.some((f) => f.items && f.items.some((i) => i.h || i.f || (i.t && i.t !== 'default')))) return 'hidden or calculated items';
    if (fields.some((f) => !f.subtotal && f.axis)) return 'subtotals turned off';
    if (!def.dataFields.length || def.dataFields.some((d) => d.showAs)) return 'its values';
    const cols = def.colFields.filter((x) => x !== -2);
    if (def.rowFields.includes(-2)) return 'values on rows';
    if (cols.length > 1) return 'more than one column field';
    if (cols.length && def.dataFields.length > 1) return 'several values with a column field';
    if (!def.rowFields.length && !cols.length) return 'no row or column field';
    return null;
  })();
  return why ? { unsupported: why } : def;
}

/** The source range of a cache: { si, r0, c0, r1, c1 } (a table by name: its header and data rows), or null. */
function sourceOf(wb, cacheRoot) {
  const ns = cacheRoot.namespaceURI;
  const src = kid(kid(cacheRoot, 'cacheSource', ns), 'worksheetSource', ns);
  if (!src || attr(kid(cacheRoot, 'cacheSource', ns), 'type') !== 'worksheet') return null;
  const name = attr(src, 'name');
  if (name) {
    const t = readTables(wb).find((x) => x.name.toLowerCase() === name.toLowerCase());
    if (t) return { si: t.si, r0: t.r0, c0: t.c0, r1: t.r1 - t.totals, c1: t.c1 };
    const d = wb.definedName(name, null);
    const m = d && /^(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/.exec(d.text);
    if (!m) return null;
    const si = wb.sheetIndexOf((m[1] ?? m[2]).replace(/''/g, "'"));
    const a = parseRange(`${m[3]}${m[4]}:${m[5]}${m[6]}`);
    return si < 0 || !a ? null : { si, ...a };
  }
  const si = wb.sheetIndexOf(attr(src, 'sheet') || '');
  const a = parseRange(attr(src, 'ref') || '');
  return si < 0 || !a ? null : { si, ...a };
}

/** A cell's value as the cache holds it: number, string, boolean, error, or null for an empty cell. */
function cellValue(cell) {
  if (!cell || cell.v == null || cell.v === '') return null;
  if (cell.t === 'n') return cell.v;
  if (cell.t === 'b') return !!cell.v;
  if (cell.t === 'e') return { error: String(cell.v) };
  return String(cell.v);
}

/** The value of a data field over records `recs`, as its function works it out (null: nothing to show). */
function aggregate(recs, fld, fn) {
  const all = recs.map((r) => r[fld]).filter((v) => v != null);
  const nums = all.filter((v) => typeof v === 'number');
  if (!recs.length) return null;
  switch (fn) {
    case 'count': return all.length;
    case 'countNums': return nums.length;
    case 'sum': return nums.reduce((a, b) => a + b, 0);
    case 'product': return nums.length ? nums.reduce((a, b) => a * b, 1) : 0;
    case 'max': return nums.length ? Math.max(...nums) : 0;
    case 'min': return nums.length ? Math.min(...nums) : 0;
    case 'average': return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : { error: '#DIV/0!' };
    case 'stdDev': case 'stdDevp': case 'var': case 'varp': {
      const n = nums.length;
      const p = fn.endsWith('p');
      if (n < (p ? 1 : 2)) return { error: '#DIV/0!' };
      const mean = nums.reduce((a, b) => a + b, 0) / n;
      const v = nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (p ? n : n - 1);
      return fn.startsWith('std') ? Math.sqrt(v) : v;
    }
    default: return null;
  }
}

/** Display order of items: the values compared as Excel sorts them (numbers before text, text case aside). */
const compareItems = (a, b) => {
  const rank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : v == null ? 4 : 3);
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a === 'number') return a - b;
  if (typeof a === 'string') return a.localeCompare(b, 'en', { sensitivity: 'base' });
  return 0;
};

/** The sharedItems element's attributes for a field's values (as Excel writes them). */
function sharedAttrs(values, shared) {
  const present = values.filter((v) => v != null);
  const nums = present.filter((v) => typeof v === 'number');
  const strings = present.filter((v) => typeof v === 'string');
  const blank = present.length < values.length;
  const a = [];
  if (nums.length && nums.length === present.length) {
    a.push('containsSemiMixedTypes="0"', 'containsString="0"');
    if (blank) a.push('containsBlank="1"');
    a.push('containsNumber="1"');
    if (nums.every(Number.isInteger)) a.push('containsInteger="1"');
    a.push(`minValue="${Math.min(...nums)}"`, `maxValue="${Math.max(...nums)}"`);
  } else {
    if (blank) a.push('containsBlank="1"');
    if (nums.length && strings.length) {
      a.push('containsMixedTypes="1"', 'containsNumber="1"');
      if (nums.every(Number.isInteger)) a.push('containsInteger="1"');
      a.push(`minValue="${Math.min(...nums)}"`, `maxValue="${Math.max(...nums)}"`);
    }
  }
  if (shared) a.push(`count="${shared.length}"`);
  return a.length ? ` ${a.join(' ')}` : '';
}

/**
 * Refresh the pivot tables of the cache of pivot `p` (readPivots) from its source, as Excel's Refresh does. Returns
 * { changes: [{ si, contents, area, labels }] } for the cells (the caller makes them, for its history; `area` the
 * table's new location, `labels` its row labels' range or null), having rewritten the cache, the records and each
 * table's location and items; or { error } with Excel's words, nothing changed. `fresh`: a PivotTable just made
 * (pivotnew.js), whose items Excel lists sorted, as Sort A to Z would, though the field isn't marked sorted.
 */
export function refreshPivot(wb, p, { fresh = false } = {}) {
  const pkg = wb.pkg;
  const cacheText = pkg.text(p.cachePath);
  const cacheRoot = pkg.xml(p.cachePath).documentElement;
  const ns = cacheRoot.namespaceURI;
  if (kid(cacheRoot, 'calculatedItems', ns)) return { error: PIVOT_MESSAGES.unsupported, why: 'calculated items' };
  const cacheFields = kids(kid(cacheRoot, 'cacheFields', ns), 'cacheField', ns);
  if (cacheFields.some((f) => attr(f, 'formula') != null || kid(f, 'fieldGroup', ns) || attr(f, 'databaseField') === '0')) return { error: PIVOT_MESSAGES.unsupported, why: 'calculated or grouped fields' };
  const tables = readPivots(wb).filter((x) => x.cachePath === p.cachePath).map((x) => ({ ...x, def: readDefinition(pkg.xml(x.path)) }));
  const bad = tables.find((x) => x.def.unsupported);
  if (bad) return { error: PIVOT_MESSAGES.unsupported, why: bad.def.unsupported };
  const src = sourceOf(wb, cacheRoot);
  if (!src) return { error: PIVOT_MESSAGES.source };
  const sheet = wb.sheet(src.si);
  const width = src.c1 - src.c0 + 1;
  // The source's headings: an empty one is Excel's refusal; one renamed or added (which Excel's refresh works into
  // the fields) isn't done here.
  const names = cacheFields.map((f) => attr(f, 'name'));
  const heads = Array.from({ length: width }, (_, k) => cellValue(sheet.cells.get(src.r0)?.get(src.c0 + k)));
  if (heads.some((h) => h == null || h === '')) return { error: PIVOT_MESSAGES.fieldName };
  if (heads.length !== names.length || heads.some((h, k) => String(h).toLowerCase() !== String(names[k]).toLowerCase())) return { error: PIVOT_MESSAGES.unsupported, why: 'a heading of its source renamed' };
  const records = [];
  for (let r = src.r0 + 1; r <= src.r1; r++) records.push(Array.from({ length: width }, (_, k) => cellValue(sheet.cells.get(r)?.get(src.c0 + k))));
  // The old shared items (what each pivot field's items point at), and which fields share their items.
  const oldShared = cacheFields.map((f) => {
    const s = kid(f, 'sharedItems', ns);
    return s ? [...s.childNodes].filter((n) => n.nodeType === 1).map(itemValue) : [];
  });
  const onAxis = new Set(tables.flatMap((x) => x.def.fields.map((f, i) => (f.axis || f.items ? i : -1))).filter((i) => i >= 0));
  const isShared = names.map((_, i) => oldShared[i].length > 0 || onAxis.has(i));
  // Shared items again, in the order the values first come in the data.
  const shared = names.map((_, i) => {
    if (!isShared[i]) return null;
    const seen = new Map();
    for (const rec of records) if (!seen.has(keyOf(rec[i]))) seen.set(keyOf(rec[i]), rec[i]);
    return [...seen.values()];
  });
  const sharedIndex = shared.map((s) => s && new Map(s.map((v, k) => [keyOf(v), k])));
  const edits = [];
  for (const t of tables) edits.push(layoutTable(wb, t, records, shared, sharedIndex, oldShared, fresh));
  // The cache: shared items, the record count; the records.
  let cacheOut = cacheText.replace(/ recordCount="\d+"/, ` recordCount="${records.length}"`);
  let k = 0;
  cacheOut = cacheOut.replace(/<cacheField\b[^>]*>[\s\S]*?<\/cacheField>/g, (field) => {
    const i = k++;
    const values = records.map((rec) => rec[i]);
    const body = shared[i] ? `<sharedItems${sharedAttrs(values, shared[i])}>${shared[i].map(itemXml).join('')}</sharedItems>` : `<sharedItems${sharedAttrs(values, null)}/>`;
    return field.replace(/<sharedItems\b[^>]*?(?:\/>|>[\s\S]*?<\/sharedItems>)/, body);
  });
  pkg.setText(p.cachePath, cacheOut);
  if (p.recordsPath) {
    const old = pkg.text(p.recordsPath);
    const head = old.match(/^[\s\S]*?<pivotCacheRecords\b[^>]*?>/)[0].replace(/ count="\d+"/, ` count="${records.length}"`);
    const body = records.map((rec) => `<r>${rec.map((v, i) => (shared[i] ? `<x v="${sharedIndex[i].get(keyOf(v))}"/>` : itemXml(v))).join('')}</r>`).join('');
    pkg.setText(p.recordsPath, `${head}${body}</pivotCacheRecords>`);
  }
  return { changes: edits };
}

/** One table laid out again: its part rewritten, and the cell changes of its sheet returned ({ si, contents }). */
function layoutTable(wb, t, records, shared, sharedIndex, oldShared, fresh = false) {
  const def = t.def;
  // Each field's items in the order shown: the old ones still in the data, then new ones as they first come; sorted
  // when the field sorts.
  const order = def.fields.map((f, i) => {
    if (!shared[i] || !(f.axis || f.items)) return null;
    const present = new Set(shared[i].map(keyOf));
    const keys = [];
    for (const it of f.items || []) {
      if (it.x == null) continue;
      const v = oldShared[i][it.x];
      if (v !== undefined && present.has(keyOf(v)) && !keys.includes(keyOf(v))) keys.push(keyOf(v));
    }
    for (const v of shared[i]) if (!keys.includes(keyOf(v))) keys.push(keyOf(v));
    const values = keys.map((key) => shared[i][sharedIndex[i].get(key)]);
    // (Fresh: a field just put on, with no items of its own yet, is listed sorted, as Excel lists a new one.)
    if (f.sortType === 'ascending' || (fresh && !f.items && f.sortType !== 'descending')) values.sort(compareItems);
    if (f.sortType === 'descending') values.sort((a, b) => compareItems(b, a));
    return values;
  });
  const pos = order.map((vals) => vals && new Map(vals.map((v, k) => [keyOf(v), k])));
  const rows = def.rowFields;
  const col = def.colFields.find((x) => x !== -2);
  const data = def.dataFields;
  const label = (v) => (v == null ? def.captions.missing ?? '(blank)' : typeof v === 'object' ? v.error : v);
  // Row entries, depth first: each item of the first row field that has records, then its items of the next...
  const entries = [];
  const walk = (depth, recs) => {
    const f = rows[depth];
    for (const v of order[f]) {
      const sub = recs.filter((rec) => keyOf(rec[f]) === keyOf(v));
      if (!sub.length) continue;
      entries.push({ depth, v, recs: sub });
      if (depth + 1 < rows.length) walk(depth + 1, sub);
    }
  };
  if (rows.length) walk(0, records);
  const cols = col == null ? [] : order[col].map((v) => ({ v, recs: records.filter((rec) => keyOf(rec[col]) === keyOf(v)) })).filter((x) => x.recs.length);
  // The cells, relative to the location's first cell.
  const grid = new Map();
  const put = (r, c, v) => { if (v != null) grid.set(`${r},${c}`, v); };
  let height;
  let widthOut;
  let firstHeaderRow;
  let firstDataRow;
  if (col == null) {
    put(0, 0, def.captions.row);
    data.forEach((d, k) => put(0, 1 + k, d.name));
    entries.forEach((e, i) => { put(1 + i, 0, label(e.v)); data.forEach((d, k) => put(1 + i, 1 + k, aggregate(e.recs, d.fld, d.subtotal))); });
    put(1 + entries.length, 0, def.captions.grand);
    data.forEach((d, k) => put(1 + entries.length, 1 + k, aggregate(records, d.fld, d.subtotal)));
    height = entries.length + 2;
    widthOut = data.length + 1;
    firstHeaderRow = data.length > 1 ? 0 : 1;
    firstDataRow = 1;
  } else {
    const d = data[0];
    if (rows.length) put(0, 0, d.name);
    put(0, 1, def.captions.col);
    if (rows.length) put(1, 0, def.captions.row);
    cols.forEach((x, j) => put(1, 1 + j, label(x.v)));
    put(1, 1 + cols.length, def.captions.grand);
    const line = (r, recs) => {
      cols.forEach((x, j) => { const sub = recs.filter((rec) => keyOf(rec[col]) === keyOf(x.v)); if (sub.length) put(r, 1 + j, aggregate(sub, d.fld, d.subtotal)); });
      put(r, 1 + cols.length, aggregate(recs, d.fld, d.subtotal));
    };
    if (rows.length) {
      entries.forEach((e, i) => { put(2 + i, 0, label(e.v)); line(2 + i, e.recs); });
      put(2 + entries.length, 0, def.captions.grand);
      line(2 + entries.length, records);
      height = entries.length + 3;
    } else {
      put(2, 0, d.name);
      line(2, records);
      height = 3;
    }
    widthOut = cols.length + 2;
    firstHeaderRow = 1;
    firstDataRow = 2;
  }
  const at = def.location;
  const area = { r0: at.r0, c0: at.c0, r1: at.r0 + height - 1, c1: at.c0 + widthOut - 1 };
  // The part: location, each field's items, the row and column items.
  const xv = (k) => (k ? `<x v="${k}"/>` : '<x/>');
  const rowItems = rows.length
    ? [...entries.map((e) => `<i${e.depth ? ` r="${e.depth}"` : ''}>${xv(pos[rows[e.depth]].get(keyOf(e.v)))}</i>`), '<i t="grand"><x/></i>']
    : ['<i/>'];
  const colItems = col != null
    ? [...cols.map((x) => `<i>${xv(pos[col].get(keyOf(x.v)))}</i>`), '<i t="grand"><x/></i>']
    : data.length > 1 ? data.map((_, k) => (k ? `<i i="${k}">${xv(k)}</i>` : '<i><x/></i>')) : ['<i/>'];
  let text = wb.pkg.text(t.path);
  text = text.replace(/<location\b[^>]*\/>/, (loc) => loc
    .replace(/ ref="[^"]*"/, ` ref="${rangeName(area)}"`)
    .replace(/ firstHeaderRow="\d+"/, ` firstHeaderRow="${firstHeaderRow}"`)
    .replace(/ firstDataRow="\d+"/, ` firstDataRow="${firstDataRow}"`));
  let f = 0;
  text = text.replace(/<pivotField\b[^>]*?(?:\/>|>[\s\S]*?<\/pivotField>)/g, (field) => {
    const i = f++;
    if (!order[i]) return field;
    const items = `<items count="${order[i].length + 1}">${order[i].map((v) => `<item x="${sharedIndex[i].get(keyOf(v))}"/>`).join('')}<item t="default"/></items>`;
    if (/<items\b/.test(field)) return field.replace(/<items\b[^>]*?(?:\/>|>[\s\S]*?<\/items>)/, items);
    return field.endsWith('/>') ? `${field.slice(0, -2)}>${items}</pivotField>` : field.replace(/^<pivotField\b[^>]*>/, (open) => open + items);
  });
  const block = (name, list) => `<${name} count="${list.length}">${list.join('')}</${name}>`;
  // (Where a part has none yet, as a new one hasn't: after the element the schema puts first among those there.)
  for (const [name, list, after] of [['rowItems', rowItems, ['</rowFields>', '</pivotFields>']], ['colItems', colItems, ['</colFields>', '</rowItems>', '</rowFields>', '</pivotFields>']]]) {
    const re = new RegExp(`<${name}\\b[^>]*?(?:/>|>[\\s\\S]*?</${name}>)`);
    if (re.test(text)) text = text.replace(re, block(name, list));
    else {
      const end = after.find((tag) => text.includes(tag));
      text = text.replace(end, `${end}${block(name, list)}`);
    }
  }
  wb.pkg.setText(t.path, text);
  // The cells: the new layout's, and the old location's emptied where the new one has nothing.
  const sheet = wb.sheet(t.si);
  const contents = [];
  const span = { r0: Math.min(at.r0, area.r0), c0: Math.min(at.c0, area.c0), r1: Math.max(at.r1, area.r1), c1: Math.max(at.c1, area.c1) };
  for (let r = span.r0; r <= span.r1; r++) {
    for (let c = span.c0; c <= span.c1; c++) {
      const inside = r <= area.r1 && c <= area.c1;
      const wasInside = r <= at.r1 && c <= at.c1;
      if (!inside && !wasInside) continue;
      let v = inside ? grid.get(`${r - area.r0},${c - area.c0}`) ?? null : null;
      if (v && typeof v === 'object') v = v.error;
      const cell = sheet.cells.get(r)?.get(c);
      const now = cell ? cellValue(cell) : null;
      const same = (now && typeof now === 'object' ? now.error : now) === v;
      if (!same || (cell && cell.f != null)) contents.push({ r, c, value: v });
    }
  }
  const labels = rows.length ? { r0: area.r0 + firstDataRow, c0: area.c0, r1: area.r1, c1: area.c0 } : null;
  return { si: t.si, contents, area, labels };
}

