// GETPIVOTDATA's lookup, as Excel 365 does it (measured over COM on PivotTables Excel made, docs/PIVOT.md): from the
// pivot's cache (its data as of its last refresh), the value of one of its data fields for the field and item pairs
// given, or the grand total for none. Only what the pivot shows can be asked for: the row fields given must be the
// outer ones (Region alone, or Region and Item; not Item alone), the column fields likewise, and the items must be ones
// that come there; anything else is #REF!. A combination shown empty is 0.
import { kids, kid, attr } from './xml.js';
import { readPivots } from './pivot.js';
import { parseRange } from './refs.js';

const AGG = {
  sum: (ns) => ns.reduce((a, b) => a + b, 0),
  count: (ns, all) => all.length,
  countNums: (ns) => ns.length,
  average: (ns) => (ns.length ? ns.reduce((a, b) => a + b, 0) / ns.length : 0),
  max: (ns) => (ns.length ? Math.max(...ns) : 0),
  min: (ns) => (ns.length ? Math.min(...ns) : 0),
  product: (ns) => (ns.length ? ns.reduce((a, b) => a * b, 1) : 0),
};

function itemValue(el) {
  const v = attr(el, 'v');
  switch (el.localName) {
    case 'n': return Number(v);
    case 'b': return v === '1' || v === 'true';
    case 'm': return null;
    case 'e': return { error: v };
    case 'd': return v;
    default: return v ?? '';
  }
}

/** The pivot whose location meets `area` ({ si, r0, c0, r1, c1 }) of the workbook, with what GETPIVOTDATA needs. */
function pivotMeeting(wb, area) {
  for (const p of readPivots(wb)) {
    if (p.si !== area.si) continue;
    const root = wb.pkg.xml(p.path).documentElement;
    const ns = root.namespaceURI;
    const a = parseRange(attr(kid(root, 'location', ns), 'ref') || '');
    if (!a || a.r0 > area.r1 || area.r0 > a.r1 || a.c0 > area.c1 || area.c0 > a.c1) continue;
    const list = (name) => kids(kid(root, name, ns), 'field', ns).map((f) => Number(attr(f, 'x'))).filter((x) => x >= 0);
    const dataFields = kids(kid(root, 'dataFields', ns), 'dataField', ns).map((d) => ({ name: attr(d, 'name') || '', fld: Number(attr(d, 'fld')), fn: attr(d, 'subtotal') || 'sum' }));
    return { p, rows: list('rowFields'), cols: list('colFields'), dataFields };
  }
  return null;
}

/** The cache's field names and records (each a value per field), or null when the file has no records. */
function cacheOf(wb, p) {
  const cache = wb.pkg.xml(p.cachePath)?.documentElement;
  if (!cache || !p.recordsPath) return null;
  const ns = cache.namespaceURI;
  const fields = kids(kid(cache, 'cacheFields', ns), 'cacheField', ns);
  const names = fields.map((f) => attr(f, 'name') || '');
  const shared = fields.map((f) => { const s = kid(f, 'sharedItems', ns); return s ? [...s.childNodes].filter((n) => n.nodeType === 1).map(itemValue) : []; });
  const recs = wb.pkg.xml(p.recordsPath)?.documentElement;
  if (!recs) return null;
  const records = kids(recs, 'r', recs.namespaceURI).map((r) => [...r.childNodes].filter((n) => n.nodeType === 1)
    .map((n, k) => (n.localName === 'x' ? shared[k][Number(attr(n, 'v'))] ?? null : itemValue(n))));
  return { names, records };
}

const same = (a, b) => a === b || (typeof a === 'string' && typeof b === 'string' && a.toLowerCase() === b.toLowerCase());

/** An item as given matches a value in the cache: text by text, case aside; a number by number, or by text for a
 *  number given as text ("2023"). */
function matches(given, v) {
  if (v == null || typeof v === 'object' || typeof v === 'boolean' || typeof given === 'boolean') return false;
  if (typeof v === 'number') return typeof given === 'number' ? given === v : String(given) !== '' && Number(given) === v;
  return typeof given === 'string' && given.toLowerCase() === String(v).toLowerCase();
}

/**
 * GETPIVOTDATA: the value of data field `dataName` (its name, "Sum of Sales", or its source field's, "Sales"; case
 * aside) of the pivot meeting `area`, for `pairs` ([[field, item]...]). Returns a number, or the error's text.
 */
export function pivotData(wb, area, dataName, pairs) {
  const found = pivotMeeting(wb, area);
  if (!found) return '#REF!';
  const cache = cacheOf(wb, found.p);
  if (!cache) return '#REF!';
  const want = String(dataName).toLowerCase();
  const data = found.dataFields.find((d) => d.name.toLowerCase() === want) || found.dataFields.find((d) => (cache.names[d.fld] || '').toLowerCase() === want);
  if (!data || !AGG[data.fn]) return '#REF!';
  // (Each field once; the same item twice is the same ask.)
  const asked = new Map();
  for (const [f, item] of pairs) {
    const k = cache.names.findIndex((n) => n.toLowerCase() === String(f).toLowerCase());
    if (k < 0 || !(found.rows.includes(k) || found.cols.includes(k))) return '#REF!';
    if (asked.has(k) && !same(asked.get(k).item, item)) return '#REF!';
    asked.set(k, { item });
  }
  const prefix = (axis) => { const n = axis.filter((k) => asked.has(k)).length; return axis.slice(0, n).every((k) => asked.has(k)); };
  if (!prefix(found.rows) || !prefix(found.cols)) return '#REF!';
  const keep = (rec, axis) => axis.every((k) => !asked.has(k) || matches(asked.get(k).item, rec[k]));
  // (The row path and the column path must each be there; where they cross may be empty.)
  if (!cache.records.some((r) => keep(r, found.rows)) || !cache.records.some((r) => keep(r, found.cols))) return '#REF!';
  const recs = cache.records.filter((r) => keep(r, found.rows) && keep(r, found.cols));
  const all = recs.map((r) => r[data.fld]).filter((v) => v != null && v !== '');
  return AGG[data.fn](all.filter((v) => typeof v === 'number'), all);
}
