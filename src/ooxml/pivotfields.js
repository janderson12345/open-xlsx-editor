// A PivotTable's fields moved after it is made, as Excel's field list moves them (measured: excel-pivot-move.ps1,
// test/pivot-move-excel.json): each field's place written into the table part (on rows, on columns, a value), the
// row and column lists and the values made again, and the pivot laid out by pivot.js's refresh. A field taken off
// keeps its items, as Excel keeps them; one put on that has none yet gets them sorted, as a new pivot's are; a value
// that stays keeps its name.
import { kids, kid, attr } from './xml.js';
import { readPivots, refreshPivot } from './pivot.js';
import { PIVOT_FUNCTIONS, layoutRefusal } from './pivotnew.js';

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** A pivot's fields and layout as the field list shows them: { names, rows, cols, values: [{ field, fn, name }] }, or
 *  null when it isn't one the field list can change here. */
export function pivotLayout(wb, p) {
  const root = wb.pkg.xml(p.path).documentElement;
  const ns = root.namespaceURI;
  const cache = wb.pkg.xml(p.cachePath).documentElement;
  const cacheFields = kids(kid(cache, 'cacheFields', cache.namespaceURI), 'cacheField', cache.namespaceURI);
  const names = cacheFields.map((f) => attr(f, 'name'));
  // (A field of numbers only, as its cache says: Sum when it becomes a value, as Excel starts it; others Count.)
  const numeric = cacheFields.map((f) => { const s = kid(f, 'sharedItems', cache.namespaceURI); return !!s && attr(s, 'containsNumber') === '1' && attr(s, 'containsString') === '0'; });
  const list = (name) => kids(kid(root, name, ns), 'field', ns).map((f) => Number(attr(f, 'x'))).filter((x) => x >= 0);
  const values = kids(kid(root, 'dataFields', ns), 'dataField', ns).map((d) => ({ field: Number(attr(d, 'fld')), fn: attr(d, 'subtotal') || 'sum', name: attr(d, 'name') }));
  if (values.some((v) => !PIVOT_FUNCTIONS[v.fn])) return null;
  return { names, numeric, rows: list('rowFields'), cols: list('colFields'), values };
}

/** The pivot `p` (readPivots) laid out with `layout` ({ rows, cols, values: [{ field, fn }] }): refreshPivot's
 *  { changes } for the caller to make, or { error } with nothing changed. */
export function movePivotFields(wb, p, layout) {
  const why = layoutRefusal(layout);
  if (why) return { error: why };
  const now = pivotLayout(wb, p);
  if (!now) return { error: "This PivotTable's values can't be changed here yet" };
  const before = wb.snapshot();
  const { rows, cols, values } = layout;
  let text = wb.pkg.text(p.path);
  // Each field's place: axis and dataField set or taken off; its items kept.
  let k = 0;
  text = text.replace(/<pivotField\b([^>]*?)(\/?)>/g, (m, attrs, close) => {
    const i = k++;
    let a = attrs.replace(/ axis="[^"]*"/, '').replace(/ dataField="[^"]*"/, '');
    const axis = rows.includes(i) ? ' axis="axisRow"' : cols.includes(i) ? ' axis="axisCol"' : '';
    const data = values.some((v) => v.field === i) ? ' dataField="1"' : '';
    // (Excel's attribute order: axis, then dataField, before the others.)
    a = `${axis}${data}${a}`;
    return `<pivotField${a}${close}>`;
  });
  // The row and column lists, and the values, made again; the old row and column items go (the refresh writes them).
  for (const name of ['rowFields', 'rowItems', 'colFields', 'colItems', 'dataFields']) {
    text = text.replace(new RegExp(`<${name}\\b[^>]*?(?:/>|>[\\s\\S]*?</${name}>)`), '');
  }
  const fields = (tag, xs) => (xs.length ? `<${tag} count="${xs.length}">${xs.map((x) => `<field x="${x}"/>`).join('')}</${tag}>` : '');
  const colList = values.length > 1 ? [...cols, -2] : cols;
  const kept = (v) => now.values.find((w) => w.field === v.field && w.fn === v.fn)?.name;
  const dataFields = `<dataFields count="${values.length}">${values.map((v) => `<dataField name="${esc(kept(v) || `${PIVOT_FUNCTIONS[v.fn]} of ${now.names[v.field]}`)}" fld="${v.field}"${v.fn === 'sum' ? '' : ` subtotal="${v.fn}"`} baseField="0" baseItem="0"/>`).join('')}</dataFields>`;
  text = text.replace('</pivotFields>', `</pivotFields>${fields('rowFields', rows)}${fields('colFields', colList)}`);
  // (dataFields after the page fields, before formats, styles and the rest, as the schema orders them.)
  const next = /<(formats|conditionalFormats|chartFormats|pivotHierarchies|pivotTableStyleInfo|filters|rowHierarchiesUsage|colHierarchiesUsage|extLst)\b/.exec(text);
  text = next ? text.slice(0, next.index) + dataFields + text.slice(next.index) : text.replace('</pivotTableDefinition>', `${dataFields}</pivotTableDefinition>`);
  wb.pkg.setText(p.path, text);
  const out = refreshPivot(wb, readPivots(wb).find((x) => x.path === p.path), { fresh: true });
  if (out.error) { wb.restore(before); return out; }
  return out;
}
