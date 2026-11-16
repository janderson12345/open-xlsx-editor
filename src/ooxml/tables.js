// Tables (ListObjects) as the file has them, and what a formula's table reference stands for: Fruit[Qty] the column's
// data rows; Fruit[#All], [#Data], [#Headers], [#Totals], and two of them together; Fruit[[Qty]:[Price]] the columns
// between; Fruit[[#This Row],[Qty]] (typed Fruit[@Qty]) the cell of the formula's own row. Column names compare case
// aside; a ' escapes the next character ([, ], # or ').
import { kids, kid, attr, NS } from './xml.js';
import { parseRange } from './refs.js';

/** Every table of the workbook's worksheets: [{ name, si, r0, c0, r1, c1, headers (rows: 0 or 1), totals (0 or 1),
 *  columns: [names] }]. */
export function readTables(wb) {
  const out = [];
  wb.sheets.forEach((info, si) => {
    if (!info.path || info.kind !== 'worksheet') return;
    for (const rel of wb.pkg.readRels(info.path).filter((r) => r.type === `${NS.r}/table`)) {
      const doc = wb.pkg.xml(rel.target);
      const root = doc?.documentElement;
      const area = root && parseRange(attr(root, 'ref'));
      if (!area) continue;
      const ns = root.namespaceURI;
      const cols = kids(kid(root, 'tableColumns', ns), 'tableColumn', ns);
      // (Its style and Table Style Options: a table without tableStyleInfo, or one without a name, has no style.)
      const info = kid(root, 'tableStyleInfo', ns);
      const on = (name) => attr(info, name) === '1' || attr(info, name) === 'true';
      const filter = kid(root, 'autoFilter', ns);
      out.push({
        name: attr(root, 'displayName') || attr(root, 'name'), si, ...area, path: rel.target,
        id: Number(attr(root, 'id') || 0),
        headers: attr(root, 'headerRowCount') === '0' ? 0 : 1,
        totals: Number(attr(root, 'totalsRowCount') || 0) > 0 ? 1 : 0,
        columns: cols.map((c) => attr(c, 'name') || ''),
        // (Each column's total: totalsRowFunction and totalsRowLabel, kept while the total row is off.)
        totalsRow: cols.map((c) => ({ fn: attr(c, 'totalsRowFunction') || null, label: attr(c, 'totalsRowLabel') })),
        style: { name: attr(info, 'name') || null, first: on('showFirstColumn'), last: on('showLastColumn'), rows: on('showRowStripes'), cols: on('showColumnStripes') },
        // (Its filter buttons: an autoFilter, less the columns whose button is hidden.)
        filter: filter ? { hidden: new Set(kids(filter, 'filterColumn', ns).filter((f) => attr(f, 'hiddenButton') === '1').map((f) => Number(attr(f, 'colId')))) } : null,
      });
    }
  });
  return out;
}

/** A specifier's items, each "#Totals", "@Qty" or a column name (unescaped), with "Qty:Price" kept as a span. */
function items(spec) {
  const unescape = (s) => s.replace(/'(.)/g, '$1');
  const s = spec.trim();
  if (!s.startsWith('[')) {
    // (One item, unbracketed: Fruit[Qty], Fruit[#All], Fruit[@Qty].)
    if (s.startsWith('@')) return [{ special: '#this row' }, ...(s.length > 1 ? [{ col: unescape(s.slice(1).replace(/^\[|\]$/g, '')) }] : [])];
    return s.startsWith('#') ? [{ special: s.toLowerCase() }] : [{ col: unescape(s) }];
  }
  const out = [];
  let i = 0;
  const read = () => {
    // ([name], the brackets' text with its escapes.)
    let j = i + 1;
    let t = '';
    while (j < s.length && s[j] !== ']') { if (s[j] === "'" && j + 1 < s.length) { t += s[j + 1]; j += 2; } else t += s[j++]; }
    i = j + 1;
    return t;
  };
  while (i < s.length) {
    const ch = s[i];
    if (ch === ',' || ch === ' ') { i++; continue; }
    if (ch === '@') { out.push({ special: '#this row' }); i++; continue; }
    if (ch !== '[') return null;
    const a = read();
    if (a.startsWith('#')) { out.push({ special: a.toLowerCase() }); continue; }
    if (s[i] === ':' && s[i + 1] === '[') { i++; out.push({ col: a, to: read() }); continue; }
    out.push({ col: a });
  }
  return out;
}

/**
 * What table reference `table`[`spec`] stands for, from cell `here` ({ si, r, c }): { si, r0, c0, r1, c1 }, or
 * { error } (#REF! for a table, column or part there isn't, #VALUE! for This Row outside the table's rows). `table`
 * null: the table `here` is in.
 */
export function tableArea(tables, table, spec, here) {
  const t = table == null
    ? tables.find((x) => x.si === here.si && here.r >= x.r0 && here.r <= x.r1 && here.c >= x.c0 && here.c <= x.c1)
    : tables.find((x) => x.name.toLowerCase() === table.toLowerCase());
  if (!t) return { error: '#REF!' };
  const list = items(spec);
  if (!list) return { error: '#REF!' };
  const first = t.r0 + t.headers;
  const last = t.r1 - t.totals;
  const parts = new Set(list.filter((x) => x.special).map((x) => x.special));
  let r0;
  let r1;
  if (parts.has('#this row')) {
    // (The formula's own row, on whatever sheet it is: implicit intersection, as a range on another sheet meets it.)
    if (here.r < first || here.r > last) return { error: '#VALUE!' };
    r0 = r1 = here.r;
  } else if (parts.has('#all')) { r0 = t.r0; r1 = t.r1; }
  else {
    const has = (p) => parts.has(p);
    const want = { h: has('#headers'), d: has('#data') || (!has('#headers') && !has('#totals')), t: has('#totals') };
    if ((want.h && !t.headers) || (want.t && !t.totals)) return { error: '#REF!' };
    r0 = want.h ? t.r0 : want.d ? first : t.r1;
    r1 = want.t ? t.r1 : want.d ? last : t.r0;
    if (r1 < r0) return { error: '#REF!' };
  }
  const cols = list.filter((x) => x.col != null);
  let c0 = t.c0;
  let c1 = t.c1;
  if (cols.length) {
    const at = (name) => t.columns.findIndex((n) => n.toLowerCase() === name.toLowerCase());
    const a = at(cols[0].col);
    const b = at(cols[cols.length - 1].to ?? cols[cols.length - 1].col);
    if (a < 0 || b < 0) return { error: '#REF!' };
    c0 = t.c0 + Math.min(a, b);
    c1 = t.c0 + Math.max(a, b);
  }
  return { si: t.si, r0, c0, r1, c1 };
}
