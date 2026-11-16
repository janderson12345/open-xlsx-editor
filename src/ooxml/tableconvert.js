// Convert to Range (Table Design > Tools, ListObject.Unlist) as Excel does it (docs/TABLES.md; measured over COM,
// tools/excel-compare/excel-tables.ps1's unlist case, test/tablemake-excel.json): the table's look kept as the cells'
// own formats, its references made ranges, its part gone.
import { serializeXml, kids, NS } from './xml.js';
import { readTables, tableArea } from './tables.js';
import { colName } from './refs.js';
import { mapFormulas } from './sheetops.js';
import { sheetPrefix } from '../calc/tokens.js';
import { tableLook, tableStyleElements, withOwnFormat } from './tablestyle.js';
import { xfOf } from './styles.js';
import { lineStyle } from './sheet.js';
import { mapStruct } from './tableedit.js';

/**
 * Convert to Range (Unlist, unlist): the table's look kept as the cells' own formats (each cell's fill, font and border
 * from the style, its own format over it), its references in formulas made absolute ranges with the sheet's name
 * (SUM(Table1[Qty]) is SUM(Sheet1!$D$3:$D$7)), and the table part, its relationship and content type gone.
 */
export function convertToRange(wb, t) {
  const tables = readTables(wb);
  const sheetName = wb.sheets[t.si].name;
  const abs = (a) => `${sheetPrefix(sheetName)}${a.r0 === a.r1 && a.c0 === a.c1 ? `$${colName(a.c0)}$${a.r0 + 1}` : `$${colName(a.c0)}$${a.r0 + 1}:$${colName(a.c1)}$${a.r1 + 1}`}`;
  const was = t.name.toLowerCase();
  mapFormulas(wb, (text, here) => mapStruct(text, (tk) => {
    if (!tk.table || tk.table.toLowerCase() !== was) return null;
    const a = tableArea(tables, tk.table, tk.spec, { si: wb.sheetIndexOf(here), r: 0, c: 0 });
    return a.error ? '#REF!' : abs(a);
  }), { name: t.name });
  keepLook(wb, t);
  const part = wb._part(t.si);
  const doc = part.doc;
  const rid = wb.pkg.readRels(wb.sheets[t.si].path).find((r) => r.target === t.path)?.id;
  for (const el of [...doc.getElementsByTagNameNS('*', 'tablePart')]) {
    if (el.getAttributeNS(NS.r, 'id') === rid) el.parentNode.removeChild(el);
  }
  const wrap = doc.getElementsByTagNameNS('*', 'tableParts')[0];
  if (wrap) {
    const left = wrap.getElementsByTagNameNS('*', 'tablePart').length;
    if (left) wrap.setAttribute('count', String(left)); else wrap.parentNode.removeChild(wrap);
  }
  if (rid) dropRel(wb.pkg, wb.sheets[t.si].path, rid);
  wb.pkg.removePart(t.path);
  wb._reread(t.si, part);
  wb.resetCalc();
}

function dropRel(pkg, partPath, id) {
  const dir = partPath.slice(0, partPath.lastIndexOf('/') + 1);
  const relsPath = `${dir}_rels/${partPath.slice(dir.length)}.rels`;
  const doc = pkg.xml(relsPath);
  if (!doc) return;
  for (const r of kids(doc.documentElement, 'Relationship', NS.rel)) if (r.getAttribute('Id') === id) r.parentNode.removeChild(r);
  pkg.setText(relsPath, serializeXml(doc));
}

/** Each cell of `t` given the look its table style gave it, as formats of its own (convertToRange). */
function keepLook(wb, t) {
  const styles = wb.styles;
  const sheet = wb.sheet(t.si);
  const look = tableLook(t, tableStyleElements(styles, t.style.name), wb.theme, styles.palette, { raw: true });
  const cells = [];
  for (const [k, l] of look) {
    const [r, c] = k.split(',').map(Number);
    const s = sheet.cells.get(r)?.get(c)?.s ?? lineStyle(sheet, r, c) ?? 0;
    const xf = xfOf(styles, s);
    const own = { fill: styles.fills[xf.fillId], font: styles.fonts[xf.fontId], border: styles.borders[xf.borderId] };
    const left = withOwnFormat(l, own, styles.fonts[styles.styleXfs?.[xf.xfId]?.fontId] || styles.normal, wb.theme, styles.palette);
    const change = {};
    const font = {};
    if (left.color && l.raw?.color && left.color === l.color) font.color = l.raw.color;
    if (left.bold !== undefined && left.bold !== !!own.font?.b) font.b = left.bold;
    if (left.italic !== undefined && left.italic !== !!own.font?.i) font.i = left.italic;
    if (Object.keys(font).length) change.font = font;
    if (left.fill && l.raw?.fill && own.fill?.pattern === 'none') change.fill = l.raw.fill;
    const border = {};
    // (Each line once, as Excel writes it (unlist): a cell's top and left, the right only in the last column and the
    // bottom only in the last row.)
    const sides = ['left', 'top', ...(c === t.c1 ? ['right'] : []), ...(r === t.r1 ? ['bottom'] : [])];
    for (const side of sides) {
      if (own.border?.[side]?.style || !l.border?.[side]) continue;
      border[side] = { style: l.border[side].style, color: l.raw?.border?.[side] || { indexed: 64 } };
    }
    if (Object.keys(border).length) change.border = border;
    if (Object.keys(change).length) cells.push({ r, c, s: wb.styleWith(s, change) });
  }
  if (cells.length) wb.setStyles(t.si, cells);
}


