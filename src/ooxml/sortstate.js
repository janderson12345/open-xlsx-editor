// The sort a sheet was last sorted by, saved as Excel saves it: its <sortState>. Measured over COM on 17 sorts of a
// made-up table (tools/excel-compare/excel-sortstate.ps1, sortstate-cases.json, test/sortstate-excel.json):
// - the sortState's ref is the rows sorted, the header row left out (A2:C6 of A1:C6), or with Left to Right the
//   columns sorted;
// - one sortCondition per level, its ref the key column's data rows (B2:B6), as Excel's Sort box passes the key, or
//   with Left to Right the key row's cells; descending="1" for Z to A, sortBy="cellColor" or "fontColor" with the
//   colour as a dxf (dxfId), customList="a,b,c" for a custom order;
// - caseSensitive="1" with Match case, columnSort="1" with Left to Right;
// - in a filtered range (sorted from its AutoFilter, as its buttons sort it) the one sortState is the autoFilter's
//   last child; otherwise it follows the autoFilter in the sheet, before the elements CT_Worksheet puts after it;
// - Excel 365 declares the xlrd2 namespace on it, written so here too;
// - the A to Z and Z to A buttons save one in the same form as the Sort box.
// A new sort replaces the sheet's last one. A sort inside a table, whose sortState Excel keeps in the table's part,
// is not written here (the table's is kept as it was).
import { kid, kids, parseXml } from './xml.js';
import { rangeName, colName } from './refs.js';
import { addDxf } from './filterops.js';
import { AFTER_MERGES } from './sheetedit.js';
import { readTables } from './tables.js';

const XLRD2 = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2';
// (The elements CT_Worksheet puts after sortState: a sheet-level one goes before the first of these there.)
const AFTER_SORT = ['dataConsolidate', 'customSheetViews', 'mergeCells', ...AFTER_MERGES];
const escAttr = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** The workbook's body font as a font colour sort's dxf names it (Excel's: every property given, its colour). */
function fontDxfXml(wb, rgb) {
  const f = wb.styles.fonts[0] || {};
  const sroot = wb.stylesDoc.documentElement;
  const first = kids(kid(sroot, 'fonts', sroot.namespaceURI), 'font', sroot.namespaceURI)[0];
  const family = first ? kid(first, 'family', sroot.namespaceURI)?.getAttribute('val') : null;
  return `<dxf><font><b val="0"/><i val="0"/><strike val="0"/><condense val="0"/><extend val="0"/><outline val="0"/>`
    + `<shadow val="0"/><u val="none"/><vertAlign val="baseline"/><sz val="${escAttr(f.size ?? 11)}"/><color rgb="${escAttr(rgb)}"/>`
    + `<name val="${escAttr(f.name || 'Calibri')}"/>${family ? `<family val="${escAttr(family)}"/>` : ''}`
    + `${f.scheme ? `<scheme val="${escAttr(f.scheme)}"/>` : ''}</font></dxf>`;
}
const fillDxfXml = (rgb) => `<dxf><fill><patternFill patternType="solid"><fgColor rgb="${escAttr(rgb)}"/><bgColor indexed="65"/></patternFill></fill></dxf>`;

/**
 * The sortState element's XML for a sort of sheet si (`plan`: { area, keys: [{ at, desc, on, color, customOrder }],
 * header, matchCase, leftToRight }, as the Sort box and quickSort give it). Colour levels add their dxf to the styles.
 */
export function sortStateXml(wb, plan) {
  const { area, keys, header = false, matchCase = false, leftToRight = false } = plan;
  const data = leftToRight ? { ...area, c0: area.c0 + (header ? 1 : 0) } : { ...area, r0: area.r0 + (header ? 1 : 0) };
  const conditions = keys.map((k) => {
    const ref = leftToRight ? `${colName(data.c0)}${k.at + 1}:${colName(data.c1)}${k.at + 1}` : `${colName(k.at)}${data.r0 + 1}:${colName(k.at)}${data.r1 + 1}`;
    const on = (k.on === 'cellColor' || k.on === 'fontColor') && k.color ? k.on : null;
    const dxf = on ? addDxf(wb, on === 'cellColor' ? fillDxfXml(k.color) : fontDxfXml(wb, k.color)) : null;
    return `<sortCondition${k.desc ? ' descending="1"' : ''}${on ? ` sortBy="${on}"` : ''} ref="${ref}"`
      + `${Array.isArray(k.customOrder) ? ` customList="${escAttr(k.customOrder.join(','))}"` : ''}${dxf != null ? ` dxfId="${dxf}"` : ''}/>`;
  }).join('');
  return `<sortState${leftToRight ? ' columnSort="1"' : ''}${matchCase ? ' caseSensitive="1"' : ''} ref="${rangeName(data)}" xmlns:xlrd2="${XLRD2}">${conditions}</sortState>`;
}

/** Sheet si's last sort saved as `plan` (sortStateXml): in its autoFilter when the sort was of the filtered range,
 *  else after it in the sheet; the sheet's earlier one replaced. Not for a sort inside a table. Returns true when
 *  written. */
export function setSortState(wb, si, plan) {
  const { area } = plan;
  const inside = (t) => t.si === si && area.r0 >= t.r0 && area.r1 <= t.r1 && area.c0 >= t.c0 && area.c1 <= t.c1;
  if (readTables(wb).some(inside)) return false;
  const part = wb._part(si);
  const root = part.doc.documentElement;
  const ns = root.namespaceURI;
  const filter = part.model.autoFilter;
  const filtered = !plan.leftToRight && filter && area.r0 >= filter.r0 && area.r1 <= filter.r1 && area.c0 >= filter.c0 && area.c1 <= filter.c1;
  const el = part.doc.importNode(parseXml(`<w xmlns="${ns}">${sortStateXml(wb, plan)}</w>`).documentElement.firstChild, true);
  const own = kid(root, 'sortState', ns);
  const auto = kid(root, 'autoFilter', ns);
  if (filtered && auto) {
    const old = kid(auto, 'sortState', ns);
    if (old) auto.removeChild(old);
    if (own) root.removeChild(own);
    auto.appendChild(el);
  } else if (own) root.replaceChild(el, own);
  else root.insertBefore(el, [...root.childNodes].find((n) => n.nodeType === 1 && AFTER_SORT.includes(n.localName)) || null);
  wb._reread(si, part);
  return true;
}
