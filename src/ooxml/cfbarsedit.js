// Data bars and icon sets added as Home > Conditional Formatting adds them (measured over COM:
// tools/excel-compare/excel-cf-bars.ps1, test/cfbars-excel.json). A data bar: a cfRule of type dataBar (min, max, its
// colour) whose x14:id points to its x14:cfRule in the worksheet's extLst (automatic min and max, a gradient bar with
// a border in its colour or a solid one without, negative bars red, a black axis). An icon set: a cfRule of type iconSet
// (the set's thresholds in percent: thirds, quarters or fifths); Excel 2010's sets (3 Stars, 3 Triangles, 5 Boxes) an
// x14:cfRule in the extLst alone.
import { kids, kid, parseXml } from './xml.js';
import { X14, XM, EXT_CF, EXT_ID } from './cfext.js';
import { ICON_COUNTS } from './cfbars.js';

/** The menu's data bar colours (Blue, Green, Red, Orange, Light Blue, Purple). */
export const BAR_COLORS = { blue: 'FF638EC6', green: 'FF63C384', red: 'FFFF555A', orange: 'FFFFB628', lightBlue: 'FF008AEF', purple: 'FFD6007B' };
/** The icon sets Excel 2010 added, kept in the extLst only. */
const X14_ONLY = new Set(['3Stars', '3Triangles', '5Boxes']);
const hex = (n) => Array.from({ length: n }, () => Math.floor(Math.random() * 16).toString(16).toUpperCase()).join('');
const guid = () => `{${hex(8)}-${hex(4)}-4${hex(3)}-${'89AB'[Math.floor(Math.random() * 4)]}${hex(3)}-${hex(12)}}`;
const thresholds = (set) => { const n = ICON_COUNTS[set] || 3; return Array.from({ length: n }, (_, i) => Math.round((i * 100) / n)); };

/** The base cfRule (or null for an x14-only icon set) and the x14 cfRule (or null) for `rule`: { kind: 'dataBar',
 *  color (ARGB), solid } or { kind: 'iconSet', set }. */
export function barRuleXml(rule, priority, id = guid()) {
  if (rule.kind === 'dataBar') {
    const c = rule.color || BAR_COLORS.blue;
    const base = `<cfRule type="dataBar" priority="${priority}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${c}"/></dataBar><extLst><ext uri="${EXT_ID}" xmlns:x14="${X14}"><x14:id>${id}</x14:id></ext></extLst></cfRule>`;
    const body = rule.solid
      ? '<x14:dataBar minLength="0" maxLength="100" gradient="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar>'
      : `<x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:borderColor rgb="${c}"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar>`;
    return { base, x14: `<x14:cfRule type="dataBar" id="${id}">${body}</x14:cfRule>` };
  }
  const set = rule.set || '3TrafficLights1';
  const ts = thresholds(set);
  if (X14_ONLY.has(set)) {
    return { base: null, x14: `<x14:cfRule type="iconSet" priority="${priority}" id="${id}"><x14:iconSet iconSet="${set}">${ts.map((t) => `<x14:cfvo type="percent"><xm:f>${t}</xm:f></x14:cfvo>`).join('')}</x14:iconSet></x14:cfRule>` };
  }
  return { base: `<cfRule type="iconSet" priority="${priority}"><iconSet${set === '3TrafficLights1' ? '' : ` iconSet="${set}"`}>${ts.map((t) => `<cfvo type="percent" val="${t}"/>`).join('')}</iconSet></cfRule>`, x14: null };
}

/** The worksheet's x14 cfRules (each conditionalFormatting's), to renumber their priorities with the others. */
export const x14Rules = (doc) => [...doc.getElementsByTagNameNS(X14, 'cfRule')];

/** Put an x14 cfRule for `sqref` in the worksheet's extLst (its conditionalFormattings ext, made first when it has
 *  none, as Excel orders its exts). */
export function addX14Rule(doc, x14, sqref) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const p = root.prefix ? `${root.prefix}:` : '';
  let extLst = kid(root, 'extLst', ns);
  if (!extLst) { extLst = doc.createElementNS(ns, `${p}extLst`); root.appendChild(extLst); }
  let ext = kids(extLst, 'ext', ns).find((e) => e.getAttribute('uri') === EXT_CF);
  if (!ext) {
    ext = doc.importNode(parseXml(`<w xmlns="${ns}"><ext uri="${EXT_CF}" xmlns:x14="${X14}"><x14:conditionalFormattings/></ext></w>`).documentElement.firstChild, true);
    extLst.insertBefore(ext, extLst.firstChild);
  }
  const list = kid(ext, 'conditionalFormattings', X14);
  const cf = doc.importNode(parseXml(`<w xmlns:x14="${X14}"><x14:conditionalFormatting xmlns:xm="${XM}">${x14}<xm:sqref>${sqref}</xm:sqref></x14:conditionalFormatting></w>`).documentElement.firstChild, true);
  list.appendChild(cf);
}

/** The x14 cfRules (and their conditionalFormattings, when left empty) with these ids taken out of the extLst. */
export function dropX14Rules(doc, ids) {
  for (const r of x14Rules(doc)) {
    if (!ids.has(r.getAttribute('id'))) continue;
    const cf = r.parentNode;
    cf.removeChild(r);
    if (!kids(cf, 'cfRule', X14).length) cf.parentNode.removeChild(cf);
  }
}
