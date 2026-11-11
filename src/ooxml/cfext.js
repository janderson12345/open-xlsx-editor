// Data bars and icon sets as Excel 2010 and later keep them (test/cfbars-excel.json): a data bar is a cfRule of type
// dataBar (its colour, min and max) with an x14:id pointing to its x14:cfRule in the worksheet's extLst, which holds the
// rest (automatic min and max, gradient or solid, border, negative bars' colours, axis); an icon set is a cfRule of type
// iconSet, or, for the sets Excel 2010 added (3 Stars, 3 Triangles, 5 Boxes), an x14:cfRule in the extLst alone.
// Read here into the rules sheet.js reads (sheet.condFormats).
import { kids, kid, attr } from './xml.js';
import { readColor } from './styles.js';
import { parseRange } from './refs.js';

export const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
export const XM = 'http://schemas.microsoft.com/office/excel/2006/main';
/** The worksheet extLst's ext for conditional formats, and a cfRule's ext pointing to its x14 rule. */
export const EXT_CF = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}';
export const EXT_ID = '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}';

const flag = (el, name, dflt) => { const v = attr(el, name); return v == null ? dflt : v === '1' || v === 'true'; };
/** A cfvo: { type, val } (an x14 cfvo's value is its xm:f). */
const cfvo = (el) => ({ type: attr(el, 'type'), val: attr(el, 'val') ?? kid(el, 'f', XM)?.textContent ?? null, gte: flag(el, 'gte', true) });

/** An x14 data bar's settings, over the cfRule's own. */
function x14Bar(el) {
  const vals = kids(el, 'cfvo', X14).map(cfvo);
  return {
    minLength: Number(attr(el, 'minLength') ?? 10), maxLength: Number(attr(el, 'maxLength') ?? 90), gradient: flag(el, 'gradient', true),
    border: flag(el, 'border', false), axis: attr(el, 'axisPosition') || 'automatic', direction: attr(el, 'direction') || 'context',
    negativeSame: flag(el, 'negativeBarColorSameAsPositive', false),
    min: vals[0], max: vals[1],
    fill: readColor(kid(el, 'fillColor', X14)), borderColor: readColor(kid(el, 'borderColor', X14)),
    negativeFill: readColor(kid(el, 'negativeFillColor', X14)), negativeBorder: readColor(kid(el, 'negativeBorderColor', X14)), axisColor: readColor(kid(el, 'axisColor', X14)),
  };
}
/** An icon set (its own, or an x14 one): { set, cfvo, reverse, showValue, icons (custom ones) }. */
export function readIconSet(el, ns) {
  return {
    set: attr(el, 'iconSet') || '3TrafficLights1', reverse: flag(el, 'reverse', false), showValue: flag(el, 'showValue', true),
    cfvo: kids(el, 'cfvo', ns).map(cfvo),
  };
}
/** A cfRule's data bar (its own settings, Excel 2007's: min length 10%, max 90%). */
export function readDataBar(el, ns) {
  const vals = kids(el, 'cfvo', ns).map(cfvo);
  return { minLength: Number(attr(el, 'minLength') ?? 10), maxLength: Number(attr(el, 'maxLength') ?? 90), gradient: true, border: false, axis: 'none', negativeSame: true, min: vals[0], max: vals[1], fill: readColor(kid(el, 'color', ns)), showValue: flag(el, 'showValue', true) };
}
/** A base cfRule's x14 id, when it has one. */
export const extIdOf = (rule, ns) => kids(kid(rule, 'extLst', ns), 'ext', ns).map((e) => kid(e, 'id', X14)?.textContent).find(Boolean) || null;

/** The worksheet's x14 rules: { bars: Map(id => settings), extra: [{ ranges, rules }] } (the extra ones, icon sets
 *  only Excel 2010 has, go with the others). */
export function readExtRules(root, ns) {
  const bars = new Map();
  const extra = [];
  const ext = kids(kid(root, 'extLst', ns), 'ext', ns).find((e) => attr(e, 'uri') === EXT_CF);
  for (const cf of kids(kid(ext, 'conditionalFormattings', X14), 'conditionalFormatting', X14)) {
    const ranges = (kid(cf, 'sqref', XM)?.textContent || '').split(/\s+/).map(parseRange).filter(Boolean);
    const rules = [];
    for (const r of kids(cf, 'cfRule', X14)) {
      const bar = kid(r, 'dataBar', X14);
      if (bar) { bars.set(attr(r, 'id'), x14Bar(bar)); continue; }
      const icons = kid(r, 'iconSet', X14);
      if (icons) rules.push({ type: 'iconSet', priority: Number(attr(r, 'priority') || 0), dxfId: null, stopIfTrue: false, formulas: [], iconSet: readIconSet(icons, X14), ext: attr(r, 'id') });
    }
    if (rules.length) extra.push({ ranges, rules });
  }
  return { bars, extra };
}
