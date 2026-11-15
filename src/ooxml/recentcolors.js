// The colour menus' Recent Colors: the colours picked in More Colors, kept in the workbook as Excel keeps them, in
// styles.xml's <colors><mruColors> (after tableStyles, before extLst; inside colors after indexedColors). Excel saved
// a list put there back as it was (tools/excel-compare/excel-drawborders.ps1); a colour set over COM didn't add one.
import { kids, kid, attr } from './xml.js';
import { stylesPart } from './stylesedit.js';

/** At most ten, as the schema allows (CT_MRUColors). */
export const MAX_RECENT = 10;
const ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];

/** The workbook's recent colours, first the latest: [{ rgb: 'FFRRGGBB' }] (others, theme or indexed, left out). */
export function recentColors(wb) {
  const doc = wb.stylesDoc;
  if (!doc) return [];
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  return kids(kid(kid(root, 'colors', ns), 'mruColors', ns), 'color', ns).map((el) => attr(el, 'rgb')).filter(Boolean).map((rgb) => ({ rgb: rgb.toUpperCase() }));
}

/** `rgb` ('FFRRGGBB') put first in the recent colours, once, the oldest past ten let go. (Which end Excel puts the
 *  latest at isn't measured: the dialog that adds them can't be driven over COM.) */
export function addRecentColor(wb, rgb) {
  const doc = stylesPart(wb);
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const make = (name) => doc.createElementNS(ns, root.prefix ? `${root.prefix}:${name}` : name);
  const value = rgb.toUpperCase();
  let colors = kid(root, 'colors', ns);
  if (!colors) {
    colors = make('colors');
    const next = [...root.childNodes].find((n) => n.nodeType === 1 && ORDER.indexOf(n.localName) > ORDER.indexOf('colors'));
    root.insertBefore(colors, next || null);
  }
  let mru = kid(colors, 'mruColors', ns);
  if (!mru) colors.appendChild((mru = make('mruColors')));
  const list = kids(mru, 'color', ns);
  if ((attr(list[0], 'rgb') || '').toUpperCase() === value) return false;
  for (const el of list) if ((attr(el, 'rgb') || '').toUpperCase() === value) mru.removeChild(el);
  const el = make('color');
  el.setAttribute('rgb', value);
  mru.insertBefore(el, mru.firstChild);
  for (const old of kids(mru, 'color', ns).slice(MAX_RECENT)) mru.removeChild(old);
  wb.dirty.add(wb.stylesPath);
  return true;
}
