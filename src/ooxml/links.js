// A sheet's hyperlinks as Excel writes them (measured over COM, Hyperlinks.Add, and the file Excel saved: docs/LINKS.md):
// <hyperlinks> of <hyperlink ref>, an address outside the workbook as a relationship of the sheet (r:id, its Target
// with TargetMode="External"), a place in the workbook as `location` ('Sheet2'!B5 or a defined name), with `display`
// and `tooltip` when they were given.
import { NS, kids, attr } from './xml.js';
import { parseRange, refName } from './refs.js';
import { clearHyperlinks } from './clearparts.js';

// (The characters Excel escapes in a hyperlink's Target, measured: test/filelinks-excel.json. `&`, `\`, `#`'s place and
// letters outside ASCII stay as they are; a `%` already starting an escape stays, any other `%` is %25.)
const ESCAPED = { ' ': '%20', '[': '%5b', ']': '%5d', '{': '%7b', '}': '%7d', '^': '%5e' };
const UNESCAPED = Object.fromEntries([...Object.entries(ESCAPED).map(([k, v]) => [v, k]), ['%25', '%']]);

/** A hyperlink's address as Excel writes it in the relationship's Target: "My File.xlsx" as "My%20File.xlsx",
 *  "R&D 100%.xlsx" as "R&D%20100%25.xlsx", "Café [1].xlsx" as "Café%20%5b1%5d.xlsx", "a%20b" as it is. */
export const encodeAddress = (a) => String(a).replace(/%(?![0-9a-f]{2})/gi, '%25').replace(/[ [\]{}^]/g, (ch) => ESCAPED[ch]);

/** A Target as Excel's Address reads it back: those escapes undone ("x%41" stays: Excel didn't make it). */
export const decodeAddress = (t) => String(t).replace(/%(20|5b|5d|7b|7d|5e|25)/gi, (m) => UNESCAPED[m.toLowerCase()]);

/** Sheet si's hyperlinks: [{ range: { r0, c0, r1, c1 }, address (outside the workbook) or null, location (in it) or
 *  null, display, tooltip }], in the file's order. */
export function readHyperlinks(wb, si) {
  const { doc } = wb._part(si);
  const ns = doc.documentElement.namespaceURI;
  const list = doc.getElementsByTagNameNS(ns, 'hyperlinks')[0];
  if (!list) return [];
  const rels = wb.pkg.readRels(wb.sheets[si].path);
  return kids(list, 'hyperlink', ns).map((el) => {
    const id = el.getAttributeNS(NS.r, 'id');
    const rel = id ? rels.find((x) => x.id === id) : null;
    return {
      range: parseRange(attr(el, 'ref') || ''),
      address: rel ? decodeAddress(rel.rawTarget) : null,
      location: attr(el, 'location') || null,
      display: attr(el, 'display') || null,
      tooltip: attr(el, 'tooltip') || null,
    };
  }).filter((h) => h.range);
}

/** The hyperlink of cell (r, c) of sheet si, or null (the first, as Excel follows it, when ranges overlap). */
export function hyperlinkAt(wb, si, r, c) {
  return readHyperlinks(wb, si).find(({ range: a }) => r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1) || null;
}

/** Where a hyperlink goes, as Excel's ScreenTip says it: the address (with its #place), or the place in the workbook. */
export function linkTarget(h) {
  if (!h) return '';
  if (h.address) return h.location ? `${h.address}#${h.location}` : h.address;
  return h.location || '';
}

/** Excel's ScreenTip for a hyperlink: its own tooltip, else where it goes; then how to follow it. */
export function screenTip(h) {
  return `${h.tooltip || linkTarget(h)} - Click once to follow. Click and hold to select this cell.`;
}

// (A browser opens these itself; a file's path, relative or not, or any other scheme isn't opened from a page.)
const OPENED = /^(https?:|mailto:|ftp:)/i;

/** What following hyperlink `h` does: { url } to open in the browser, { place } to go to in the workbook, or
 *  { refused: Excel's-style words } for an address a page can't open. */
export function followOf(h) {
  // (Opened as the file writes it, escaped: a "100%" read back from "100%25" goes as 100%25.)
  const url = encodeAddress(h.address || '');
  if (h.address) return OPENED.test(h.address) ? { url: h.location ? `${url}#${h.location}` : url } : { refused: 'Cannot open the specified file.' };
  if (h.location) return { place: h.location };
  return { refused: 'Reference isn\'t valid.' };
}

const LINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
// (CT_Worksheet: hyperlinks come before these.)
const AFTER_LINKS = ['printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches',
  'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls',
  'webPublishItems', 'tableParts', 'extLst'];

/** Give cell (r, c) of sheet si a hyperlink, as Excel's Hyperlinks.Add writes one (measured): to an `address` outside
 *  the workbook, or a `location` in it ('Sheet2'!B5), with a `tooltip` (ScreenTip) and the `text` to display. The
 *  cell's own hyperlink goes first. An empty cell gets the text (else the address or place) as its value; one with
 *  something in it keeps it, and `display` then holds the address. A place in the workbook keeps its text in
 *  `display` too. The cell gets the Hyperlink cell style. */
export function addHyperlink(wb, si, r, c, { address = '', location = '', tooltip = '', text = '' }) {
  clearHyperlinks(wb, si, { r0: r, c0: c, r1: r, c1: c });
  const cell = wb.sheet(si).cells.get(r)?.get(c);
  const filled = cell != null && (cell.v != null || cell.f != null) && cell.v !== '';
  const display = address ? (filled ? address : '') : text || location;
  appendLink(wb, si, { range: { r0: r, c0: c, r1: r, c1: c }, address, location, tooltip, display });
  if (!filled) wb.setCell(si, { r, c }, text || address || location);
  wb.setStyles(si, [{ r, c, s: wb.styleWith(cell?.s || 0, { style: 'Hyperlink' }) }]);
}

/** A <hyperlink> added to sheet si as `h` has it (readHyperlinks' shape): its address a relationship of the sheet,
 *  written as Excel writes it (encodeAddress). */
function appendLink(wb, si, { range, address, location, tooltip, display }) {
  const { doc } = wb._part(si);
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const el = doc.createElementNS(ns, root.prefix ? `${root.prefix}:hyperlink` : 'hyperlink');
  el.setAttribute('ref', range.r0 === range.r1 && range.c0 === range.c1 ? refName(range.r0, range.c0) : `${refName(range.r0, range.c0)}:${refName(range.r1, range.c1)}`);
  if (address) el.setAttributeNS(NS.r, 'r:id', wb.pkg.addRel(LINK, encodeAddress(address), true, wb.sheets[si].path));
  if (location) el.setAttribute('location', location);
  if (tooltip) el.setAttribute('tooltip', tooltip);
  if (display) el.setAttribute('display', display);
  let list = doc.getElementsByTagNameNS(ns, 'hyperlinks')[0];
  if (!list) {
    list = doc.createElementNS(ns, root.prefix ? `${root.prefix}:hyperlinks` : 'hyperlinks');
    root.insertBefore(list, [...root.childNodes].find((n) => n.nodeType === 1 && AFTER_LINKS.includes(n.localName)) || null);
  }
  list.appendChild(el);
  wb.dirty.add(wb.sheets[si].path);
  wb.edited = true;
}

/** Sheet si's hyperlinks inside `area` put at the same cells of sheet `to` (Fill Across Worksheets' All, measured),
 *  after taking away those that meet it there. Returns whether anything changed. */
export function copyHyperlinks(wb, si, to, area) {
  const inside = (a) => a.r0 >= area.r0 && a.r1 <= area.r1 && a.c0 >= area.c0 && a.c1 <= area.c1;
  const links = readHyperlinks(wb, si).filter((h) => inside(h.range));
  const gone = clearHyperlinks(wb, to, area).length > 0;
  for (const h of links) appendLink(wb, to, h);
  return gone || links.length > 0;
}
