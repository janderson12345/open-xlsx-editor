// Sparklines made and changed as Excel makes and changes them (tools/excel-compare/excel-sparkedit.ps1,
// test/sparkedit-excel.json, docs/SPARKLINES.md): Insert > Sparklines' groups, and the Sparkline tab's edits (type,
// points shown, colours, line weight, axes, empty and hidden cells, Edit Data, Group, Ungroup, Clear), written into
// the worksheet's x14:sparklineGroups. Each function changes the sheet's document; the workbook reads it again.
import { kids, kid, attr } from './xml.js';
import { X14, XM } from './cfext.js';
import { EXT_SPARKLINES, groupElements, readGroup } from './sparklines.js';
import { parseRef, colName } from './refs.js';
import { sheetPrefix } from '../calc/tokens.js';
import { styleColors } from './sparklinestyles.js';

const XMLNS = 'http://www.w3.org/2000/xmlns/';
/** The group's attributes in the order Excel writes them (the schema's), each only when it isn't its default. */
const ATTRS = ['manualMax', 'manualMin', 'lineWeight', 'type', 'dateAxis', 'displayEmptyCellsAs', 'markers', 'high', 'low', 'first', 'last', 'negative',
  'displayXAxis', 'displayHidden', 'minAxisType', 'maxAxisType', 'rightToLeft'];
const COLORS = ['Series', 'Negative', 'Axis', 'Markers', 'First', 'Last', 'High', 'Low'];
/** A new group's colours as Excel writes them through Range.SparklineGroups.Add (sparklines-xml-excel.json). */
export const NEW_COLORS = { series: 'FF376092', negative: 'FFD00000', axis: 'FF000000', markers: 'FFD00000', first: 'FFD00000', last: 'FFD00000', high: 'FFD00000', low: 'FFD00000' };

export class SparklineError extends Error {}

const cellName = (r, c) => `${colName(c)}${r + 1}`;
const rangeText = (r0, c0, r1, c1, pair) => (!pair && r0 === r1 && c0 === c1 ? cellName(r0, c0) : `${cellName(r0, c0)}:${cellName(r1, c1)}`);

/**
 * The sparklines a location ({ r0, c0, r1, c1 }) and a data range ({ sheet, range }) make, as Excel's Add makes them:
 * [{ r, c, source }]. The location is one row or one column of n cells; with one cell, the data is one row or column
 * and the sparkline takes it all; with n, the data is cut into its rows when it has n of them (B2:M3 into O2:O3, and
 * a 3 by 3 square into rows whichever way the location runs), else into its columns when it has n (B2:M3 into
 * B6:M6). Pieces of one cell are written B2:B2, as Excel writes them. Anything else is refused (Excel: 0x800A03EC).
 */
export function sparkPieces(loc, data) {
  const n = (loc.r1 - loc.r0 + 1) * (loc.c1 - loc.c0 + 1);
  if (loc.r0 !== loc.r1 && loc.c0 !== loc.c1) throw new SparklineError('Location reference is not valid.');
  if (!data?.range) throw new SparklineError('Data Source reference is not valid.');
  const { r0, c0, r1, c1 } = data.range;
  const rows = r1 - r0 + 1;
  const cols = c1 - c0 + 1;
  const prefix = data.sheet ? sheetPrefix(data.sheet) : '';
  const cells = [];
  for (let r = loc.r0; r <= loc.r1; r++) for (let c = loc.c0; c <= loc.c1; c++) cells.push({ r, c });
  if (n === 1) {
    if (rows !== 1 && cols !== 1) throw new SparklineError('Location reference or data range is not valid.');
    return [{ ...cells[0], source: prefix + rangeText(r0, c0, r1, c1, false) }];
  }
  if (rows === n) return cells.map((at, k) => ({ ...at, source: prefix + rangeText(r0 + k, c0, r0 + k, c1, true) }));
  if (cols === n) return cells.map((at, k) => ({ ...at, source: prefix + rangeText(r0, c0 + k, r1, c0 + k, true) }));
  throw new SparklineError('Location reference or data range is not valid.');
}

/** The x14:sparklineGroups element, made (with its extLst and ext, the extLst last in the worksheet) if `make`. */
function groupsElement(doc, make) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  let list = kid(root, 'extLst', ns);
  let ext = kids(list, 'ext', ns).find((e) => attr(e, 'uri') === EXT_SPARKLINES);
  let groups = kid(ext, 'sparklineGroups', X14);
  if (groups || !make) return groups;
  if (!list) { list = doc.createElementNS(ns, root.prefix ? `${root.prefix}:extLst` : 'extLst'); root.appendChild(list); }
  if (!ext) {
    ext = doc.createElementNS(ns, root.prefix ? `${root.prefix}:ext` : 'ext');
    ext.setAttribute('uri', EXT_SPARKLINES);
    ext.setAttributeNS(XMLNS, 'xmlns:x14', X14);
    list.appendChild(ext);
  }
  groups = doc.createElementNS(X14, 'x14:sparklineGroups');
  groups.setAttributeNS(XMLNS, 'xmlns:xm', XM);
  ext.appendChild(groups);
  return groups;
}

const x14 = (doc, like, name) => doc.createElementNS(X14, `${like?.prefix || 'x14'}:${name}`);
const xm = (doc, name, text) => { const e = doc.createElementNS(XM, `xm:${name}`); e.textContent = text; return e; };
const sparkEls = (g) => kids(kid(g, 'sparklines', X14), 'sparkline', X14);
const cellOf = (s) => parseRef((kid(s, 'sqref', XM)?.textContent || '').trim());
const inAreas = (at, areas) => !!at && areas.some((a) => at.r >= a.r0 && at.r <= a.r1 && at.c >= a.c0 && at.c <= a.c1);

/** The group drawn in cell (r, c), or null. */
export function groupAt(doc, r, c) {
  const root = doc.documentElement;
  return groupElements(root, root.namespaceURI).find((g) => sparkEls(g).some((s) => { const at = cellOf(s); return at && at.r === r && at.c === c; })) || null;
}

/** A sparkline element (source, cell) in group `g`'s namespace prefixes. */
function sparkEl(doc, g, { r, c, source }) {
  const s = x14(doc, g, 'sparkline');
  s.appendChild(xm(doc, 'f', source));
  s.appendChild(xm(doc, 'sqref', cellName(r, c)));
  return s;
}

/** Group `g`'s sparklines in `areas` taken out (a group left empty goes too); the elements taken. */
function takeOut(g, areas) {
  const out = sparkEls(g).filter((s) => inAreas(cellOf(s), areas));
  for (const s of out) s.parentNode.removeChild(s);
  if (!sparkEls(g).length) g.parentNode.removeChild(g);
  return out;
}

/** Sparklines in `areas` taken out of every group. */
function clearAt(doc, areas) {
  const root = doc.documentElement;
  for (const g of groupElements(root, root.namespaceURI)) takeOut(g, areas);
}

/** A copy of group `g`'s settings with no sparklines, put first (Excel lists the newest group first). */
function copyGroup(doc, g) {
  const copy = g.cloneNode(true);
  const list = kid(copy, 'sparklines', X14);
  while (list.firstChild) list.removeChild(list.firstChild);
  for (const a of [...copy.attributes]) if (/uid$/.test(a.name)) copy.removeAttributeNode(a);
  g.parentNode.insertBefore(copy, g.parentNode.firstChild);
  return copy;
}

/**
 * Insert > Sparklines: a new group of `type` ('line', 'column', 'stacked') at `loc` over `data` (sparkPieces), first
 * in the list, with the settings and colours Excel's Add gives (empty cells as gaps; sparklines-xml-excel.json); a
 * sparkline already in one of its cells is taken out of its group (Excel's Add over O3 of O2:O4: O2 and O4 stay).
 */
export function addSparklines(doc, type, loc, data) {
  const pieces = sparkPieces(loc, data);
  clearAt(doc, [loc]);
  const groups = groupsElement(doc, true);
  const g = x14(doc, null, 'sparklineGroup');
  if (type !== 'line') g.setAttribute('type', type);
  g.setAttribute('displayEmptyCellsAs', 'gap');
  for (const n of COLORS) { const e = x14(doc, null, `color${n}`); e.setAttribute('rgb', NEW_COLORS[n.toLowerCase()]); g.appendChild(e); }
  const list = x14(doc, null, 'sparklines');
  for (const p of pieces) list.appendChild(sparkEl(doc, g, p));
  g.appendChild(list);
  groups.insertBefore(g, groups.firstChild);
  return pieces.length;
}

/** A group's settings (sparklines.js readGroup's names) as Excel's attributes, defaults left out. */
function attrsOf(o) {
  const custom = (t) => o[t] === 'custom';
  return {
    manualMax: custom('maxType') ? o.max : null, manualMin: custom('minType') ? o.min : null,
    lineWeight: o.lineWeight !== 0.75 ? o.lineWeight : null, type: o.type !== 'line' ? o.type : null, dateAxis: o.dates ? 1 : null,
    displayEmptyCellsAs: o.empty !== 'zero' ? o.empty : null,
    markers: o.markers ? 1 : null, high: o.high ? 1 : null, low: o.low ? 1 : null, first: o.first ? 1 : null, last: o.last ? 1 : null, negative: o.negative ? 1 : null,
    displayXAxis: o.axis ? 1 : null, displayHidden: o.hidden ? 1 : null,
    minAxisType: o.minType !== 'individual' ? o.minType : null, maxAxisType: o.maxType !== 'individual' ? o.maxType : null, rightToLeft: o.rightToLeft ? 1 : null,
  };
}

/**
 * The Sparkline tab's settings for the group drawn in (r, c): `changes` over its own, in readGroup's names (type,
 * markers, high, low, first, last, negative, lineWeight, axis, empty, hidden, rightToLeft, dates (a reference or
 * null), minType, maxType, min, max). Its attributes written again in Excel's order (manualMin="-2" lineWeight="1.5"
 * minAxisType="custom", sparkedit-excel.json "weight"), any it doesn't model (xr2:uid) after them. False if there is
 * no group there.
 */
export function setGroupOptions(doc, r, c, changes) {
  const g = groupAt(doc, r, c);
  if (!g) return false;
  const o = { ...readGroup(g), ...changes };
  const others = [...g.attributes].filter((a) => !ATTRS.includes(a.name) && a.name !== 'xmlns:x14' && a.name !== 'xmlns:xm');
  for (const a of [...g.attributes]) if (ATTRS.includes(a.name) || others.includes(a)) g.removeAttributeNode(a);
  const values = attrsOf(o);
  for (const name of ATTRS) if (values[name] != null) g.setAttribute(name, String(values[name]));
  for (const a of others) g.setAttributeNode(a);
  // (The date axis's range: an xm:f in the group, before its sparklines.)
  const f = kid(g, 'f', XM);
  if (f) g.removeChild(f);
  if (o.dates) g.insertBefore(xm(doc, 'f', o.dates), kid(g, 'sparklines', X14));
  return true;
}

/** A colour of the group drawn in (r, c) (`name`: series, negative, axis, markers, first, last, high, low): { rgb:
 *  'FFRRGGBB' } or { theme, tint }, written as Excel writes them (theme="4" tint="-0.249977111117893" for Accent 1,
 *  Darker 25%; sparkedit-excel.json "themeColor"). */
export function setGroupColor(doc, r, c, name, color) {
  const g = groupAt(doc, r, c);
  return !!g && colorOn(doc, g, name, color);
}

/** Group `g`'s colour `name` set to `color`. */
function colorOn(doc, g, name, color) {
  const key = COLORS.find((n) => n.toLowerCase() === name);
  if (!key) return false;
  let e = kid(g, `color${key}`, X14);
  if (!e) {
    e = x14(doc, g, `color${key}`);
    const next = COLORS.slice(COLORS.indexOf(key) + 1).map((n) => kid(g, `color${n}`, X14)).find(Boolean) || kid(g, 'f', XM) || kid(g, 'sparklines', X14);
    g.insertBefore(e, next);
  }
  for (const a of [...e.attributes]) e.removeAttributeNode(a);
  if (color.rgb) e.setAttribute('rgb', color.rgb.length === 6 ? `FF${color.rgb}` : color.rgb);
  else {
    e.setAttribute('theme', String(color.theme));
    if (color.tint) e.setAttribute('tint', String(color.tint));
  }
  return true;
}

/** Clear: the sparklines in `areas` (Clear Selected Sparklines), or (`whole`) every group with one there (Clear
 *  Selected Sparkline Groups); a group left empty goes. The number of sparklines taken out. */
export function clearSparklines(doc, areas, whole = false) {
  const root = doc.documentElement;
  let count = 0;
  for (const g of groupElements(root, root.namespaceURI)) {
    if (whole && sparkEls(g).some((s) => inAreas(cellOf(s), areas))) { count += sparkEls(g).length; g.parentNode.removeChild(g); }
    else if (!whole) count += takeOut(g, areas).length;
  }
  return count;
}

/** Group: the sparklines in `areas` moved into the group drawn in (r, c), after its own, taking its settings (Excel's
 *  Group with O6 for O3 and O6: O3 joins O6's group, O2 and O4 stay in theirs). */
export function groupSparklines(doc, areas, r, c) {
  const into = groupAt(doc, r, c);
  if (!into) return false;
  const root = doc.documentElement;
  const moved = [];
  for (const g of groupElements(root, root.namespaceURI)) if (g !== into) moved.push(...takeOut(g, areas));
  for (const s of moved) kid(into, 'sparklines', X14).appendChild(s);
  return moved.length > 0;
}

/** Ungroup: each sparkline in `areas` a group of its own with its group's settings, put first in turn (Excel's
 *  Ungroup of O2:O3 in O2:O4: groups O3, O2, then O4's). */
export function ungroupSparklines(doc, areas) {
  const root = doc.documentElement;
  let count = 0;
  for (const g of groupElements(root, root.namespaceURI)) {
    const picked = sparkEls(g).filter((s) => inAreas(cellOf(s), areas));
    if (!picked.length || (picked.length === 1 && sparkEls(g).length === 1)) continue;
    for (const s of picked) {
      const own = copyGroup(doc, g);
      s.parentNode.removeChild(s);
      kid(own, 'sparklines', X14).appendChild(s);
      count++;
    }
    if (!sparkEls(g).length) g.parentNode.removeChild(g);
  }
  return count > 0;
}

/** Edit Group Location & Data: the group drawn in (r, c) given the sparklines `loc` and `data` make (sparkPieces), its
 *  settings kept (Excel's Modify of O2:O4 to P2:P4 over C2:M4). Sparklines of other groups in the new cells go. */
export function setGroupData(doc, r, c, loc, data) {
  const g = groupAt(doc, r, c);
  if (!g) return false;
  const pieces = sparkPieces(loc, data);
  const list = kid(g, 'sparklines', X14);
  while (list.firstChild) list.removeChild(list.firstChild);
  const root = doc.documentElement;
  for (const other of groupElements(root, root.namespaceURI)) if (other !== g) takeOut(other, [loc]);
  for (const p of pieces) list.appendChild(sparkEl(doc, g, p));
  return true;
}

/** Edit Single Sparkline's Data: the sparkline in (r, c) reads `source` (one row or column; Excel's ModifySourceData). */
export function setSparklineData(doc, r, c, data) {
  const g = groupAt(doc, r, c);
  if (!g) return false;
  const { r0, c0, r1, c1 } = data.range;
  if (r0 !== r1 && c0 !== c1) throw new SparklineError('Location reference or data range is not valid.');
  const s = sparkEls(g).find((x) => { const at = cellOf(x); return at.r === r && at.c === c; });
  kid(s, 'f', XM).textContent = (data.sheet ? sheetPrefix(data.sheet) : '') + rangeText(r0, c0, r1, c1, false);
  return true;
}

/** A style of the gallery (sparklinestyles.js, 0 to 35) given to the group drawn in (r, c) and every group with a
 *  sparkline in `areas`: its eight colours, nothing else (excel-sparkstyles.ps1: the group's attributes unchanged). */
export function applyStyle(doc, areas, r, c, k) {
  const root = doc.documentElement;
  const colors = styleColors(k);
  const picked = groupElements(root, root.namespaceURI).filter((g) => sparkEls(g).some((s) => { const at = cellOf(s); return inAreas(at, areas) || (!!at && at.r === r && at.c === c); }));
  for (const g of picked) for (const [name, color] of Object.entries(colors)) colorOn(doc, g, name, color);
  return picked.length > 0;
}
