// Home > Cell Styles > New Cell Style and a style's Modify, as Excel makes them (docs/FORMATS.md; measured over COM,
// tools/excel-compare/excel-newstyle.ps1, test/newstyle-excel.json):
// - New Cell Style is "by example": the style's cellStyleXfs entry takes the active cell's number format, font, fill
//   and border ids and its alignment and protection, each part the Style box leaves out marked applyX="0"; its
//   cellStyle (a name, no builtinId) goes in name order. The cell itself keeps its own style.
// - Modify changes the style's own entry in place, a part changed written as a new font, fill or border (or an equal
//   one found); every cell format of that style then takes each part the style includes that it doesn't mark as its
//   own (applyX="1"), in place too, so the cells using it change with it. A built-in style changed is marked
//   customBuiltin="1".
import { kids, kid, attr, serializeXml } from './xml.js';
import { readXf } from './styles.js';
import { formatsOf } from './stylesedit.js';
import { namedStyle } from './cellstyleedit.js';
import { BUILTIN_CELL_STYLES } from '../engine/cellstyles.js';

/** The Style box's six parts: its key, the xf's apply flag and its id attribute (none: a child element). */
export const STYLE_PARTS = [['number', 'NumberFormat', 'numFmtId'], ['alignment', 'Alignment', null], ['font', 'Font', 'fontId'],
  ['border', 'Border', 'borderId'], ['fill', 'Fill', 'fillId'], ['protection', 'Protection', null]];
const CHILD = { Alignment: 'alignment', Protection: 'protection' };

const cellStyleEls = (f) => kids(kid(f.root, 'cellStyles', f.ns), 'cellStyle', f.ns);
const styleXfEls = (f) => kids(kid(f.root, 'cellStyleXfs', f.ns), 'xf', f.ns);
const cellXfEls = (f) => kids(kid(f.root, 'cellXfs', f.ns), 'xf', f.ns);

/** Why a new style can't take `name` (Excel's own refusals), or null. A built-in style's name is taken too. */
export function styleNameRefusal(wb, name, keep = null) {
  const n = String(name || '').trim();
  if (!n) return 'Type a name for the style.';
  if (n.length > 255) return 'The style name is too long.';
  const taken = [...BUILTIN_CELL_STYLES.map((d) => d.name), ...(wb.stylesDoc ? cellStyleEls(formatsOf(wb)).map((el) => attr(el, 'name')) : [])];
  if (taken.some((t) => t && t.toLowerCase() === n.toLowerCase() && t !== keep)) return 'A style with this name already exists.';
  return null;
}

/** Style `name` as the Style box shows it: { name, builtin, xfId, xf (styles.js readXf), includes: { number, ... } },
 *  or null when the workbook has no style of that name and it isn't one of Excel's built-in ones. */
export function styleInfo(wb, name) {
  const def = BUILTIN_CELL_STYLES.find((d) => d.name === name);
  const f = formatsOf(wb);
  const found = cellStyleEls(f).find((el) => attr(el, 'name') === name || (def && attr(el, 'builtinId') === String(def.builtinId)));
  if (!found && !def) return null;
  // (A built-in style the workbook hasn't got: as Excel would add it, without adding it yet.)
  const el = found ? styleXfEls(f)[Number(attr(found, 'xfId') || 0)] : null;
  const includes = Object.fromEntries(STYLE_PARTS.map(([k, p]) => [k, el ? attr(el, `apply${p}`) !== '0' : !def.off.includes(p)]));
  return { name, builtin: !!def || attr(found, 'builtinId') != null, xfId: found ? Number(attr(found, 'xfId') || 0) : null, xf: el ? readXf(el, f.ns, wb.styles) : null, includes };
}

/** The include flags written on a style's xf: none for a part included, applyX="0" for one left out, in Excel's order
 *  (NumberFormat, Font, Fill, Border, Alignment, Protection). */
function setIncludes(el, includes) {
  const key = Object.fromEntries(STYLE_PARTS.map(([k, p]) => [p, k]));
  for (const p of ['NumberFormat', 'Font', 'Fill', 'Border', 'Alignment', 'Protection']) {
    el.removeAttribute(`apply${p}`);
    if (includes[key[p]] === false) el.setAttribute(`apply${p}`, '0');
  }
}

/** Entry `i` of `list` replaced by `el` in place, the writer's lookup and the model kept in step. */
function replaceAt(f, list, i, el) {
  const old = kids(kid(f.root, list, f.ns), 'xf', f.ns)[i];
  const oldKey = serializeXml(old);
  if (f.seen[list].get(oldKey) === i) f.seen[list].delete(oldKey);
  old.parentNode.replaceChild(el, old);
  const key = serializeXml(el);
  if (!f.seen[list].has(key)) f.seen[list].set(key, i);
  const read = readXf(el, f.ns, f.wb.styles);
  if (list === 'cellXfs') f.wb.styles.cellXfs[i] = read;
  else {
    (f.wb.styles.styleXfs ||= [])[i] = read;
    if (i === 0) f.wb.styles.normal = f.wb.styles.fonts[read.fontId];
  }
  f.wb.dirty.add(f.wb.stylesPath);
}

/** The Format Cells box's change for one format (formatops.js formatCells's, its Border tab's outline as the sides). */
function oneCell(change) {
  const { borders, ...rest } = change || {};
  if (!borders) return rest;
  const o = borders.outline || {};
  const border = {};
  for (const k of ['left', 'right', 'top', 'bottom']) if (o[k] !== undefined) border[k] = o[k];
  for (const k of ['diagonal', 'diagonalUp', 'diagonalDown']) if (borders[k] !== undefined) border[k] = borders[k];
  return Object.keys(border).length ? { ...rest, border } : rest;
}

/**
 * New Cell Style: style `name` made from cell format `s` by example, `includes` ({ number, alignment, font, border,
 * fill, protection }, false for a part left out) and `change` (what the Style box's Format... changed, the Format
 * Cells box's change). Returns { xfId } or { error } (Excel's refusal of the name).
 */
export function newCellStyle(wb, s, name, includes = {}, change = null) {
  const error = styleNameRefusal(wb, name);
  if (error) return { error };
  const f = formatsOf(wb);
  const from = f.copy('cellXfs', 'xf', s);
  const xf = f.doc.createElementNS(f.ns, f.root.prefix ? `${f.root.prefix}:xf` : 'xf');
  for (const k of ['numFmtId', 'fontId', 'fillId', 'borderId']) xf.setAttribute(k, attr(from, k) || '0');
  for (const tag of ['alignment', 'protection']) { const c = kid(from, tag, f.ns); if (c) xf.appendChild(c.cloneNode(true)); }
  setIncludes(xf, includes);
  const parent = f.list('cellStyleXfs');
  parent.appendChild(xf);
  const xfId = styleXfEls(f).length - 1;
  parent.setAttribute('count', String(xfId + 1));
  f.seen.cellStyleXfs.set(serializeXml(xf), xfId);
  (wb.styles.styleXfs ||= [])[xfId] = readXf(xf, f.ns, wb.styles);
  const cs = f.doc.createElementNS(f.ns, f.root.prefix ? `${f.root.prefix}:cellStyle` : 'cellStyle');
  cs.setAttribute('name', String(name).trim());
  cs.setAttribute('xfId', String(xfId));
  // (In name order, as Excel keeps them: Mine before Normal.)
  const list = f.list('cellStyles');
  const lower = String(name).trim().toLowerCase();
  list.insertBefore(cs, kids(list, 'cellStyle', f.ns).find((el) => (attr(el, 'name') || '').toLowerCase() > lower) || null);
  list.setAttribute('count', String(kids(list, 'cellStyle', f.ns).length));
  wb.dirty.add(wb.stylesPath);
  if (change && Object.keys(change).length) modifyCellStyle(wb, String(name).trim(), change, includes);
  return { xfId };
}

/**
 * Modify: style `name` (the workbook's own, or a built-in one, added first if the workbook hasn't got it) given
 * `change` (the Format Cells box's) and `includes`, and renamed to `rename` (a style of the workbook's own). Every cell
 * format of the style takes the parts it doesn't hold as its own. Returns { xfId, formats } (the cell formats changed,
 * by index) or { error }.
 */
export function modifyCellStyle(wb, name, change = null, includes = null, rename = null) {
  const info = styleInfo(wb, name);
  if (!info) return { error: `There is no style named ${name}.` };
  if (rename && rename !== name) {
    if (info.builtin) return { error: "A built-in style can't be renamed." };
    const error = styleNameRefusal(wb, rename, name);
    if (error) return { error };
  }
  const f = formatsOf(wb);
  const { xfId, el: styleEl } = namedStyle(f, name);
  const before = styleEl.cloneNode(true);
  const now = f.changed(styleEl.cloneNode(true), readXf(styleEl, f.ns, wb.styles), oneCell(change));
  // (The changed() marks applied what it changed, as a cell's format does; a style's xf marks only what it leaves out.)
  setIncludes(now, includes || info.includes);
  replaceAt(f, 'cellStyleXfs', xfId, now);
  const cs = cellStyleEls(f).find((el) => Number(attr(el, 'xfId') || 0) === xfId);
  if (cs && info.builtin && attr(cs, 'customBuiltin') !== '1') cs.setAttribute('customBuiltin', '1');
  if (cs && rename && rename !== name) cs.setAttribute('name', String(rename).trim());
  // (Each part of the style that changed, or that it now includes, given to the cell formats that follow it.)
  const moved = STYLE_PARTS.filter(([, p, id]) => {
    if (attr(now, `apply${p}`) === '0') return false;
    if (attr(before, `apply${p}`) === '0') return true;
    if (id) return (attr(before, id) || '0') !== (attr(now, id) || '0');
    const a = kid(before, CHILD[p], f.ns);
    const b = kid(now, CHILD[p], f.ns);
    return (a ? serializeXml(a) : '') !== (b ? serializeXml(b) : '');
  });
  const formats = [];
  cellXfEls(f).forEach((el, i) => {
    if (Number(attr(el, 'xfId') || 0) !== xfId) return;
    const parts = moved.filter(([, p]) => attr(el, `apply${p}`) !== '1');
    if (!parts.length) return;
    const next = el.cloneNode(true);
    for (const [, p, id] of parts) {
      if (id) { next.setAttribute(id, attr(now, id) || '0'); continue; }
      const old = kid(next, CHILD[p], f.ns);
      const src = kid(now, CHILD[p], f.ns);
      if (old) next.removeChild(old);
      if (src) {
        const at = CHILD[p] === 'alignment' ? kid(next, 'protection', f.ns) || kid(next, 'extLst', f.ns) : kid(next, 'extLst', f.ns);
        next.insertBefore(src.cloneNode(true), at || null);
      }
    }
    replaceAt(f, 'cellXfs', i, next);
    formats.push(i);
  });
  wb.dirty.add(wb.stylesPath);
  return { xfId, formats };
}
