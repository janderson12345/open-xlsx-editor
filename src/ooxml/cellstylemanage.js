// Home > Cell Styles' Delete, Duplicate and Merge Styles, as Excel does them (docs/FORMATS.md; measured over COM,
// tools/excel-compare/excel-stylesmore.ps1, test/stylesmore-excel.json):
// - Delete: a style of the workbook's own goes, its cellStyle and its cellStyleXfs entry (the entries after it, and the
//   xfIds pointing at them, one down); a built-in style stays, its cellStyle marked hidden="1", and can't be given to
//   cells again. Each cell format of the style goes back to Normal in place: xfId 0 and Normal's number format, font,
//   fill, border, alignment and protection, but for each part it holds as its own (applyX="1"). Normal can't be deleted.
// - Duplicate: a copy of the style named "<name> 2" (no builtinId), its parts as the style's, each part it leaves out
//   Normal's (Excel's Style box makes the copy by example from a cell of the style alone); the cells keep their style.
// - Merge Styles: each style of the other workbook (its fonts, fills, borders and number formats written into this
//   one) added, and one of the same name (a built-in style by its builtinId) put in place of this workbook's, as Modify
//   puts a style's change, the cells using it changing with it, when "Merge styles that have the same names?" is
//   answered Yes; a built-in style changed in either keeps customBuiltin="1".
import { kids, kid, attr, serializeXml } from './xml.js';
import { readStyles } from './styles.js';
import { formatsOf, forgetFormats } from './stylesedit.js';
import { namedStyle, builtinStyleXf } from './cellstyleedit.js';
import { styleInfo, styleNameRefusal, setIncludes, putStyleXf, modifyCellStyle, STYLE_PARTS } from './cellstyledefine.js';
import { BUILTIN_CELL_STYLES } from '../engine/cellstyles.js';

const ID_PARTS = [['numFmtId', 'NumberFormat'], ['fontId', 'Font'], ['fillId', 'Fill'], ['borderId', 'Border']];
const APPLY = ['NumberFormat', 'Font', 'Fill', 'Border', 'Alignment', 'Protection'];
const IDS = new Set(['numFmtId', 'fontId', 'fillId', 'borderId', 'xfId', ...APPLY.map((p) => `apply${p}`)]);

const make = (f, name) => f.doc.createElementNS(f.ns, f.root.prefix ? `${f.root.prefix}:${name}` : name);
const listOf = (root, list, item) => kids(kid(root, list, root.namespaceURI), item, root.namespaceURI);
const lower = (s) => String(s || '').trim().toLowerCase();
const builtinDef = (name) => BUILTIN_CELL_STYLES.find((d) => lower(d.name) === lower(name));

/** The cellStyle of `name` in a styles part (`root`): a built-in style's by its builtinId, any other's by its name
 *  (capitals aside). */
function findStyle(root, name, builtinId = builtinDef(name)?.builtinId) {
  const list = listOf(root, 'cellStyles', 'cellStyle');
  return (builtinId != null && list.find((el) => attr(el, 'builtinId') === String(builtinId) && attr(el, 'iLevel') == null))
    || list.find((el) => attr(el, 'builtinId') == null && lower(attr(el, 'name')) === lower(name)) || null;
}

/** The styles part changed as a whole (entries taken out, moved or hidden): the model read again and the writer made
 *  anew, as a snapshot put back is (workbookstate.js). */
function reread(wb) {
  forgetFormats(wb);
  Object.assign(wb.styles, readStyles(wb.stylesDoc), { palette: wb.styles.palette });
  wb.dirty.add(wb.stylesPath);
}

/** `cs` put among the cellStyles in name order, as Excel keeps them. */
function placeStyle(f, cs) {
  const list = f.list('cellStyles');
  const name = lower(attr(cs, 'name'));
  list.insertBefore(cs, kids(list, 'cellStyle', f.ns).find((el) => lower(attr(el, 'name')) > name) || null);
  list.setAttribute('count', String(kids(list, 'cellStyle', f.ns).length));
}

/** `el` appended to cellStyleXfs: its index. */
function appendStyleXf(f, el) {
  const parent = f.list('cellStyleXfs');
  parent.appendChild(el);
  const i = kids(parent, 'xf', f.ns).length - 1;
  parent.setAttribute('count', String(i + 1));
  return i;
}

/** Cell format `el` gone back to Normal (style entry `normal`): xfId 0, each part Normal's but those it marks as its
 *  own, in Excel's order (the ids, xfId, the apply flags kept, any other attribute, then the children). */
function toNormal(f, el, normal) {
  const out = make(f, 'xf');
  const own = (p) => attr(el, `apply${p}`) === '1';
  for (const [k, p] of ID_PARTS) out.setAttribute(k, (own(p) ? attr(el, k) : attr(normal, k)) || '0');
  out.setAttribute('xfId', '0');
  for (const p of APPLY) if (own(p)) out.setAttribute(`apply${p}`, '1');
  for (const a of [...el.attributes]) if (!IDS.has(a.name)) out.setAttribute(a.name, a.value);
  for (const [tag, p] of [['alignment', 'Alignment'], ['protection', 'Protection']]) {
    const from = kid(own(p) ? el : normal, tag, f.ns);
    if (from) out.appendChild(from.cloneNode(true));
  }
  const ext = kid(el, 'extLst', f.ns);
  if (ext) out.appendChild(ext.cloneNode(true));
  return out;
}

/**
 * Delete: cell style `name` taken out of the workbook as Excel takes it (see the top of this file). Returns
 * { formats } (the cell formats gone back to Normal, by index) or { error }.
 */
export function deleteCellStyle(wb, name) {
  const def = builtinDef(name);
  if (def?.builtinId === 0 || lower(name) === 'normal') return { error: "The Normal style can't be deleted." };
  const f = formatsOf(wb);
  let cs = findStyle(f.root, name);
  if (!cs && !def) return { error: `There is no style named ${name}.` };
  if (cs && attr(cs, 'hidden') === '1') return { error: `There is no style named ${name}.` };
  // (A built-in style the workbook hasn't got: Excel adds it, hidden.)
  if (!cs) { namedStyle(f, def.name); cs = findStyle(f.root, def.name); }
  const xfId = Number(attr(cs, 'xfId') || 0);
  const normal = listOf(f.root, 'cellStyleXfs', 'xf')[0];
  const formats = [];
  listOf(f.root, 'cellXfs', 'xf').forEach((el, i) => {
    if (Number(attr(el, 'xfId') || 0) !== xfId) return;
    el.parentNode.replaceChild(toNormal(f, el, normal), el);
    formats.push(i);
  });
  if (attr(cs, 'builtinId') != null) cs.setAttribute('hidden', '1');
  else {
    const styles = cs.parentNode;
    styles.removeChild(cs);
    styles.setAttribute('count', String(kids(styles, 'cellStyle', f.ns).length));
    // (Its entry out, and the ones after it a place down, with the xfIds pointing at them: none is left on it.)
    const xfs = listOf(f.root, 'cellStyleXfs', 'xf');
    const others = listOf(f.root, 'cellStyles', 'cellStyle').filter((el) => Number(attr(el, 'xfId') || 0) === xfId);
    if (xfId > 0 && !others.length) {
      xfs[xfId].parentNode.removeChild(xfs[xfId]);
      kid(f.root, 'cellStyleXfs', f.ns).setAttribute('count', String(xfs.length - 1));
      for (const el of [...listOf(f.root, 'cellXfs', 'xf'), ...listOf(f.root, 'cellStyles', 'cellStyle')]) {
        const at = Number(attr(el, 'xfId') || 0);
        if (at > xfId) el.setAttribute('xfId', String(at - 1));
      }
    }
  }
  reread(wb);
  return { formats };
}

/** Duplicate's name for a copy of style `name`: "<name> 2", or the first number after it the workbook hasn't got. */
export function duplicateName(wb, name) {
  let k = 2;
  while (styleNameRefusal(wb, `${name} ${k}`)) k++;
  return `${name} ${k}`;
}

/**
 * Duplicate: style `newName` made a copy of style `name` (see the top of this file), with `includes` ({ number, ... }:
 * false for a part left out; null: the style's) and `change` (the Style box's Format..., as Modify takes it). Returns
 * { xfId } or { error }.
 */
export function duplicateCellStyle(wb, name, newName, includes = null, change = null) {
  const info = styleInfo(wb, name);
  if (!info) return { error: `There is no style named ${name}.` };
  const error = styleNameRefusal(wb, newName);
  if (error) return { error };
  const f = formatsOf(wb);
  const cs = findStyle(f.root, name);
  let xfId;
  let el;
  if (cs) {
    el = listOf(f.root, 'cellStyleXfs', 'xf')[Number(attr(cs, 'xfId') || 0)].cloneNode(true);
    xfId = appendStyleXf(f, el);
  } else ({ xfId, el } = builtinStyleXf(f, builtinDef(name)));
  const inc = includes || info.includes;
  setIncludes(el, inc);
  // (Each part the copy leaves out: Normal's, as on a cell of the style alone.)
  const normal = listOf(f.root, 'cellStyleXfs', 'xf')[0];
  for (const [key, , id] of STYLE_PARTS) {
    if (inc[key] !== false) continue;
    if (id) { el.setAttribute(id, attr(normal, id) || '0'); continue; }
    const tag = key === 'alignment' ? 'alignment' : 'protection';
    const mine = kid(el, tag, f.ns);
    if (mine) el.removeChild(mine);
    const theirs = kid(normal, tag, f.ns);
    if (theirs) el.insertBefore(theirs.cloneNode(true), tag === 'alignment' ? kid(el, 'protection', f.ns) || kid(el, 'extLst', f.ns) : kid(el, 'extLst', f.ns));
  }
  const copy = make(f, 'cellStyle');
  copy.setAttribute('name', String(newName).trim());
  copy.setAttribute('xfId', String(xfId));
  placeStyle(f, copy);
  reread(wb);
  if (change && Object.keys(change).length) {
    const out = modifyCellStyle(wb, String(newName).trim(), change, inc);
    if (out?.error) return out;
  }
  return { xfId };
}

/** An element of another styles part as one of this one's (its namespace and prefix), attributes but xmlns kept. */
function rebuild(f, node, ns) {
  if (node.namespaceURI !== ns) return f.doc.importNode(node, true);
  const el = make(f, node.localName);
  for (const a of [...node.attributes]) if (a.name !== 'xmlns' && !a.name.startsWith('xmlns:')) el.setAttribute(a.name, a.value);
  for (const c of [...node.childNodes]) if (c.nodeType === 1) el.appendChild(rebuild(f, c, ns));
  return el;
}

/** A style entry's look as text, to tell whether two workbooks' styles of one name differ: its number format's code,
 *  font, fill and border written out, its alignment, protection and apply flags. */
function lookOf(book, xf) {
  const root = book.stylesDoc.documentElement;
  const n = Number(attr(xf, 'numFmtId') || 0);
  const item = (list, tag, k) => listOf(root, list, tag)[Number(attr(xf, k) || 0)];
  const text = (el) => (el ? serializeXml(el).replace(/ xmlns(:\w+)?="[^"]*"/g, '').replace(/<(\/?)\w+:/g, '<$1') : '');
  return JSON.stringify([book.styles.numFmts[n] ?? n, text(item('fonts', 'font', 'fontId')), text(item('fills', 'fill', 'fillId')),
    text(item('borders', 'border', 'borderId')), text(kid(xf, 'alignment', root.namespaceURI)), text(kid(xf, 'protection', root.namespaceURI)),
    APPLY.map((p) => attr(xf, `apply${p}`) === '0')]);
}

/** The styles of workbook `from` that Merge Styles would take: [{ cs, xf, mine }] (mine: this workbook's cellStyle of
 *  the same name, or null). Hidden ones (deleted) are left. */
function mergeable(wb, from) {
  if (!from?.stylesDoc) return [];
  const src = from.stylesDoc.documentElement;
  const xfs = listOf(src, 'cellStyleXfs', 'xf');
  const root = wb.stylesDoc?.documentElement;
  return listOf(src, 'cellStyles', 'cellStyle').filter((cs) => attr(cs, 'hidden') !== '1' && attr(cs, 'iLevel') == null && xfs[Number(attr(cs, 'xfId') || 0)])
    .map((cs) => {
      const builtinId = attr(cs, 'builtinId');
      const mine = root ? findStyle(root, attr(cs, 'name'), builtinId != null ? Number(builtinId) : null) : null;
      // (A built-in style deleted here, kept hidden: the other workbook's put in its place and shown again.)
      const hidden = mine && attr(mine, 'hidden') === '1' ? mine : null;
      return { cs, xf: xfs[Number(attr(cs, 'xfId') || 0)], mine: hidden ? null : mine, hidden };
    });
}

/** The names of the styles in both workbooks that differ, for Excel's question "Merge styles that have the same
 *  names?" (none: no question). */
export function sameNameStyles(wb, from) {
  const xfs = wb.stylesDoc ? listOf(wb.stylesDoc.documentElement, 'cellStyleXfs', 'xf') : [];
  return mergeable(wb, from).filter(({ xf, mine }) => mine && lookOf(from, xf) !== lookOf(wb, xfs[Number(attr(mine, 'xfId') || 0)])).map(({ mine }) => attr(mine, 'name'));
}

/**
 * Merge Styles: the styles of workbook `from` (a Workbook) taken into this one (see the top of this file); `replace`:
 * the answer to "Merge styles that have the same names?" (false: this workbook's own kept). Returns { added, replaced }
 * (style names) and { formats } (the cell formats changed with a style replaced, by index).
 */
export function mergeCellStyles(wb, from, replace = true) {
  const out = { added: [], replaced: [], formats: [] };
  if (!from?.stylesDoc || from === wb) return out;
  const src = from.stylesDoc.documentElement;
  const sns = src.namespaceURI;
  const take = mergeable(wb, from).filter(({ mine }) => !mine || replace);
  if (!take.length) return out;
  const f = formatsOf(wb);
  const xfs = () => listOf(f.root, 'cellStyleXfs', 'xf');
  const entry = (list, tag, k, xf) => {
    const el = listOf(src, list, tag)[Number(attr(xf, k) || 0)];
    return el ? f.add(list, rebuild(f, el, sns)) : 0;
  };
  for (const { cs, xf, mine, hidden } of take) {
    if (mine && lookOf(from, xf) === lookOf(wb, xfs()[Number(attr(mine, 'xfId') || 0)])) continue;
    const el = make(f, 'xf');
    // (Its number format by its code: one of the workbook's own, or a new id from 164 up, as Excel numbers them.)
    const n = Number(attr(xf, 'numFmtId') || 0);
    const code = from.styles.numFmts[n];
    el.setAttribute('numFmtId', String(code != null ? f.numFmt(code) : n));
    el.setAttribute('fontId', String(entry('fonts', 'font', 'fontId', xf)));
    el.setAttribute('fillId', String(entry('fills', 'fill', 'fillId', xf)));
    el.setAttribute('borderId', String(entry('borders', 'border', 'borderId', xf)));
    for (const p of APPLY) if (attr(xf, `apply${p}`) === '0') el.setAttribute(`apply${p}`, '0');
    for (const tag of ['alignment', 'protection']) { const c = kid(xf, tag, sns); if (c) el.appendChild(rebuild(f, c, sns)); }
    if (mine) {
      out.formats.push(...putStyleXf(f, Number(attr(mine, 'xfId') || 0), el));
      if (attr(cs, 'customBuiltin') === '1' && attr(mine, 'builtinId') != null) mine.setAttribute('customBuiltin', '1');
      out.replaced.push(attr(mine, 'name'));
      continue;
    }
    if (hidden) {
      putStyleXf(f, Number(attr(hidden, 'xfId') || 0), el);
      hidden.removeAttribute('hidden');
      if (attr(cs, 'customBuiltin') === '1') hidden.setAttribute('customBuiltin', '1');
      out.added.push(attr(hidden, 'name'));
      continue;
    }
    const added = make(f, 'cellStyle');
    added.setAttribute('name', attr(cs, 'name'));
    added.setAttribute('xfId', String(appendStyleXf(f, el)));
    if (attr(cs, 'builtinId') != null) added.setAttribute('builtinId', attr(cs, 'builtinId'));
    if (attr(cs, 'customBuiltin') === '1') added.setAttribute('customBuiltin', '1');
    placeStyle(f, added);
    out.added.push(attr(cs, 'name'));
  }
  reread(wb);
  return out;
}
