// Home > Cell Styles as Excel applies them (docs/FORMATS.md; measured over COM, tools/excel-compare/excel-cellstyles.ps1,
// test/cellstyles-excel.json): a built-in style the workbook hasn't got is added as Excel adds it (its font, fill and
// border written, a cellStyleXfs entry marking what it doesn't apply, a cellStyle in name order), and a cell given a
// style takes from it each part the style applies, keeping its own of the rest ("Good then Heading 1": Heading 1's font
// and border, Good's fill kept); the cell's xf marks applied each part that isn't the style's.
import { kids, kid, attr, serializeXml, parseXml, esc } from './xml.js';
import { readFont, readFill, readBorder, readXf } from './styles.js';
import { BUILTIN_CELL_STYLES } from '../engine/cellstyles.js';
import { BUILTIN_FORMATS } from './format.js';

/** The parts of an xf, as its id attributes and apply flags name them. */
const PARTS = [['numFmtId', 'NumberFormat'], ['fontId', 'Font'], ['fillId', 'Fill'], ['borderId', 'Border']];
const APPLY = ['NumberFormat', 'Font', 'Fill', 'Border', 'Alignment', 'Protection'];
/** Number formats whose codes Excel writes out (they depend on the locale): stylesedit.js's list. */
const LOCALE_BUILTINS = new Set([5, 6, 7, 8, 41, 42, 43, 44]);

const make = (f, name) => f.doc.createElementNS(f.ns, f.root.prefix ? `${f.root.prefix}:${name}` : name);
const parse = (f, xml) => f.doc.importNode(parseXml(`<w xmlns="${f.ns}">${xml}</w>`).documentElement.firstChild, true);

/** Append `el` to list `name` of the styles part (no matching: Excel writes a style's font anew); its index. */
function append(f, list, el, read) {
  const parent = f.list(list);
  parent.appendChild(el);
  const i = kids(parent, el.localName, f.ns).length - 1;
  parent.setAttribute('count', String(i + 1));
  f.seen[list]?.set(serializeXml(el), i);
  read?.(el, i);
  f.wb.dirty.add(f.wb.stylesPath);
  return i;
}

/** The font of built-in style `def` in this workbook: the theme's body font (as the Normal style's font names it) or
 *  heading font in place of MINOR and MAJOR, and the Normal style's size for a style of Excel's 11pt. */
function styleFont(f, def) {
  const normalEl = kids(kid(f.root, 'fonts', f.ns), 'font', f.ns)[0];
  const tail = normalEl ? kids(normalEl, null).filter((n) => ['name', 'family', 'charset', 'scheme'].includes(n.localName)).map((n) => serializeXml(n).replace(/ xmlns="[^"]*"/, '')).join('') : '<name val="Calibri"/><family val="2"/><scheme val="minor"/>';
  // (The theme's heading font is the file's text: escaped, so a name with quotes or tags stays one attribute.)
  const major = `<name val="${esc(f.wb.theme?.fonts?.major || 'Calibri Light')}"/><family val="2"/><scheme val="major"/>`;
  const size = f.wb.styles.normal?.size || 11;
  return def.font.replace('MINOR', tail).replace('MAJOR', major).replace('<sz val="11"/>', `<sz val="${esc(size)}"/>`);
}

/** The cellStyleXfs index of the cell style `name`: the workbook's own of that name, or a built-in one (found by its
 *  builtinId, else added as Excel adds it). Returns { xfId, el } (the style's xf), or null for a name there isn't. */
export function namedStyle(f, name) {
  const def = BUILTIN_CELL_STYLES.find((x) => x.name === name);
  const list = kids(kid(f.root, 'cellStyles', f.ns), 'cellStyle', f.ns);
  const found = list.find((el) => (def && attr(el, 'builtinId') === String(def.builtinId)) || attr(el, 'name') === name);
  const xfs = () => kids(kid(f.root, 'cellStyleXfs', f.ns), 'xf', f.ns);
  // (A built-in style deleted is kept hidden, and can't be given to cells again, as in Excel: cellstylemanage.js.)
  if (found && attr(found, 'hidden') === '1') return null;
  if (found) { const xfId = Number(attr(found, 'xfId') || 0); return { xfId, el: xfs()[xfId] }; }
  if (!def) return null;
  const { xfId, el: xf } = builtinStyleXf(f, def);
  // (Excel keeps the cellStyles in name order: Good before Normal, Normal before Title.)
  const cs = make(f, 'cellStyle');
  cs.setAttribute('name', name);
  cs.setAttribute('xfId', String(xfId));
  cs.setAttribute('builtinId', String(def.builtinId));
  const parent = f.list('cellStyles');
  const after = kids(parent, 'cellStyle', f.ns).find((el) => (attr(el, 'name') || '').toLowerCase() > name.toLowerCase());
  parent.insertBefore(cs, after || null);
  parent.setAttribute('count', String(kids(parent, 'cellStyle', f.ns).length));
  f.wb.dirty.add(f.wb.stylesPath);
  return { xfId, el: xf };
}

/** Built-in style `def`'s cellStyleXfs entry added as Excel adds it (its font written anew, its fill and border, what it
 *  doesn't apply marked applyX="0"), without its cellStyle: { xfId, el }. Duplicate's copy of a built-in style the
 *  workbook hasn't got is this entry (cellstylemanage.js). */
export function builtinStyleXf(f, def) {
  const styles = f.wb.styles;
  const fontId = append(f, 'fonts', parse(f, styleFont(f, def)), (x, i) => { styles.fonts[i] = readFont(x, f.ns); });
  const fillId = def.fill ? f.add('fills', parse(f, def.fill), (x, i) => { styles.fills[i] = readFill(x, f.ns); }) : 0;
  const borderId = def.border ? f.add('borders', parse(f, def.border), (x, i) => { styles.borders[i] = readBorder(x, f.ns); }) : 0;
  if (def.code && LOCALE_BUILTINS.has(def.numFmtId) && styles.numFmts[def.numFmtId] == null) f.writeNumFmt(def.numFmtId, def.code ?? BUILTIN_FORMATS[def.numFmtId]);
  const xf = make(f, 'xf');
  for (const [k, v] of [['numFmtId', def.numFmtId], ['fontId', fontId], ['fillId', fillId], ['borderId', borderId]]) xf.setAttribute(k, String(v));
  for (const p of APPLY) if (def.off.includes(p)) xf.setAttribute(`apply${p}`, '0');
  return { xfId: append(f, 'cellStyleXfs', xf), el: xf };
}

/** Cell format `s` given cell style `name` (Formats.xf's `cellStyle` change): the index of its new cell format. */
export function withCellStyle(f, s, name) {
  const style = namedStyle(f, name);
  if (!style) return s;
  const styles = f.wb.styles;
  const old = f.copy('cellXfs', 'xf', s);
  const applies = (p) => attr(style.el, `apply${p}`) !== '0';
  const num = (el, k) => Number(attr(el, k) || 0);
  const el = make(f, 'xf');
  for (const [k, p] of PARTS) el.setAttribute(k, String(applies(p) ? num(style.el, k) : num(old, k)));
  el.setAttribute('xfId', String(style.xfId));
  for (const [k, p] of PARTS) if (num(el, k) !== num(style.el, k)) el.setAttribute(`apply${p}`, '1');
  // (Alignment and protection: the style's when it applies them, else the cell's own, marked applied.)
  for (const [tag, p] of [['alignment', 'Alignment'], ['protection', 'Protection']]) {
    const from = applies(p) ? kid(style.el, tag, f.ns) : kid(old, tag, f.ns);
    if (from) el.appendChild(from.cloneNode(true));
    if (!applies(p) && from) el.setAttribute(`apply${p}`, '1');
  }
  return f.add('cellXfs', el, (x, i) => { styles.cellXfs[i] = readXf(x, f.ns, styles); });
}

/** The cell styles the workbook has of its own (not built-in): their names, for the gallery's Custom. */
export function customCellStyles(wb) {
  const doc = wb.stylesDoc;
  if (!doc) return [];
  const root = doc.documentElement;
  return kids(kid(root, 'cellStyles', root.namespaceURI), 'cellStyle', root.namespaceURI).filter((el) => attr(el, 'builtinId') == null && attr(el, 'customBuiltin') !== '1' && attr(el, 'hidden') !== '1').map((el) => attr(el, 'name'));
}

/** The built-in styles deleted from the workbook (kept hidden, as Excel keeps them): their builtinIds, for the gallery
 *  to leave out. */
export function deletedBuiltins(wb) {
  const doc = wb.stylesDoc;
  if (!doc) return new Set();
  const root = doc.documentElement;
  return new Set(kids(kid(root, 'cellStyles', root.namespaceURI), 'cellStyle', root.namespaceURI).filter((el) => attr(el, 'hidden') === '1' && attr(el, 'builtinId') != null).map((el) => Number(attr(el, 'builtinId'))));
}
