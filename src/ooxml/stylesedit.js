// Changing cell formats as Excel writes them (docs/FORMATS.md): a cell format (cellXfs) is the cell's old one with a
// change, and the font, fill or border it points to is the old one changed the same way. One that matches an entry
// already there is reused; a new one goes at the end, so indexes cells already use never move. Percent, comma and
// currency are Excel's built-in cell styles (Percent, Comma, Currency), added when the workbook hasn't got them.
import { kids, kid, attr, serializeXml, NS, REL_TYPES } from './xml.js';
import { readStyles, readFont, readFill, readBorder, readXf } from './styles.js';
import { BUILTIN_FORMATS } from './format.js';
import { withCellStyle } from './cellstyleedit.js';

/** Excel's order of a font's children, and of a border's sides. */
const FONT_ORDER = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name', 'family', 'charset', 'scheme'];
const SIDE_ORDER = ['start', 'left', 'end', 'right', 'top', 'bottom', 'diagonal', 'vertical', 'horizontal'];
/** Excel's order of an alignment's attributes (its saved styles.xml: tools/excel-compare/excel-formatcells.ps1). */
const ALIGN_ORDER = ['horizontal', 'vertical', 'textRotation', 'wrapText', 'indent', 'relativeIndent', 'justifyLastLine', 'shrinkToFit', 'readingOrder'];
/** The alignment change's keys and the attributes they set, each with the value Excel leaves out. */
const ALIGN_KEYS = { horizontal: ['horizontal', 'general'], vertical: ['vertical', 'bottom'], rotation: ['textRotation', 0], wrap: ['wrapText', false],
  indent: ['indent', 0], shrink: ['shrinkToFit', false], readingOrder: ['readingOrder', 0] };
/** Excel's built-in cell styles the Number group applies, and Hyperlink (a hyperlink's cell, links.js): builtinId,
 *  number format, and the font of its own (Hyperlink's: underlined, the theme's colour 10, measured). */
export const CELL_STYLES = { Comma: { builtinId: 3, numFmtId: 43 }, Currency: { builtinId: 4, numFmtId: 44 }, Percent: { builtinId: 5, numFmtId: 9 },
  Hyperlink: { builtinId: 8, numFmtId: 0, font: { u: 'single', color: { theme: 10 } } } };
// (The outline's automatic styles, outlinestyles.js: RowLevel_n and ColLevel_n, builtinId 1 and 2 told apart by iLevel;
// level 1 bold, level 2 italic, deeper the Normal font. Declared here, so they exist whatever module loads first.)
for (let n = 1; n <= 7; n++) {
  const font = n === 1 ? { b: true } : n === 2 ? { i: true } : {};
  CELL_STYLES[`RowLevel_${n}`] = { builtinId: 1, iLevel: n - 1, numFmtId: 0, font };
  CELL_STYLES[`ColLevel_${n}`] = { builtinId: 2, iLevel: n - 1, numFmtId: 0, font };
}
/** Built-in formats whose codes depend on the computer's locale: Excel writes their codes out (numFmts) when it uses
 *  them, so another locale reads the same thing. */
const LOCALE_BUILTINS = new Set([5, 6, 7, 8, 41, 42, 43, 44]);

/** A minimal styles part, as a workbook without one gets when it is first formatted: Excel's defaults (Calibri 11, no
 *  fill and the gray125 fill Excel always keeps second, no border, the Normal style). */
const EMPTY_STYLES = `<styleSheet xmlns="${NS.x}"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

const make = (doc, name) => {
  const root = doc.documentElement;
  return doc.createElementNS(root.namespaceURI, root.prefix ? `${root.prefix}:${name}` : name);
};

/** A colour element's attributes from { rgb: 'FFRRGGBB' } or { theme, tint } or { indexed }. */
function setColor(el, color) {
  for (const a of ['rgb', 'theme', 'tint', 'indexed', 'auto']) el.removeAttribute(a);
  if (color.rgb) el.setAttribute('rgb', color.rgb.length === 6 ? `FF${color.rgb}` : color.rgb);
  else if (color.theme != null) {
    el.setAttribute('theme', String(color.theme));
    if (color.tint) el.setAttribute('tint', color.tintText && Number(color.tintText) === color.tint ? color.tintText : String(color.tint));
  } else if (color.indexed != null) el.setAttribute('indexed', String(color.indexed));
  return el;
}

/** Put `child` among `parent`'s children in `order` (replacing one of the same name). */
function placeChild(parent, child, order) {
  const name = child.localName;
  const old = [...parent.childNodes].find((n) => n.nodeType === 1 && n.localName === name);
  if (old === child) return child;
  if (old) return parent.replaceChild(child, old);
  const after = order.indexOf(name);
  const next = [...parent.childNodes].find((n) => n.nodeType === 1 && order.indexOf(n.localName) > after);
  parent.insertBefore(child, next || null);
}
const removeChild = (parent, name) => {
  for (const n of [...parent.childNodes]) if (n.nodeType === 1 && n.localName === name) parent.removeChild(n);
};

/** The styles part of a workbook, made when it has none. */
export function stylesPart(wb) {
  if (wb.stylesDoc) return wb.stylesDoc;
  const path = 'xl/styles.xml';
  wb.pkg.setText(path, EMPTY_STYLES);
  wb.pkg.addOverride(path, 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml');
  wb.pkg.addRel(REL_TYPES.styles, 'styles.xml');
  wb.stylesPath = path;
  wb.stylesDoc = wb.pkg.xml(path);
  // (The model now reads the new part: the same defaults, with its second fill.)
  Object.assign(wb.styles, readStyles(wb.stylesDoc), { palette: wb.styles.palette });
  return wb.stylesDoc;
}

/**
 * Writes changed cell formats into a workbook's styles part and keeps `wb.styles` (styles.js) in step. Made once per
 * workbook (`formatsOf`).
 */
class Formats {
  constructor(wb) {
    this.wb = wb;
    this.doc = stylesPart(wb);
    this.root = this.doc.documentElement;
    this.ns = this.root.namespaceURI;
    // (Each list's entries as text, to find one that matches.)
    this.seen = {};
    for (const [list, item] of [['fonts', 'font'], ['fills', 'fill'], ['borders', 'border'], ['cellXfs', 'xf'], ['cellStyleXfs', 'xf']]) {
      this.seen[list] = new Map(this.items(list, item).map((el, i) => [serializeXml(el), i]).reverse());
    }
  }

  list(name) {
    let el = kid(this.root, name, this.ns);
    if (!el) {
      el = make(this.doc, name);
      const ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];
      placeChild(this.root, el, ORDER);
    }
    return el;
  }

  items(list, item) { return kids(kid(this.root, list, this.ns), item, this.ns); }

  /** The index of `el` in `list`: an equal entry's, or a new one at the end. */
  add(list, el, read) {
    const key = serializeXml(el);
    const seen = this.seen[list];
    if (seen.has(key)) return seen.get(key);
    const parent = this.list(list);
    parent.appendChild(el);
    const i = this.items(list, el.localName).length - 1;
    parent.setAttribute('count', String(i + 1));
    seen.set(key, i);
    if (read) read(el, i);
    this.wb.dirty.add(this.wb.stylesPath);
    return i;
  }

  /** A copy of entry i of a list (the first when there's none). */
  copy(list, item, i) {
    const el = this.items(list, item)[i] || this.items(list, item)[0];
    return el ? el.cloneNode(true) : make(this.doc, item);
  }

  /** A font changed by { b, i, strike, u ('single', 'double' or null), size, name, color ({ rgb } / { theme, tint },
   *  null: automatic) }; returns its index. */
  font(fontId, change) {
    const el = this.copy('fonts', 'font', fontId);
    const toggle = (name, on) => (on ? placeChild(el, make(this.doc, name), FONT_ORDER) : removeChild(el, name));
    if ('b' in change) toggle('b', change.b);
    if ('i' in change) toggle('i', change.i);
    if ('strike' in change) toggle('strike', change.strike);
    if ('u' in change) {
      removeChild(el, 'u');
      if (change.u) {
        const u = make(this.doc, 'u');
        if (change.u !== 'single') u.setAttribute('val', change.u);
        placeChild(el, u, FONT_ORDER);
      }
    }
    if (change.size != null) {
      const sz = make(this.doc, 'sz');
      sz.setAttribute('val', String(change.size));
      placeChild(el, sz, FONT_ORDER);
    }
    if (change.name != null) {
      const name = make(this.doc, 'name');
      name.setAttribute('val', change.name);
      placeChild(el, name, FONT_ORDER);
      // (A font chosen by name is no longer the theme's: Excel drops the scheme, and the family it doesn't know.)
      removeChild(el, 'scheme');
      removeChild(el, 'family');
      removeChild(el, 'charset');
    }
    if ('color' in change) {
      if (change.color) placeChild(el, setColor(make(this.doc, 'color'), change.color), FONT_ORDER);
      else placeChild(el, setColor(make(this.doc, 'color'), { theme: 1 }), FONT_ORDER);
    }
    if ('vertAlign' in change) {
      removeChild(el, 'vertAlign');
      if (change.vertAlign) {
        const v = make(this.doc, 'vertAlign');
        v.setAttribute('val', change.vertAlign);
        placeChild(el, v, FONT_ORDER);
      }
    }
    return this.add('fonts', el, (x, i) => { this.wb.styles.fonts[i] = readFont(x, this.ns); });
  }

  /** A fill: a colour (solid), { pattern, fg, bg } (a hatch or shade: its pattern colour and the cell's, either left
   *  out for automatic), { gradient: { type, degree, left, right, top, bottom, stops: [{ position, color }] } }, or
   *  none (null: fill 0, which every workbook has as "none"). Written as Excel writes each (excel-formatcells.ps1). */
  fill(change) {
    if (!change) return 0;
    const el = make(this.doc, 'fill');
    if (change.gradient) {
      const g = el.appendChild(make(this.doc, 'gradientFill'));
      const { type, degree, stops } = change.gradient;
      if (type === 'path') {
        g.setAttribute('type', 'path');
        for (const k of ['left', 'right', 'top', 'bottom']) if (change.gradient[k]) g.setAttribute(k, String(change.gradient[k]));
      } else if (degree) g.setAttribute('degree', String(degree));
      for (const s of stops) {
        const stop = g.appendChild(make(this.doc, 'stop'));
        stop.setAttribute('position', String(s.position));
        stop.appendChild(setColor(make(this.doc, 'color'), s.color));
      }
    } else {
      const p = el.appendChild(make(this.doc, 'patternFill'));
      const pattern = change.pattern || 'solid';
      p.setAttribute('patternType', pattern);
      // (Solid: the colour, with indexed 64 behind it. A pattern: the colours chosen.)
      const fg = pattern === 'solid' ? change.fg || change : change.fg;
      const bg = pattern === 'solid' ? { indexed: 64 } : change.bg;
      if (fg) p.appendChild(setColor(make(this.doc, 'fgColor'), fg));
      if (bg) p.appendChild(setColor(make(this.doc, 'bgColor'), bg));
    }
    return this.add('fills', el, (x, i) => { this.wb.styles.fills[i] = readFill(x, this.ns); });
  }

  /** A border with some sides changed: { left, right, top, bottom, diagonal }, each { style, color } or null (no
   *  line), and which diagonals are drawn (diagonalUp, diagonalDown: true or false; left out, unchanged). */
  border(borderId, change) {
    const el = this.copy('borders', 'border', borderId);
    const { diagonalUp, diagonalDown, ...sides } = change;
    // (Excel writes diagonalUp before diagonalDown, and neither when off.)
    const up = diagonalUp ?? el.getAttribute('diagonalUp') === '1';
    const down = diagonalDown ?? el.getAttribute('diagonalDown') === '1';
    el.removeAttribute('diagonalUp');
    el.removeAttribute('diagonalDown');
    if (up) el.setAttribute('diagonalUp', '1');
    if (down) el.setAttribute('diagonalDown', '1');
    for (const [name, side] of Object.entries(sides)) {
      // (A right-to-left file may name the sides start and end.)
      if (name === 'left') removeChild(el, 'start');
      if (name === 'right') removeChild(el, 'end');
      const s = make(this.doc, name);
      if (side) {
        s.setAttribute('style', side.style);
        s.appendChild(setColor(make(this.doc, 'color'), side.color || { indexed: 64 }));
      }
      placeChild(el, s, SIDE_ORDER);
    }
    for (const name of ['left', 'right', 'top', 'bottom', 'diagonal']) if (!kid(el, name, this.ns) && !(name === 'left' && kid(el, 'start', this.ns)) && !(name === 'right' && kid(el, 'end', this.ns))) placeChild(el, make(this.doc, name), SIDE_ORDER);
    return this.add('borders', el, (x, i) => { this.wb.styles.borders[i] = readBorder(x, this.ns); });
  }

  /** The id of a number format code: a built-in one's, the workbook's own, or a new one (164 up, as Excel numbers
   *  them). */
  numFmt(code) {
    const styles = this.wb.styles;
    for (const [id, c] of Object.entries(styles.numFmts)) if (c === code) return Number(id);
    for (const [id, c] of Object.entries(BUILTIN_FORMATS)) if (c === code && !LOCALE_BUILTINS.has(Number(id))) return Number(id);
    const builtin = Object.entries(BUILTIN_FORMATS).find(([, c]) => c === code);
    const id = builtin ? Number(builtin[0]) : Math.max(163, ...Object.keys(styles.numFmts).map(Number)) + 1;
    this.writeNumFmt(id, code);
    return id;
  }

  writeNumFmt(id, code) {
    const el = make(this.doc, 'numFmt');
    el.setAttribute('numFmtId', String(id));
    el.setAttribute('formatCode', code);
    const parent = this.list('numFmts');
    parent.appendChild(el);
    parent.setAttribute('count', String(kids(parent, 'numFmt', this.ns).length));
    this.wb.styles.numFmts[id] = code;
    this.wb.dirty.add(this.wb.stylesPath);
  }

  /** The cellStyleXfs index of one of Excel's built-in cell styles (CELL_STYLES), added with its cellStyle entry when
   *  the workbook hasn't got it. */
  cellStyle(name) {
    const { builtinId, numFmtId, font, iLevel } = CELL_STYLES[name];
    // (The outline's RowLevel_n and ColLevel_n share a builtinId, told apart by iLevel: outlinestyles.js.)
    const found = kids(kid(this.root, 'cellStyles', this.ns), 'cellStyle', this.ns).find((el) => attr(el, 'builtinId') === String(builtinId) && (iLevel == null || attr(el, 'iLevel') === String(iLevel)));
    if (found) return Number(attr(found, 'xfId') || 0);
    if (LOCALE_BUILTINS.has(numFmtId) && this.wb.styles.numFmts[numFmtId] == null) this.writeNumFmt(numFmtId, BUILTIN_FORMATS[numFmtId]);
    // (As Excel writes a style's own format: the Normal font, no fill or border, and only its number format applied;
    // a style with a font of its own, that font, and only it applied.)
    const xf = make(this.doc, 'xf');
    const fontId = font ? this.font(0, font) : 0;
    const applied = font ? [['applyNumberFormat', 0]] : [['applyFont', 0]];
    for (const [k, v] of [['numFmtId', numFmtId], ['fontId', fontId], ['fillId', 0], ['borderId', 0], ...applied, ['applyFill', 0], ['applyBorder', 0], ['applyAlignment', 0], ['applyProtection', 0]]) xf.setAttribute(k, String(v));
    const parent = this.list('cellStyleXfs');
    parent.appendChild(xf);
    const xfId = kids(parent, 'xf', this.ns).length - 1;
    parent.setAttribute('count', String(xfId + 1));
    const styles = this.list('cellStyles');
    const cs = make(this.doc, 'cellStyle');
    cs.setAttribute('name', name);
    cs.setAttribute('xfId', String(xfId));
    cs.setAttribute('builtinId', String(builtinId));
    if (iLevel != null) cs.setAttribute('iLevel', String(iLevel));
    styles.appendChild(cs);
    styles.setAttribute('count', String(kids(styles, 'cellStyle', this.ns).length));
    this.wb.dirty.add(this.wb.stylesPath);
    return xfId;
  }

  /**
   * Cell format `s` with a change: { font: {...} (see font()), fill (see fill()), border: { side: ... } (see
   * border()), align: { horizontal, vertical, wrap, indent, rotation, shrink, readingOrder }, protection: { locked,
   * hidden }, numFmt: code, style: 'Percent' | 'Comma' | 'Currency' } or { cellStyle: a cell style's name }.
   * Returns the index of the cell format.
   */
  xf(s, change) {
    // (A cell style of Home > Cell Styles: cellstyleedit.js, as Excel applies one.)
    if (change.cellStyle) return withCellStyle(this, s, change.cellStyle);
    const styles = this.wb.styles;
    const old = styles.cellXfs[s] || styles.cellXfs[0];
    const el = this.copy('cellXfs', 'xf', s);
    const set = (k, v, apply) => { el.setAttribute(k, String(v)); if (apply) el.setAttribute(apply, '1'); };
    if (change.font) set('fontId', this.font(old.fontId, change.font), 'applyFont');
    if ('fill' in change) set('fillId', this.fill(change.fill), 'applyFill');
    if (change.border) set('borderId', this.border(old.borderId, change.border), 'applyBorder');
    if (change.style && CELL_STYLES[change.style].font) {
      // (Hyperlink: the style's own font given to the cell, nothing marked applied, as Excel writes it.)
      const xfId = this.cellStyle(change.style);
      set('xfId', xfId);
      set('fontId', Number(attr(kids(this.list('cellStyleXfs'), 'xf', this.ns)[xfId], 'fontId') || 0));
    } else if (change.style) {
      set('xfId', this.cellStyle(change.style));
      set('numFmtId', CELL_STYLES[change.style].numFmtId, 'applyNumberFormat');
    } else if (change.numFmt != null) set('numFmtId', this.numFmt(change.numFmt), 'applyNumberFormat');
    if (change.align) {
      const a = kid(el, 'alignment', this.ns) || make(this.doc, 'alignment');
      const values = Object.fromEntries([...a.attributes].map((x) => [x.name, x.value]));
      for (const [key, [name, dflt]] of Object.entries(ALIGN_KEYS)) {
        if (change.align[key] === undefined) continue;
        const v = change.align[key];
        if (v == null || v === dflt || v === false) delete values[name];
        else values[name] = v === true ? '1' : String(v);
      }
      // (An indented or Fill-aligned cell can't be rotated: Excel drops the rotation, whichever came first.)
      if (Number(values.indent || 0) > 0 || values.horizontal === 'fill') delete values.textRotation;
      for (const x of [...a.attributes]) a.removeAttribute(x.name);
      const rank = (n) => (ALIGN_ORDER.includes(n) ? ALIGN_ORDER.indexOf(n) : 99);
      for (const name of Object.keys(values).sort((p, q) => rank(p) - rank(q))) a.setAttribute(name, values[name]);
      if (a.attributes.length) placeChild(el, a, ['alignment', 'protection', 'extLst']);
      else removeChild(el, 'alignment');
      el.setAttribute('applyAlignment', '1');
    }
    if (change.protection) {
      const p = kid(el, 'protection', this.ns) || make(this.doc, 'protection');
      const locked = change.protection.locked ?? p.getAttribute('locked') !== '0';
      const hidden = change.protection.hidden ?? p.getAttribute('hidden') === '1';
      p.removeAttribute('locked');
      p.removeAttribute('hidden');
      if (!locked) p.setAttribute('locked', '0');
      if (hidden) p.setAttribute('hidden', '1');
      if (p.attributes.length) placeChild(el, p, ['alignment', 'protection', 'extLst']);
      else removeChild(el, 'protection');
      el.setAttribute('applyProtection', '1');
    }
    if (!el.getAttribute('xfId')) el.setAttribute('xfId', '0');
    return this.add('cellXfs', el, (x, i) => { styles.cellXfs[i] = readXf(x, this.ns, styles); });
  }
}

const formats = new WeakMap();
/** The format writer of a workbook. */
export function formatsOf(wb) {
  if (!formats.has(wb)) formats.set(wb, new Formats(wb));
  return formats.get(wb);
}

