// The Format Cells box's Number, Alignment and Font tabs (formatbox.js). Each is built on the active cell's look and
// marks in `changed` what was touched; its get() gives formatops.js formatCells's change for those only.
import { make, option, labelled, check, select, colorField, HORIZONTAL, VERTICAL, UNDERLINES } from './formatbox-kit.js';
import { FONT_SIZES } from '../ooxml/formatops.js';
import { cellDisplay } from '../ooxml/numfmt.js';
import { NUMBER_CATEGORIES, categoryOf, codeFor } from './formatcodes.js';

/** The font and size lists' ids. */
let lists = 0;

/** The Number tab. */
export function numberTab(look, changed) {
  const cat = categoryOf(look.code);
  const list = make('select', { className: 'oxe-dialog-list', size: 8 }, ...NUMBER_CATEGORIES.map((c) => option(c.name, c.name, c.name === cat.name)));
  const decimals = make('input', { type: 'number', min: 0, max: 30, value: String(cat.decimals ?? 2), className: 'oxe-fc-num' });
  const thousands = check('Use 1000 Separator (,)', cat.thousands);
  const types = make('select', { className: 'oxe-dialog-list', size: 6 });
  const custom = make('input', { type: 'text', className: 'oxe-dialog-field', value: look.code, spellcheck: false });
  const sample = make('output', { className: 'oxe-fc-sample' });
  const opts = make('div', { className: 'oxe-fc-opts' });
  const code = () => {
    const c = NUMBER_CATEGORIES.find((x) => x.name === list.value);
    return c.custom ? custom.value : codeFor(c, { decimals: Number(decimals.value) || 0, thousands: thousands.input.checked, type: types.value });
  };
  const show = () => {
    const c = NUMBER_CATEGORIES.find((x) => x.name === list.value);
    const kids = [];
    if (c.decimals !== undefined) kids.push(labelled('Decimal places:', decimals));
    if (c.thousands !== undefined) kids.push(thousands.el);
    if (c.types) {
      if (![...types.options].length || types.dataset.cat !== c.name) {
        types.replaceChildren(...c.types.map((t) => option(t, t)));
        types.dataset.cat = c.name;
        types.value = c.types.includes(look.code) ? look.code : c.types[0];
      }
      kids.push(labelled('Type:', types));
    }
    if (c.custom) kids.push(labelled('Type:', custom));
    opts.replaceChildren(...kids);
    const text = cellDisplay(look.cell, Infinity, code(), look.date1904).text;
    sample.textContent = `Sample: ${text}`;
  };
  for (const x of [list, decimals, thousands.input, types, custom]) x.addEventListener(x === custom || x === decimals ? 'input' : 'change', () => { changed.numFmt = true; show(); });
  show();
  return { el: make('div', { className: 'oxe-fc-cols' }, labelled('Category:', list), make('div', {}, sample, opts)), get: () => ({ numFmt: code() }) };
}

/** The Alignment tab. Excel's degrees run -90 to 90; stored as textRotation (90 + n down). */
export function alignmentTab(look, changed) {
  const a = look.xf.align;
  const h = select(HORIZONTAL, a.horizontal);
  const v = select(VERTICAL, a.vertical);
  const indent = make('input', { type: 'number', min: 0, max: 250, value: String(a.indent || 0), className: 'oxe-fc-num' });
  const deg = a.rotation === 255 ? 0 : a.rotation > 90 ? 90 - a.rotation : a.rotation;
  const degrees = make('input', { type: 'number', min: -90, max: 90, value: String(deg), className: 'oxe-fc-num' });
  const stacked = check('Vertical text', a.rotation === 255);
  const wrap = check('Wrap text', a.wrap);
  const shrink = check('Shrink to fit', a.shrink);
  const dir = select([['0', 'Context'], ['1', 'Left-to-Right'], ['2', 'Right-to-Left']], String(a.readingOrder || 0));
  // (As in Excel: Shrink to fit can't be chosen with Wrap text; an indent only with Left, Right or Distributed.)
  const sync = () => {
    shrink.input.disabled = wrap.input.checked;
    indent.disabled = !['left', 'right', 'distributed'].includes(h.value);
    degrees.disabled = stacked.input.checked;
  };
  const mark = (key) => () => { changed[key] = true; sync(); };
  h.addEventListener('change', mark('horizontal'));
  v.addEventListener('change', mark('vertical'));
  indent.addEventListener('input', mark('indent'));
  degrees.addEventListener('input', mark('rotation'));
  stacked.input.addEventListener('change', mark('rotation'));
  wrap.input.addEventListener('change', mark('wrap'));
  shrink.input.addEventListener('change', mark('shrink'));
  dir.addEventListener('change', mark('readingOrder'));
  sync();
  const get = () => {
    const align = {};
    if (changed.horizontal) align.horizontal = h.value;
    if (changed.vertical) align.vertical = v.value;
    if (changed.indent || changed.horizontal) align.indent = indent.disabled ? 0 : Math.max(0, Math.min(250, Math.trunc(Number(indent.value) || 0)));
    if (changed.rotation) {
      const n = Math.max(-90, Math.min(90, Math.trunc(Number(degrees.value) || 0)));
      align.rotation = stacked.input.checked ? 255 : n >= 0 ? n : 90 - n;
    }
    if (changed.wrap) align.wrap = wrap.input.checked;
    if (changed.shrink || (changed.wrap && wrap.input.checked)) align.shrink = shrink.input.checked && !wrap.input.checked;
    if (changed.readingOrder) align.readingOrder = Number(dir.value);
    return Object.keys(align).length ? { align } : {};
  };
  const el = make('div', { className: 'oxe-fc-grid' },
    make('fieldset', {}, make('legend', { textContent: 'Text alignment' }), labelled('Horizontal:', h), labelled('Vertical:', v), labelled('Indent:', indent)),
    make('fieldset', {}, make('legend', { textContent: 'Orientation' }), stacked.el, labelled('Degrees:', degrees)),
    make('fieldset', {}, make('legend', { textContent: 'Text control' }), wrap.el, shrink.el),
    make('fieldset', {}, make('legend', { textContent: 'Right-to-left' }), labelled('Text direction:', dir)));
  return { el, get };
}

/** The Font tab. */
export function fontTab(look, changed) {
  const f = look.font;
  const name = make('input', { type: 'text', className: 'oxe-dialog-field', value: look.fontName, spellcheck: false });
  const names = make('datalist', { id: `oxe-fc-fonts-${++lists}` }, ...look.fontNames.map((n) => option(n, n)));
  name.setAttribute('list', names.id);
  const style = select([['', 'Regular'], ['i', 'Italic'], ['b', 'Bold'], ['bi', 'Bold Italic']], `${f.b ? 'b' : ''}${f.i ? 'i' : ''}`);
  const size = make('input', { type: 'number', min: 1, max: 409, step: 0.5, value: String(f.size), className: 'oxe-fc-num' });
  const sizes = make('datalist', { id: `oxe-fc-sizes-${++lists}` }, ...FONT_SIZES.map((n) => option(String(n), String(n))));
  size.setAttribute('list', sizes.id);
  const u = select(UNDERLINES, f.u || '');
  const color = colorField(f.color && !f.color.auto && !(f.color.theme === 1 && !f.color.tint) ? f.color : null, 'Automatic', () => { changed.color = true; });
  const strike = check('Strikethrough', f.strike);
  const sup = check('Superscript', f.vertAlign === 'superscript');
  const sub = check('Subscript', f.vertAlign === 'subscript');
  const preview = make('div', { className: 'oxe-fc-preview', textContent: 'AaBbCcYyZz' });
  const paint = () => {
    preview.style.fontFamily = name.value;
    preview.style.fontWeight = style.value.includes('b') ? 'bold' : 'normal';
    preview.style.fontStyle = style.value.includes('i') ? 'italic' : 'normal';
    preview.style.fontSize = `${Math.min(36, Number(size.value) || 11)}pt`;
    preview.style.textDecoration = [u.value ? 'underline' : '', strike.input.checked ? 'line-through' : ''].filter(Boolean).join(' ') || 'none';
  };
  const mark = (key) => () => { changed[key] = true; paint(); };
  name.addEventListener('input', mark('name'));
  style.addEventListener('change', mark('style'));
  size.addEventListener('input', mark('size'));
  u.addEventListener('change', mark('u'));
  strike.input.addEventListener('change', mark('strike'));
  // (Superscript and subscript: one or the other.)
  sup.input.addEventListener('change', () => { if (sup.input.checked) sub.input.checked = false; changed.vertAlign = true; });
  sub.input.addEventListener('change', () => { if (sub.input.checked) sup.input.checked = false; changed.vertAlign = true; });
  paint();
  const get = () => {
    const font = {};
    if (changed.name && name.value.trim()) font.name = name.value.trim();
    if (changed.style) { font.b = style.value.includes('b'); font.i = style.value.includes('i'); }
    if (changed.size && Number(size.value) > 0) font.size = Math.min(409, Math.round(Number(size.value) * 2) / 2);
    if (changed.u) font.u = u.value || null;
    if (changed.color) font.color = color.value();
    if (changed.strike) font.strike = strike.input.checked;
    if (changed.vertAlign) font.vertAlign = sup.input.checked ? 'superscript' : sub.input.checked ? 'subscript' : null;
    return Object.keys(font).length ? { font } : {};
  };
  const el = make('div', { className: 'oxe-fc-grid' }, labelled('Font:', name), names, labelled('Font style:', style), labelled('Size:', size), sizes,
    labelled('Underline:', u), labelled('Color:', color.el),
    make('fieldset', {}, make('legend', { textContent: 'Effects' }), strike.el, sup.el, sub.el), preview);
  return { el, get };
}

