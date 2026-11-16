// The cell style gallery (Home > Cell Styles): the workbook's own styles under Custom, then Excel's built-in ones under
// its gallery's headings (Good, Bad and Neutral; Data and Model; Titles and Headings; Themed Cell Styles; Number
// Format), each named and shown in its fill, font colour, weight and bottom line as the workbook's theme gives them.
// Picked, the selection takes it (ooxml/cellstyleedit.js, as Excel applies one); its menu (a right click, or the menu
// key) offers Modify..., and New Cell Style... ends the gallery, as in Excel (acts-styles.js).
import { el } from './ribbonkit.js';
import { parseXml, NS } from '../ooxml/xml.js';
import { readFont, readFill, readBorder } from '../ooxml/styles.js';
import { resolveColor } from '../ooxml/theme.js';
import { BUILTIN_CELL_STYLES } from '../engine/cellstyles.js';
import { customCellStyles } from '../ooxml/cellstyleedit.js';

/** Excel's gallery, heading by heading. */
export const CELL_STYLE_GALLERY = [
  ['Good, Bad and Neutral', ['Normal', 'Bad', 'Good', 'Neutral']],
  ['Data and Model', ['Calculation', 'Check Cell', 'Explanatory Text', 'Input', 'Linked Cell', 'Note', 'Output', 'Warning Text']],
  ['Titles and Headings', ['Heading 1', 'Heading 2', 'Heading 3', 'Heading 4', 'Title', 'Total']],
  ['Themed Cell Styles', ['20% - ', '40% - ', '60% - ', ''].flatMap((p) => [1, 2, 3, 4, 5, 6].map((k) => `${p}Accent${k}`))],
  ['Number Format', ['Comma', 'Comma [0]', 'Currency', 'Currency [0]', 'Percent']],
];

const read = (xml, fn) => (xml ? fn(parseXml(`<w xmlns="${NS.x}">${xml}</w>`).documentElement.firstChild, NS.x) : null);

/** A built-in style's look as CSS: background, colour, weight, italic and bottom line. */
function lookCss(name, wb) {
  const def = BUILTIN_CELL_STYLES.find((x) => x.name === name);
  if (!def) return '';
  const color = (c) => resolveColor(c, wb.theme, wb.styles.palette);
  const font = read(def.font.replace(/MINOR|MAJOR/, ''), readFont);
  const fill = read(def.fill, readFill);
  const border = read(def.border, readBorder);
  const out = [];
  if (fill?.fg && fill.pattern !== 'none') out.push(`background:${color(fill.fg) || '#ffffff'}`);
  if (font?.color) out.push(`color:${color(font.color) || '#000000'}`);
  if (font?.b) out.push('font-weight:700');
  if (font?.i) out.push('font-style:italic');
  if (border?.bottom?.style) out.push(`border-bottom:${border.bottom.style === 'thick' ? 3 : border.bottom.style === 'medium' ? 2 : 1}px ${border.bottom.style === 'double' ? 'double' : 'solid'} ${color(border.bottom.color) || '#000000'}`);
  if (border?.left?.style) out.push(`box-shadow:inset 0 0 0 1px ${color(border.left.color) || '#7f7f7f'}`);
  return out.join(';');
}

/** The gallery's menu items; picking one calls `onPick(name)`, its Modify... `onModify(name)`, New Cell Style...
 *  `onNew()`. */
export function cellStyleGallery(kit, { wb, onPick, onModify = null, onNew = null }) {
  const { run, item, openMenu } = kit;
  if (!wb) return [];
  const chip = (name) => {
    const b = el('button', { type: 'button', class: 'oxe-cstyle', title: name, style: lookCss(name, wb), 'aria-haspopup': onModify ? 'menu' : null, onmousedown: (e) => { if (e.button !== 2) run(() => onPick(name))(e); } }, name);
    // (A right click opens its menu instead of applying it, as in Excel.)
    if (onModify) b.addEventListener('contextmenu', (e) => { e.preventDefault(); openMenu(b, () => [item('Apply', () => onPick(name)), item('Modify...', () => onModify(name))]); });
    return b;
  };
  const own = customCellStyles(wb);
  return [
    ...(own.length ? [el('div', { class: 'oxe-cap' }, 'Custom'), el('div', { class: 'oxe-cstyles' }, ...own.map(chip))] : []),
    ...CELL_STYLE_GALLERY.flatMap(([label, names]) => [el('div', { class: 'oxe-cap' }, label), el('div', { class: 'oxe-cstyles' }, ...names.map(chip))]),
    ...(onNew ? [el('hr'), item('New Cell Style...', onNew)] : []),
  ];
}
