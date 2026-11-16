// The table style gallery (Home > Format as Table, Table Design > Table Styles): the workbook's own styles under Custom,
// then Excel's 60 built-in ones under Light, Medium and Dark, each shown as a small table drawn in it with the Table
// Style Options given (ooxml/tablestyle.js works out each cell's fill and lines), named as Excel names them.
import { el } from './ribbonkit.js';
import { tableLook, tableStyleElements, BUILTIN_TABLE_STYLES } from '../ooxml/tablestyle.js';

const ROWS = 5;
const COLS = 5;

/** "Table Style Medium 2" for TableStyleMedium2; a workbook's own style by its own name. */
export const styleTitle = (name) => {
  const m = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(name);
  return m ? `Table Style ${m[1]} ${m[2]}` : name;
};

/** One swatch: a ROWS by COLS table in style `name`, with `options` ({ headers, totals, first, last, rows, cols }). */
function swatch(wb, name, options, onPick, run) {
  const t = { r0: 0, c0: 0, r1: ROWS - 1, c1: COLS - 1, headers: options.headers === false ? 0 : 1, totals: options.totals ? 1 : 0, style: { first: !!options.first, last: !!options.last, rows: options.rows !== false, cols: !!options.cols } };
  const look = tableLook(t, tableStyleElements(wb.styles, name), wb.theme, wb.styles.palette);
  const cells = [];
  for (let r = 0; r < ROWS; r++) {
    for (let c = 0; c < COLS; c++) {
      const l = look.get(`${r},${c}`) || {};
      const line = (side) => (l.border?.[side] ? `${l.border[side].style === 'thick' || l.border[side].style === 'double' || l.border[side].style === 'medium' ? 2 : 1}px solid ${l.border[side].color}` : '0');
      // (A colour of the text: a bar in the cell, as Excel's gallery shows it.)
      cells.push(el('span', { class: 'oxe-tstyle-cell', style: `background:${l.fill || '#ffffff'};border-top:${line('top')};border-bottom:${line('bottom')};border-left:${line('left')};border-right:${line('right')};--oxe-tstyle-ink:${l.color || '#000000'}` }));
    }
  }
  const title = styleTitle(name);
  return el('button', { type: 'button', class: 'oxe-tstyle', title, 'aria-label': title, onmousedown: run(() => onPick(name)) }, ...cells);
}

/** The gallery's menu items: Custom (the workbook's own), Light, Medium, Dark; `withNone` adds Excel's Clear. */
export function tableGallery(kit, { wb, options = {}, onPick, withNone = false }) {
  const { run, item } = kit;
  if (!wb) return [];
  const own = Object.keys(wb.styles.tableStyles || {});
  const section = (label, names) => (names.length ? [el('div', { class: 'oxe-cap' }, label), el('div', { class: 'oxe-tstyles' }, ...names.map((n) => swatch(wb, n, options, onPick, run)))] : []);
  return [
    ...section('Custom', own),
    ...section('Light', BUILTIN_TABLE_STYLES.filter((n) => n.includes('Light'))),
    ...section('Medium', BUILTIN_TABLE_STYLES.filter((n) => n.includes('Medium'))),
    ...section('Dark', BUILTIN_TABLE_STYLES.filter((n) => n.includes('Dark'))),
    ...(withNone ? [el('hr'), item('Clear', () => onPick(null))] : []),
  ];
}
