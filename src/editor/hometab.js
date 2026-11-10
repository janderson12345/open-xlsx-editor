// The Home tab, Excel's groups in Excel's order: Undo, Font, Alignment, Number, Cells, Editing. Every control runs a
// named command (commands.js, through `act`); update() sets the controls from the active cell, as Excel's do.
import { el } from './ribbonkit.js';
import { STANDARD_COLORS, THEME_COLORS } from './sheetmenu.js';
import { FONT_SIZES, BORDER_PRESETS } from '../ooxml/formatops.js';

const FONTS = ['Aptos', 'Aptos Narrow', 'Arial', 'Calibri', 'Cambria', 'Consolas', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Verdana'];
/** The Number Format list: its names and the codes Excel gives (en-US). */
export const NUMBER_FORMATS = [
  ['General', 'General'], ['Number', '0.00'], ['Currency', '"$"#,##0.00'],
  ['Accounting', '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'], ['Short Date', 'm/d/yyyy'],
  ['Long Date', '[$-x-sysdate]dddd, mmmm dd, yyyy'], ['Time', '[$-x-systime]h:mm:ss AM/PM'], ['Percentage', '0.00%'],
  ['Fraction', '# ?/?'], ['Scientific', '0.00E+00'], ['Text', '@'],
];
/** The Borders menu, as Excel lists it. */
const BORDERS = [['bottom', 'Bottom Border'], ['top', 'Top Border'], ['left', 'Left Border'], ['right', 'Right Border'], null,
  ['none', 'No Border'], ['all', 'All Borders'], ['outside', 'Outside Borders'], ['thickOutside', 'Thick Outside Borders'], null,
  ['doubleBottom', 'Bottom Double Border'], ['thickBottom', 'Thick Bottom Border'], ['topBottom', 'Top and Bottom Border'],
  ['topThickBottom', 'Top and Thick Bottom Border'], ['topDoubleBottom', 'Top and Double Bottom Border']];
/** The tints of each theme colour's column in Excel's palette (rows 2 to 6), as Excel writes them. */
const LIGHTER = [0.7999816888943144, 0.5999938962981048, 0.3999755851924192, -0.249977111117893, -0.499984740745262];
const TINTS = {
  0: [-0.0499893185216834, -0.1499984740745262, -0.249977111117893, -0.3499862666707358, -0.499984740745262],
  1: [0.499984740745262, 0.3499862666707358, 0.249977111117893, 0.1499984740745262, 0.0499893185216834],
  2: [-0.0999786370433668, -0.249977111117893, -0.499984740745262, -0.749992370372631, -0.8999908444471572],
  3: [0.8999908444471572, 0.749992370372631, 0.499984740745262, 0.249977111117893, 0.0999786370433668],
};
const pct = (t) => `${t > 0 ? 'Lighter' : 'Darker'} ${Math.round(Math.abs(t) * 100)}%`;

/** The Home tab. `act(name, ...args)` runs a command; `color(c)` resolves a colour for a swatch. */
export function homeTab(kit, { act, color }) {
  const { btn, toggle, grp, col, line, item, menu, split, pick, closeMenu, run } = kit;
  const f = (name, ...args) => () => act(name, ...args);

  // A colour palette: the theme's ten in six rows, Excel's ten standard ones, and `none` (Automatic or No Fill).
  const palette = (onPick, none) => () => {
    const sw = (c, title) => el('button', { type: 'button', class: 'oxe-swatch', title, 'aria-label': title, style: `background:${color(c) || '#000'}`, onmousedown: run(() => onPick(c)) });
    const theme = el('div', { class: 'oxe-swatches oxe-swatches-theme' });
    for (let row = 0; row < 6; row++) {
      for (let t = 0; t < 10; t++) {
        const tint = row ? (TINTS[t] || LIGHTER)[row - 1] : 0;
        theme.append(sw(tint ? { theme: t, tint } : { theme: t }, `${THEME_COLORS[t]}${tint ? `, ${pct(tint)}` : ''}`));
      }
    }
    const std = el('div', { class: 'oxe-swatches' }, ...STANDARD_COLORS.map(([n, hex]) => sw({ rgb: `FF${hex}` }, n)));
    return [item(none, () => onPick(null)), el('div', { class: 'oxe-cap' }, 'Theme Colors'), theme, el('div', { class: 'oxe-cap' }, 'Standard Colors'), std];
  };
  // (The buttons apply the last colour chosen: red and yellow to start with, as in Excel.)
  let fontColor = { rgb: 'FFFF0000' };
  let fillColor = { rgb: 'FFFFFF00' };
  let border = 'bottom';
  const fontColorBtn = split('A', 'Font Color', () => act('format', { font: { color: fontColor } }),
    palette((c) => { fontColor = c; paintBars(); act('format', { font: { color: c } }); }, 'Automatic'), { cls: 'oxe-colorbtn oxe-fontcolor' });
  const fillBtn = split('🪣', 'Fill Color', () => act('format', { fill: fillColor }),
    palette((c) => { fillColor = c; paintBars(); act('format', { fill: c }); }, 'No Fill'), { cls: 'oxe-colorbtn oxe-fillcolor' });
  const paintBars = () => {
    fontColorBtn.main.style.setProperty('--oxe-swatch', fontColor ? color(fontColor) : '#000000');
    fillBtn.main.style.setProperty('--oxe-swatch', fillColor ? color(fillColor) : 'transparent');
  };
  const bordersBtn = split('⊞', 'Bottom Border', () => act('borders', border),
    () => BORDERS.map((b) => (b ? item(b[1], () => { border = b[0]; bordersBtn.main.title = b[1]; bordersBtn.main.setAttribute('aria-label', b[1]); act('borders', b[0]); }) : el('hr'))));
  if (BORDERS.filter(Boolean).some(([k]) => !BORDER_PRESETS.includes(k))) throw new Error('a border preset without a rule');

  const fontSel = pick('Font', 'oxe-font', [], (v) => act('format', { font: { name: v } }));
  const sizeSel = pick('Font Size', 'oxe-size', [], (v) => act('format', { font: { size: Number(v) } }));
  const numSel = pick('Number Format', 'oxe-numfmt', NUMBER_FORMATS.map(([n, code]) => el('option', { value: code }, n)), (v) => { if (v) act('format', { numFmt: v }); });
  const b = {
    bold: toggle('B', 'Bold (Ctrl+B)', f('toggleFont', 'b'), { cls: 'oxe-b' }),
    italic: toggle('I', 'Italic (Ctrl+I)', f('toggleFont', 'i'), { cls: 'oxe-i' }),
    underline: toggle('U', 'Underline (Ctrl+U)', f('toggleFont', 'u'), { cls: 'oxe-u' }),
    strike: toggle('S', 'Strikethrough (Ctrl+5)', f('toggleFont', 'strike'), { cls: 'oxe-s' }),
    top: toggle('⤒', 'Top Align', f('align', 'vertical', 'top')),
    middle: toggle('↕', 'Middle Align', f('align', 'vertical', 'center')),
    bottom: toggle('⤓', 'Bottom Align', f('align', 'vertical', 'bottom')),
    wrap: toggle('⏎', 'Wrap Text', f('toggleWrap'), { label: 'Wrap Text' }),
    left: toggle('⫷', 'Align Left', f('align', 'horizontal', 'left')),
    center: toggle('≡', 'Center', f('align', 'horizontal', 'center')),
    right: toggle('⫸', 'Align Right', f('align', 'horizontal', 'right')),
    undo: btn('↶', 'Undo (Ctrl+Z)', f('undo')),
    redo: btn('↷', 'Redo (Ctrl+Y)', f('redo')),
  };
  const sheetMenu = (icon, label, title, items) => menu(icon, title, items, { label, big: false });
  const groups = [
    grp('Undo', col(b.undo, b.redo)),
    grp('Clipboard', split('📋', 'Paste (Ctrl+V)', f('clipPaste'), () => [item('Paste', f('clipPaste'), { hint: 'Ctrl+V' }), item('Formulas', f('pasteAs', 'formulas')),
      item('Values', f('pasteAs', 'values')), item('Formatting', f('pasteAs', 'formats')), item('Transpose', f('pasteAs', 'transpose')), item('Paste Link', f('pasteAs', 'link')),
      el('hr'), item('Paste Special...', f('pasteSpecialBox'), { hint: 'Ctrl+Alt+V' })], { label: 'Paste', big: true }),
      col(btn('✂', 'Cut (Ctrl+X)', f('clipCut'), { label: 'Cut' }), btn('⧉', 'Copy (Ctrl+C)', f('clipCopy'), { label: 'Copy' }))),
    grp('Font', col(
      line(fontSel, sizeSel, btn('A⁺', 'Increase Font Size', f('stepFont', 1)), btn('A⁻', 'Decrease Font Size', f('stepFont', -1))),
      line(b.bold, b.italic, b.underline, b.strike, bordersBtn, fillBtn, fontColorBtn))),
    grp('Alignment', col(
      line(b.top, b.middle, b.bottom, b.wrap),
      line(b.left, b.center, b.right, btn('⇤', 'Decrease Indent', f('stepIndent', -1)), btn('⇥', 'Increase Indent', f('stepIndent', 1)),
        split('⊟', 'Merge & Center', f('merge', 'center'), () => [item('Merge & Center', f('merge', 'center')), item('Merge Across', f('merge', 'across')),
          item('Merge Cells', f('merge', 'cells')), item('Unmerge Cells', f('merge', 'unmerge'))], { label: 'Merge & Center' })))),
    grp('Number', col(
      line(numSel),
      line(btn('$', 'Accounting Number Format', f('format', { style: 'Currency' })), btn('%', 'Percent Style (Ctrl+Shift+%)', f('format', { style: 'Percent' })),
        btn(',', 'Comma Style', f('format', { style: 'Comma' })), btn('.0←', 'Increase Decimal', f('stepDecimal', 1)), btn('.0→', 'Decrease Decimal', f('stepDecimal', -1))))),
    grp('Cells', col(
      sheetMenu('⊕', 'Insert', 'Insert', () => [item('Insert Cells...', f('insertCells'), { hint: 'Ctrl++' }), item('Insert Sheet Rows', f('insertRows')), item('Insert Sheet Columns', f('insertColumns')),
        item('Insert Sheet', f('insertSheet'), { hint: 'Shift+F11' })]),
      sheetMenu('🗑', 'Delete', 'Delete', () => [item('Delete Cells...', f('deleteCells'), { hint: 'Ctrl+-' }), item('Delete Sheet Rows', f('deleteRows')), item('Delete Sheet Columns', f('deleteColumns')),
        item('Delete Sheet', f('deleteShownSheet'))]),
      sheetMenu('✎', 'Format', 'Format', () => [el('div', { class: 'oxe-cap' }, 'Cell Size'), item('Row Height...', f('rowHeightBox')), item('AutoFit Row Height', f('fitLines', 'row')),
        item('Column Width...', f('columnWidthBox')), item('AutoFit Column Width', f('fitLines', 'col')), el('hr'),
        el('div', { class: 'oxe-cap' }, 'Visibility'), item('Hide Rows', f('hideLines', 'row', true), { hint: 'Ctrl+9' }), item('Hide Columns', f('hideLines', 'col', true), { hint: 'Ctrl+0' }),
        item('Unhide Rows', f('hideLines', 'row', false), { hint: 'Ctrl+Shift+9' }), item('Unhide Columns', f('hideLines', 'col', false)),
        item('Hide Sheet', f('hideShownSheet')), item('Unhide Sheet...', f('unhideSheet')), el('hr'),
        el('div', { class: 'oxe-cap' }, 'Organize Sheets'), item('Rename Sheet', f('renameShownSheet'))]))),
    grp('Editing', col(
      split('∑', 'AutoSum (Alt+=)', f('autoSum', 'sum'), () => [item('Sum', f('autoSum', 'sum'), { hint: 'Alt+=' }), item('Average', f('autoSum', 'average')),
        item('Count Numbers', f('autoSum', 'count')), item('Max', f('autoSum', 'max')), item('Min', f('autoSum', 'min'))], { label: 'AutoSum' }),
      sheetMenu('⬇', 'Fill', 'Fill', () => [item('Down', f('fill', 'down'), { hint: 'Ctrl+D' }), item('Right', f('fill', 'right'), { hint: 'Ctrl+R' }),
        item('Up', f('fill', 'up')), item('Left', f('fill', 'left'))]),
      sheetMenu('⌫', 'Clear', 'Clear', () => [item('Clear All', f('clearAs', 'all')), item('Clear Formats', f('clearFormats')), item('Clear Contents', f('clear'), { hint: 'Delete' })]),
      sheetMenu('⇅', 'Sort & Filter', 'Sort & Filter', () => [item('Sort A to Z', f('sortAsc')), item('Sort Z to A', f('sortDesc')), item('Custom Sort...', f('sortBox'))]),
      sheetMenu('🔍', 'Find & Select', 'Find & Select', () => [item('Find...', f('find'), { hint: 'Ctrl+F' }), item('Replace...', f('replace'), { hint: 'Ctrl+H' })]))),
  ];
  paintBars();

  const fillSelect = (sel, values, current, label = String) => {
    const list = values.includes(current) ? values : [current, ...values];
    if (sel.options.length !== list.length || [...sel.options].some((o, i) => o.value !== String(list[i]))) sel.replaceChildren(...list.map((v) => el('option', { value: String(v) }, label(v))));
    sel.value = String(current);
  };
  return {
    groups,
    /** The controls from the active cell's look: { font, fontName, xf, code, fonts: { minor, major }, canUndo, canRedo }. */
    update(look) {
      closeMenu();
      const { font, xf, code } = look;
      b.bold.setOn(font.b);
      b.italic.setOn(font.i);
      b.underline.setOn(!!font.u);
      b.strike.setOn(font.strike);
      b.wrap.setOn(xf.align.wrap);
      b.top.setOn(xf.align.vertical === 'top');
      b.middle.setOn(xf.align.vertical === 'center');
      b.bottom.setOn(xf.align.vertical === 'bottom');
      b.left.setOn(xf.align.horizontal === 'left');
      b.center.setOn(xf.align.horizontal === 'center');
      b.right.setOn(xf.align.horizontal === 'right');
      b.undo.disabled = !look.canUndo;
      b.redo.disabled = !look.canRedo;
      const names = [...new Set([look.fonts.minor, look.fonts.major, ...FONTS])];
      fillSelect(fontSel, names, look.fontName);
      fillSelect(sizeSel, FONT_SIZES, font.size);
      const known = NUMBER_FORMATS.find(([, c]) => c === code);
      if (!known && numSel.options[0].value !== '') numSel.prepend(el('option', { value: '' }, 'Custom'));
      if (known && numSel.options[0].value === '') numSel.options[0].remove();
      numSel.value = known ? code : '';
    },
  };
}
