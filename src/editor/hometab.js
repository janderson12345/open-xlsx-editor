// The Home tab, Excel's groups in Excel's order: Undo, Font, Alignment, Number, Cells, Editing. Every control runs a
// named command (commands.js, through `act`); update() sets the controls from the active cell, as Excel's do.
import { tableGallery } from './tablegallery.js';
import { cellStyleGallery } from './cellstylegallery.js';
import { CF_GALLERY } from '../ooxml/cfedit.js';
import { el } from './ribbonkit.js';
import { STANDARD_COLORS, THEME_COLORS } from './sheetmenu.js';
import { FONT_SIZES, BORDER_PRESETS } from '../ooxml/formatops.js';
import { LINES } from './formatbox-kit.js';

const FONTS = ['Aptos', 'Aptos Narrow', 'Arial', 'Calibri', 'Cambria', 'Consolas', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Verdana'];
/** The Number Format list: its names and the codes Excel gives (en-US). */
export const NUMBER_FORMATS = [
  ['General', 'General'], ['Number', '0.00'], ['Currency', '"$"#,##0.00'],
  ['Accounting', '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'], ['Short Date', 'm/d/yyyy'],
  ['Long Date', '[$-x-sysdate]dddd, mmmm dd, yyyy'], ['Time', '[$-x-systime]h:mm:ss AM/PM'], ['Percentage', '0.00%'],
  ['Fraction', '# ?/?'], ['Scientific', '0.00E+00'], ['Text', '@'],
];
/** The Data Bars gallery: gradient then solid fill, each in Excel's six colours. */
const BAR_MENU = ['gradient', 'solid'].flatMap((fill) => [['blue', 'Blue'], ['green', 'Green'], ['red', 'Red'], ['orange', 'Orange'], ['lightBlue', 'Light Blue'], ['purple', 'Purple']].map(([c, l]) => [fill, c, l]));
/** The Icon Sets gallery, by Excel's names: Directional, Shapes, Indicators, Ratings. */
const ICON_MENU = [['3Arrows', '3 Arrows (Colored)'], ['3ArrowsGray', '3 Arrows (Gray)'], ['3Triangles', '3 Triangles'], ['4Arrows', '4 Arrows (Colored)'], ['4ArrowsGray', '4 Arrows (Gray)'],
  ['5Arrows', '5 Arrows (Colored)'], ['5ArrowsGray', '5 Arrows (Gray)'], ['3TrafficLights1', '3 Traffic Lights (Unrimmed)'], ['3TrafficLights2', '3 Traffic Lights (Rimmed)'], ['3Signs', '3 Signs'],
  ['4RedToBlack', 'Red To Black'], ['4TrafficLights', '4 Traffic Lights'], ['3Symbols', '3 Symbols (Circled)'], ['3Symbols2', '3 Symbols (Uncircled)'], ['3Flags', '3 Flags'],
  ['3Stars', '3 Stars'], ['4Rating', '4 Ratings'], ['5Quarters', '5 Quarters'], ['5Rating', '5 Ratings'], ['5Boxes', '5 Boxes']];
/** The Borders menu, as Excel lists it. */
const BORDERS = [['bottom', 'Bottom Border'], ['top', 'Top Border'], ['left', 'Left Border'], ['right', 'Right Border'], null,
  ['none', 'No Border'], ['all', 'All Borders'], ['outside', 'Outside Borders'], ['thickOutside', 'Thick Outside Borders'], null,
  ['doubleBottom', 'Bottom Double Border'], ['thickBottom', 'Thick Bottom Border'], ['topBottom', 'Top and Bottom Border'],
  ['topThickBottom', 'Top and Thick Bottom Border'], ['topDoubleBottom', 'Top and Double Bottom Border']];
/** The Borders menu's Draw Borders section: its three modes (drawborders.js). */
const DRAW = [['draw', 'Draw Border'], ['grid', 'Draw Border Grid'], ['erase', 'Erase Border']];
/** The tints of each theme colour's column in Excel's palette (rows 2 to 6), as Excel writes them. */
export const LIGHTER = [0.7999816888943144, 0.5999938962981048, 0.3999755851924192, -0.249977111117893, -0.499984740745262];
export const TINTS = {
  0: [-0.0499893185216834, -0.1499984740745262, -0.249977111117893, -0.3499862666707358, -0.499984740745262],
  1: [0.499984740745262, 0.3499862666707358, 0.249977111117893, 0.1499984740745262, 0.0499893185216834],
  2: [-0.0999786370433668, -0.249977111117893, -0.499984740745262, -0.749992370372631, -0.8999908444471572],
  3: [0.8999908444471572, 0.749992370372631, 0.499984740745262, 0.249977111117893, 0.0999786370433668],
};
const pct = (t) => `${t > 0 ? 'Lighter' : 'Darker'} ${Math.round(Math.abs(t) * 100)}%`;

/** The Home tab. `act(name, ...args)` runs a command; `color(c)` resolves a colour for a swatch. */
export function homeTab(kit, { act, color, workbook = () => null }) {
  const { btn, toggle, grp, col, line, item, menu, split, pick, closeMenu, openMenu, run } = kit;
  const f = (name, ...args) => () => act(name, ...args);

  // A colour palette: the theme's ten in six rows, Excel's ten standard ones, the workbook's Recent Colors (those
  // picked in More Colors, recentcolors.js), `none` (Automatic or No Fill) and More Colors... (`more()`: the colour
  // the box gave, or false).
  const palette = (onPick, none, more) => () => {
    const sw = (c, title) => el('button', { type: 'button', class: 'oxe-swatch', title, 'aria-label': title, style: `background:${color(c) || '#000'}`, onmousedown: run(() => onPick(c)) });
    const theme = el('div', { class: 'oxe-swatches oxe-swatches-theme' });
    for (let row = 0; row < 6; row++) {
      for (let t = 0; t < 10; t++) {
        const tint = row ? (TINTS[t] || LIGHTER)[row - 1] : 0;
        theme.append(sw(tint ? { theme: t, tint } : { theme: t }, `${THEME_COLORS[t]}${tint ? `, ${pct(tint)}` : ''}`));
      }
    }
    const std = el('div', { class: 'oxe-swatches' }, ...STANDARD_COLORS.map(([n, hex]) => sw({ rgb: `FF${hex}` }, n)));
    const recent = act('recentColors') || [];
    const mru = recent.length ? [el('div', { class: 'oxe-cap' }, 'Recent Colors'), el('div', { class: 'oxe-swatches' }, ...recent.map((c) => sw(c, `#${c.rgb.slice(-6)}`)))] : [];
    return [item(none, () => onPick(null)), el('div', { class: 'oxe-cap' }, 'Theme Colors'), theme, el('div', { class: 'oxe-cap' }, 'Standard Colors'), std, ...mru,
      el('hr'), item('More Colors...', () => Promise.resolve(more()).catch(() => false))];
  };
  // (The buttons apply the last colour chosen: red and yellow to start with, as in Excel.)
  let fontColor = { rgb: 'FFFF0000' };
  let fillColor = { rgb: 'FFFFFF00' };
  let border = 'bottom';
  const fontColorBtn = split('A', 'Font Color', () => act('format', { font: { color: fontColor } }),
    palette((c) => { fontColor = c; paintBars(); act('format', { font: { color: c } }); }, 'Automatic',
      () => Promise.resolve(act('moreColorsBox', 'font', fontColor)).then((c) => { if (c) { fontColor = c; paintBars(); } })), { cls: 'oxe-colorbtn oxe-fontcolor' });
  const fillBtn = split('🪣', 'Fill Color', () => act('format', { fill: fillColor }),
    palette((c) => { fillColor = c; paintBars(); act('format', { fill: c }); }, 'No Fill',
      () => Promise.resolve(act('moreColorsBox', 'fill', fillColor)).then((c) => { if (c) { fillColor = c; paintBars(); } })), { cls: 'oxe-colorbtn oxe-fillcolor' });
  const paintBars = () => {
    fontColorBtn.main.style.setProperty('--oxe-swatch', fontColor ? color(fontColor) : '#000000');
    fillBtn.main.style.setProperty('--oxe-swatch', fillColor ? color(fillColor) : 'transparent');
  };
  // (Line Color and Line Style open in the menu's place, as their own menus under the Borders button.)
  const sub = (build) => () => setTimeout(() => openMenu(bordersBtn.more, build));
  const lineColors = palette((c) => act('borderLineColor', c), 'Automatic', () => act('moreColorsBox', 'line'));
  const lineStyles = () => LINES.map(([v, l]) => item(l, f('borderLineStyle', v), { on: act('borderLine')?.style === v, icon: el('span', { class: `oxe-linestyle oxe-ls-${v}` }) }));
  const bordersBtn = split('⊞', 'Bottom Border', () => act('borders', border),
    () => [...BORDERS.map((b) => (b ? item(b[1], () => { border = b[0]; bordersBtn.main.title = b[1]; bordersBtn.main.setAttribute('aria-label', b[1]); act('borders', b[0]); }) : el('hr'))),
      el('hr'), el('div', { class: 'oxe-cap' }, 'Draw Borders'), ...DRAW.map(([mode, label]) => item(label, f('drawBorderMode', mode), { on: act('drawing') === mode })),
      item('Line Color ▸', sub(lineColors)), item('Line Style ▸', sub(lineStyles))]);
  if (BORDERS.filter(Boolean).some(([k]) => !BORDER_PRESETS.includes(k))) throw new Error('a border preset without a rule');

  const fontSel = pick('Font', 'oxe-font', [], (v) => act('format', { font: { name: v } }));
  const sizeSel = pick('Font Size', 'oxe-size', [], (v) => act('format', { font: { size: Number(v) } }));
  // (Picked while part of a cell's text is selected in the cell editor or the formula bar: that part, the edit kept
  // open while the list has focus, richtrack.js keepsEdit.)
  for (const sel of [fontSel, sizeSel]) sel.dataset.oxeKeepsEdit = '';
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
    painter: toggle('🖌', 'Format Painter: press twice to keep it on', f('formatPainter'), { label: 'Format Painter' }),
    redo: btn('↷', 'Redo (Ctrl+Y)', f('redo')),
  };
  const sheetMenu = (icon, label, title, items) => menu(icon, title, items, { label, big: false });
  const groups = [
    grp('Undo', col(b.undo, b.redo)),
    grp('Clipboard', split('📋', 'Paste (Ctrl+V)', f('clipPaste'), () => [item('Paste', f('clipPaste'), { hint: 'Ctrl+V' }), item('Formulas', f('pasteAs', 'formulas')),
      item('Values', f('pasteAs', 'values')), item('Formatting', f('pasteAs', 'formats')), item('Transpose', f('pasteAs', 'transpose')), item('Paste Link', f('pasteAs', 'link')),
      el('hr'), item('Paste Special...', f('pasteSpecialBox'), { hint: 'Ctrl+Alt+V' })], { label: 'Paste', big: true }),
      col(btn('✂', 'Cut (Ctrl+X)', f('clipCut'), { label: 'Cut' }), btn('⧉', 'Copy (Ctrl+C)', f('clipCopy'), { label: 'Copy' }), b.painter)),
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
    grp('Styles', menu('▦', 'Conditional Formatting', () => [
      el('div', { class: 'oxe-cap' }, 'Highlight Cells Rules'), item('Greater Than...', f('cfAdd', 'greaterThan')), item('Less Than...', f('cfAdd', 'lessThan')),
      item('Between...', f('cfAdd', 'between')), item('Equal To...', f('cfAdd', 'equal')), item('Text that Contains...', f('cfAdd', 'containsText')),
      item('A Date Occurring...', f('cfAdd', 'timePeriod')), item('Duplicate Values...', f('cfAdd', 'duplicate')), el('hr'),
      el('div', { class: 'oxe-cap' }, 'Top/Bottom Rules'), item('Top 10 Items...', f('cfAdd', 'top')), item('Top 10%...', f('cfAdd', 'topPercent')),
      item('Bottom 10 Items...', f('cfAdd', 'bottom')), item('Bottom 10%...', f('cfAdd', 'bottomPercent')), item('Above Average...', f('cfAdd', 'above')),
      item('Below Average...', f('cfAdd', 'below')), el('hr'),
      el('div', { class: 'oxe-cap' }, 'Data Bars'), ...BAR_MENU.map(([fill, color, label]) => item(`${label} Data Bar (${fill === 'solid' ? 'Solid' : 'Gradient'} Fill)`, f('cfAdd', 'dataBar', color, fill))), el('hr'),
      el('div', { class: 'oxe-cap' }, 'Icon Sets'), ...ICON_MENU.map(([set, label]) => item(label, f('cfAdd', 'iconSet', set))), el('hr'),
      el('div', { class: 'oxe-cap' }, 'Color Scales'), ...CF_GALLERY.map(([name], k) => item(`${name} Color Scale`, f('cfAdd', 'gallery', k))), el('hr'),
      item('New Rule...', f('cfNewRule')), item('Clear Rules from Selected Cells', f('cfClear', 'selection')), item('Clear Rules from Entire Sheet', f('cfClear', 'sheet')), item('Manage Rules...', f('cfManager'))],
      { label: 'Conditional Formatting', big: true }),
      // (Format as Table: Excel's gallery, the selection made a table in the style picked; acts-tables.js.)
      menu('▤', 'Format as Table', () => [el('div', { class: 'oxe-tstyle-menu' }, ...tableGallery(kit, { wb: workbook(), onPick: (n) => act('formatAsTable', n) }))], { label: 'Format as Table', big: true }),
      // (Cell Styles: Excel's gallery, the selection given the style picked; cellstyleedit.js.)
      menu('▧', 'Cell Styles', () => [el('div', { class: 'oxe-tstyle-menu' }, ...cellStyleGallery(kit, { wb: workbook(), onPick: (n) => act('format', { cellStyle: n }) }))], { label: 'Cell Styles', big: true })),
    grp('Cells', col(
      sheetMenu('⊕', 'Insert', 'Insert', () => [item('Insert Cells...', f('insertCells'), { hint: 'Ctrl++' }), item('Insert Sheet Rows', f('insertRows')), item('Insert Sheet Columns', f('insertColumns')),
        item('Insert Sheet', f('insertSheet'), { hint: 'Shift+F11' })]),
      sheetMenu('🗑', 'Delete', 'Delete', () => [item('Delete Cells...', f('deleteCells'), { hint: 'Ctrl+-' }), item('Delete Sheet Rows', f('deleteRows')), item('Delete Sheet Columns', f('deleteColumns')),
        item('Delete Sheet', f('deleteShownSheet'))]),
      sheetMenu('✎', 'Format', 'Format', () => [el('div', { class: 'oxe-cap' }, 'Cell Size'), item('Row Height...', f('rowHeightBox')), item('AutoFit Row Height', f('fitLines', 'row')),
        item('Column Width...', f('columnWidthBox')), item('AutoFit Column Width', f('fitLines', 'col')), item('Default Width...', f('defaultWidthBox')), el('hr'),
        el('div', { class: 'oxe-cap' }, 'Visibility'), item('Hide Rows', f('hideLines', 'row', true), { hint: 'Ctrl+9' }), item('Hide Columns', f('hideLines', 'col', true), { hint: 'Ctrl+0' }),
        item('Unhide Rows', f('hideLines', 'row', false), { hint: 'Ctrl+Shift+9' }), item('Unhide Columns', f('hideLines', 'col', false)),
        item('Hide Sheet', f('hideShownSheet')), item('Unhide Sheet...', f('unhideSheet')), el('hr'),
        el('div', { class: 'oxe-cap' }, 'Organize Sheets'), item('Rename Sheet', f('renameShownSheet')), el('hr'),
        item('Format Cells...', f('formatCellsBox'), { hint: 'Ctrl+1' })]))),
    grp('Editing', col(
      split('∑', 'AutoSum (Alt+=)', f('autoSum', 'sum'), () => [item('Sum', f('autoSum', 'sum'), { hint: 'Alt+=' }), item('Average', f('autoSum', 'average')),
        item('Count Numbers', f('autoSum', 'count')), item('Max', f('autoSum', 'max')), item('Min', f('autoSum', 'min')),
        // (More Functions...: Insert Function, then Function Arguments, as Excel's AutoSum menu opens them.)
        el('hr'), item('More Functions...', f('insertFunction'))], { label: 'AutoSum' }),
      sheetMenu('⬇', 'Fill', 'Fill', () => [item('Down', f('fill', 'down'), { hint: 'Ctrl+D' }), item('Right', f('fill', 'right'), { hint: 'Ctrl+R' }),
        item('Up', f('fill', 'up')), item('Left', f('fill', 'left')), el('hr'), item('Series...', f('seriesBox')), item('Justify', f('fillJustify')), item('Flash Fill', f('flashFill'), { hint: 'Ctrl+E' })]),
      sheetMenu('⌫', 'Clear', 'Clear', () => [item('Clear All', f('clearAs', 'all')), item('Clear Formats', f('clearFormats')), item('Clear Contents', f('clear'), { hint: 'Delete' }),
        item('Clear Comments and Notes', f('clearAs', 'comments')), item('Clear Hyperlinks', f('clearAs', 'hyperlinks')), item('Remove Hyperlinks', f('clearAs', 'removeHyperlinks'))])),
    // (Sort & Filter and Find & Select big beside them, as in Excel: five stacked made the group, and with it the
    // whole ribbon, five rows tall.)
    menu('⇅', 'Sort & Filter', () => [item('Sort A to Z', f('sortAsc')), item('Sort Z to A', f('sortDesc')), item('Custom Sort...', f('sortBox')), el('hr'),
      item('Filter', f('filterToggle'), { hint: 'Ctrl+Shift+L' }), item('Clear', f('filterClear')), item('Reapply', f('filterReapply'))], { label: 'Sort & Filter', big: true }),
    menu('🔍', 'Find & Select', () => [item('Find...', f('find'), { hint: 'Ctrl+F' }), item('Replace...', f('replace'), { hint: 'Ctrl+H' }), item('Go To...', f('goToBox'), { hint: 'Ctrl+G' }), item('Go To Special...', f('specialBox')), el('hr'),
      // (Excel's shortcuts to Go To Special's choices, the whole sheet looked in when one cell is selected.)
      item('Formulas', f('goToSpecial', { kind: 'formulas', types: 23 })), item('Notes', f('goToSpecial', { kind: 'notes' })),
      item('Conditional Formatting', f('goToSpecial', { kind: 'cfAll' })), item('Constants', f('goToSpecial', { kind: 'constants', types: 23 })),
      item('Data Validation', f('goToSpecial', { kind: 'dvAll' }))], { label: 'Find & Select', big: true })),
  ];
  paintBars();

  const fillSelect = (sel, values, current, label = String) => {
    const list = values.includes(current) ? values : [current, ...values];
    if (sel.options.length !== list.length || [...sel.options].some((o, i) => o.value !== String(list[i]))) sel.replaceChildren(...list.map((v) => el('option', { value: String(v) }, label(v))));
    sel.value = String(current);
  };
  return {
    groups,
    /** Format Painter's button pressed while it waits for cells. */
    painting(on) { b.painter.setOn(on); },
    /** Draw Borders on ('draw', 'grid', 'erase') or off: the Borders button pressed meanwhile. */
    drawing(mode) { bordersBtn.main.classList.toggle('on', !!mode); bordersBtn.main.setAttribute('aria-pressed', String(!!mode)); },
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
