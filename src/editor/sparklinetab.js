// The Sparkline tab, Excel's contextual tab, shown while the active cell holds a sparkline: Sparkline (Edit Data),
// Type (Line, Column, Win/Loss), Show (the points marked), Style (Sparkline Color with its Weight, Marker Color),
// Group (Axis, Group, Ungroup, Clear), and Excel's 36 styles in a gallery (sparklinestyles.js). Its actions are
// acts-sparklines.js's.
import { el } from './ribbonkit.js';
import { STANDARD_COLORS, THEME_COLORS } from './sheetmenu.js';
import { LIGHTER, TINTS } from './hometab.js';
import { SPARKLINE_STYLES, styleColors } from '../ooxml/sparklinestyles.js';

const SHOW = [['high', 'High Point'], ['low', 'Low Point'], ['negative', 'Negative Points'], ['first', 'First Point'], ['last', 'Last Point'], ['markers', 'Markers']];
const MARKERS = [['negative', 'Negative Points'], ['markers', 'Markers'], ['high', 'High Point'], ['low', 'Low Point'], ['first', 'First Point'], ['last', 'Last Point']];
/** Excel's line weights, in points. */
const WEIGHTS = [[0.25, '¼ pt'], [0.5, '½ pt'], [0.75, '¾ pt'], [1, '1 pt'], [1.5, '1½ pt'], [2.25, '2¼ pt'], [3, '3 pt'], [4.5, '4½ pt'], [6, '6 pt']];
const pct = (t) => `${t > 0 ? 'Lighter' : 'Darker'} ${Math.round(Math.abs(t) * 100)}%`;

/** The Sparkline tab. `act(name, ...args)` runs a command; `color(c)` resolves a colour for a swatch. */
export function sparklineTab(kit, { act, color }) {
  const { btn, toggle, grp, col, item, menu, openMenu, run } = kit;
  const f = (name, ...args) => () => act(name, ...args);
  let g = null;
  /** The theme's and the standard colours, as Home's palettes have them, each to `onPick`. */
  const palette = (onPick) => {
    const sw = (c, title) => el('button', { type: 'button', class: 'oxe-swatch', title, 'aria-label': title, style: `background:${color(c) || '#000'}`, onmousedown: run(() => onPick(c)) });
    const theme = el('div', { class: 'oxe-swatches oxe-swatches-theme' });
    for (let row = 0; row < 6; row++) {
      for (let t = 0; t < 10; t++) {
        const tint = row ? (TINTS[t] || LIGHTER)[row - 1] : 0;
        theme.append(sw(tint ? { theme: t, tint } : { theme: t }, `${THEME_COLORS[t]}${tint ? `, ${pct(tint)}` : ''}`));
      }
    }
    return [el('div', { class: 'oxe-cap' }, 'Theme Colors'), theme, el('div', { class: 'oxe-cap' }, 'Standard Colors'),
      el('div', { class: 'oxe-swatches' }, ...STANDARD_COLORS.map(([n, hex]) => sw({ rgb: `FF${hex}` }, n)))];
  };
  // (A submenu opens in its menu's place, as Home's Borders menu opens Line Color.)
  const sub = (owner, build) => () => setTimeout(() => openMenu(owner(), build));
  const typeBtns = { line: toggle('📈', 'Line', f('sparklineSet', { type: 'line' }), { label: 'Line', big: true }),
    column: toggle('▥', 'Column', f('sparklineSet', { type: 'column' }), { label: 'Column', big: true }),
    stacked: toggle('◫', 'Win/Loss', f('sparklineSet', { type: 'stacked' }), { label: 'Win/Loss', big: true }) };
  const showBtns = Object.fromEntries(SHOW.map(([k, label]) => [k, toggle('☐', label, () => act('sparklineSet', { [k]: !g?.[k] }), { label })]));
  const colorBtn = menu('🖌', 'Sparkline Color', () => [...palette((c) => act('sparklineColor', 'series', c)), el('hr'),
    item('Weight', sub(() => colorBtn, () => WEIGHTS.map(([w, l]) => item(l, f('sparklineSet', { lineWeight: w }), { on: g?.lineWeight === w }))))], { label: 'Sparkline Color' });
  // The style gallery: each style shown by its series colour with its high and low points' colours, named as Excel names it.
  const stylesBtn = menu('▤', 'Sparkline Styles', () => {
    const swatch = (k) => {
      const c = styleColors(k);
      const tone = (x) => color({ ...x, tint: x.tint ? Number(x.tint) : 0 }) || '#000';
      const title = `Sparkline Style ${SPARKLINE_STYLES[k][0]}`;
      return el('button', { type: 'button', class: 'oxe-spark-style', title, 'aria-label': title, onmousedown: run(() => act('sparklineStyle', k)) },
        el('span', { class: 'oxe-spark-style-line', style: `background:${tone(c.series)}` }),
        el('span', { class: 'oxe-spark-style-dot oxe-spark-style-high', style: `background:${tone(c.high)}` }),
        el('span', { class: 'oxe-spark-style-dot oxe-spark-style-low', style: `background:${tone(c.low)}` }));
    };
    return [el('div', { class: 'oxe-spark-styles' }, ...SPARKLINE_STYLES.map((x, k) => swatch(k)))];
  }, { label: 'Style', big: true });
  const markerBtn = menu('◆', 'Marker Color', () => MARKERS.map(([k, label]) => item(label, sub(() => markerBtn, () => palette((c) => act('sparklineColor', k, c))))), { label: 'Marker Color' });
  const axisBtn = menu('↔', 'Sparkline Axis', () => {
    const minT = g?.minType || 'individual';
    const maxT = g?.maxType || 'individual';
    return [el('div', { class: 'oxe-cap' }, 'Horizontal Axis Options'),
      item('General Axis Type', f('sparklineSet', { dates: null }), { on: !g?.dates }),
      item('Date Axis Type...', f('sparklineDateAxis'), { on: !!g?.dates }),
      item('Show Axis', f('sparklineSet', { axis: !g?.axis }), { on: !!g?.axis }),
      item('Plot Data Right-to-Left', f('sparklineSet', { rightToLeft: !g?.rightToLeft }), { on: !!g?.rightToLeft }),
      el('div', { class: 'oxe-cap' }, 'Vertical Axis Minimum Value Options'),
      item('Automatic for Each Sparkline', f('sparklineSet', { minType: 'individual' }), { on: minT === 'individual' }),
      item('Same for All Sparklines', f('sparklineSet', { minType: 'group' }), { on: minT === 'group' }),
      item('Custom Value...', f('sparklineAxisValue', 'min'), { on: minT === 'custom' }),
      el('div', { class: 'oxe-cap' }, 'Vertical Axis Maximum Value Options'),
      item('Automatic for Each Sparkline', f('sparklineSet', { maxType: 'individual' }), { on: maxT === 'individual' }),
      item('Same for All Sparklines', f('sparklineSet', { maxType: 'group' }), { on: maxT === 'group' }),
      item('Custom Value...', f('sparklineAxisValue', 'max'), { on: maxT === 'custom' })];
  }, { label: 'Axis' });
  const groups = [
    grp('Sparkline', menu('✎', 'Edit Data', () => [item('Edit Group Location & Data...', f('sparklineGroupData')), item("Edit Single Sparkline's Data...", f('sparklineSingleData')),
      el('hr'), item('Hidden & Empty Cells...', f('sparklineEmpty'))], { label: 'Edit Data', big: true })),
    grp('Type', typeBtns.line, typeBtns.column, typeBtns.stacked),
    grp('Show', col(showBtns.high, showBtns.low, showBtns.negative), col(showBtns.first, showBtns.last, showBtns.markers)),
    grp('Style', stylesBtn, col(colorBtn, markerBtn)),
    grp('Group', axisBtn, col(btn('⊞', 'Group Selected Sparklines', f('sparklineGroup'), { label: 'Group' }), btn('⊟', 'Ungroup Selected Sparklines', f('sparklineUngroup'), { label: 'Ungroup' }),
      menu('✕', 'Clear Sparklines', () => [item('Clear Selected Sparklines', f('sparklineClear', false)), item('Clear Selected Sparkline Groups', f('sparklineClear', true))], { label: 'Clear' }))),
  ];
  return {
    groups,
    /** Shown only while the active cell holds a sparkline. */
    contextual: () => !!act('sparklineAt'),
    update() {
      g = act('sparklineAt') || null;
      for (const [k, b] of Object.entries(typeBtns)) b.setOn(g?.type === k);
      for (const [k, b] of Object.entries(showBtns)) {
        b.setOn(!!g?.[k]);
        b.querySelector('.oxe-ico').textContent = g?.[k] ? '☑' : '☐';
        // (Markers belong to a line: Excel greys the box for columns and win/loss.)
        b.disabled = k === 'markers' && g?.type !== 'line';
      }
    },
  };
}
