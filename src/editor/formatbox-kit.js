// The Format Cells box's pieces (formatbox.js): labelled fields, check boxes, lists and a colour chooser, and Excel's
// names for its alignments, underlines, line styles, patterns and shading styles.
import { STANDARD_COLORS, THEME_COLORS } from './sheetmenu.js';

export const make = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids.filter((k) => k != null)); return e; };
export const option = (value, label, selected = false) => make('option', { value, textContent: label, selected });
let fieldCount = 0;
/** A labelled control: [label, control]. */
export const labelled = (text, control) => {
  control.id ||= `oxe-fc-${++fieldCount}`;
  return make('div', { className: 'oxe-fc-field' }, make('label', { htmlFor: control.id, textContent: text }), control);
};
export const check = (text, on) => {
  const input = make('input', { type: 'checkbox', checked: !!on, id: `oxe-fc-${++fieldCount}` });
  return { input, el: make('div', { className: 'oxe-fc-check' }, input, make('label', { htmlFor: input.id, textContent: text })) };
};
export const select = (pairs, value) => make('select', { className: 'oxe-dialog-select' }, ...pairs.map(([v, l]) => option(v, l, v === value)));

/** Excel's horizontal and vertical alignments, underlines, line styles and patterns, by the names its box gives. */
export const HORIZONTAL = [['general', 'General'], ['left', 'Left (Indent)'], ['center', 'Center'], ['right', 'Right (Indent)'], ['fill', 'Fill'],
  ['justify', 'Justify'], ['centerContinuous', 'Center Across Selection'], ['distributed', 'Distributed (Indent)']];
export const VERTICAL = [['top', 'Top'], ['center', 'Center'], ['bottom', 'Bottom'], ['justify', 'Justify'], ['distributed', 'Distributed']];
export const UNDERLINES = [['', 'None'], ['single', 'Single'], ['double', 'Double'], ['singleAccounting', 'Single Accounting'], ['doubleAccounting', 'Double Accounting']];
export const LINES = [['hair', 'Hair'], ['dotted', 'Dotted'], ['dashDotDot', 'Dash dot dot'], ['dashDot', 'Dash dot'], ['dashed', 'Dashed'], ['thin', 'Thin'],
  ['mediumDashDotDot', 'Medium dash dot dot'], ['slantDashDot', 'Slanted dash dot'], ['mediumDashDot', 'Medium dash dot'], ['mediumDashed', 'Medium dashed'],
  ['medium', 'Medium'], ['thick', 'Thick'], ['double', 'Double']];
export const PATTERNS = [['none', 'None'], ['solid', 'Solid'], ['mediumGray', '50% Gray'], ['darkGray', '75% Gray'], ['lightGray', '25% Gray'],
  ['gray125', '12.5% Gray'], ['gray0625', '6.25% Gray'], ['darkHorizontal', 'Horizontal Stripe'], ['darkVertical', 'Vertical Stripe'],
  ['darkDown', 'Reverse Diagonal Stripe'], ['darkUp', 'Diagonal Stripe'], ['darkGrid', 'Diagonal Crosshatch'], ['darkTrellis', 'Thick Diagonal Crosshatch'],
  ['lightHorizontal', 'Thin Horizontal Stripe'], ['lightVertical', 'Thin Vertical Stripe'], ['lightDown', 'Thin Reverse Diagonal Stripe'],
  ['lightUp', 'Thin Diagonal Stripe'], ['lightGrid', 'Thin Horizontal Crosshatch'], ['lightTrellis', 'Thin Diagonal Crosshatch']];
/** Fill Effects' shading styles as gradients: Horizontal runs top to bottom (degree 90), Vertical left to right (0);
 *  the diagonals as 45 and 135, From center a path from the middle. (The degrees are what COM's Gradient.Degree saves,
 *  excel-formatcells.ps1; which degree each of the box's own variants gives has not been measured.) */
export const SHADINGS = [['horizontal', 'Horizontal', { degree: 90 }], ['vertical', 'Vertical', { degree: 0 }], ['diagonalUp', 'Diagonal up', { degree: 45 }],
  ['diagonalDown', 'Diagonal down', { degree: 135 }], ['center', 'From center', { type: 'path', left: 0.5, right: 0.5, top: 0.5, bottom: 0.5 }]];

/** A colour chooser: Automatic (or No Color), the theme's ten, Excel's ten standard ones, or one picked. `value()`
 *  gives { rgb } / { theme } or null. */
export function colorField(current, noneLabel, onChange) {
  const sel = make('select', { className: 'oxe-dialog-select' });
  sel.append(option('none', noneLabel), ...THEME_COLORS.map((n, t) => option(`t${t}`, n)), ...STANDARD_COLORS.map(([n, hex]) => option(`FF${hex}`, n)), option('more', 'More Colors...'));
  const custom = make('input', { type: 'color', value: '#000000', hidden: true, title: 'Color' });
  const known = !current ? 'none' : current.theme != null && !current.tint ? `t${current.theme}` : current.rgb && [...sel.options].some((o) => o.value === current.rgb.toUpperCase()) ? current.rgb.toUpperCase() : current.rgb ? 'more' : 'none';
  sel.value = known;
  if (known === 'more') { custom.hidden = false; custom.value = `#${current.rgb.slice(-6)}`; }
  sel.addEventListener('change', () => { custom.hidden = sel.value !== 'more'; onChange?.(); });
  custom.addEventListener('input', () => onChange?.());
  const value = () => (sel.value === 'none' ? null : sel.value === 'more' ? { rgb: `FF${custom.value.slice(1).toUpperCase()}` } : sel.value[0] === 't' ? { theme: Number(sel.value.slice(1)) } : { rgb: sel.value });
  return { el: make('span', { className: 'oxe-fc-color' }, sel, custom), sel, value };
}

