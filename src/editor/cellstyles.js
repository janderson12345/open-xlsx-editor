// The workbook's cell formats as CSS: one class per cellXfs index (.oxe-s3: font, size, colour), fill classes
// (.oxe-f3) and border-side classes (.oxe-e3-left: line style and colour, as --oxe-es and --oxe-ec), scoped to one editor in a <style> of its
// own. Cells carry the classes of their `s`; nothing about a cell's look is written on the cell itself.
import { fontFamily } from '../ooxml/styles.js';
import { resolveColor } from '../ooxml/theme.js';
import { fontStack } from './fonts.js';
import { cellTextVars, setCellText } from './celltext.js';

let nextId = 1;

/** How much of a cell a pattern's foreground covers; a hatch or shade is drawn as the blend of its two colours for
 *  now (Excel prints the pattern itself). */
const DENSITY = { darkGray: 0.75, mediumGray: 0.5, lightGray: 0.25, gray125: 0.125, gray0625: 0.0625 };
/** CSS line styles for Excel's border styles (dash lengths are CSS's own, not yet Excel's). */
const LINE = { dotted: 'dotted', hair: 'solid', dashed: 'dashed', mediumDashed: 'dashed', dashDot: 'dashed', mediumDashDot: 'dashed',
  dashDotDot: 'dotted', mediumDashDotDot: 'dotted', slantDashDot: 'dashed', double: 'double' };

const mix = (a, b, t) => '#' + [1, 3, 5].map((i) => Math.round(parseInt(a.slice(i, i + 2), 16) * t + parseInt(b.slice(i, i + 2), 16) * (1 - t)).toString(16).padStart(2, '0')).join('');

/** A fill's colour, or null when there is none. */
function fillColor(fill, theme, palette) {
  if (!fill || fill.pattern === 'none') return null;
  const fg = resolveColor(fill.fg, theme, palette) || '#000000';
  if (fill.pattern === 'solid') return fg;
  // (A pattern's background is white unless given; indexed 64 there is the system background.)
  const bg = fill.bg && fill.bg.indexed !== 64 ? resolveColor(fill.bg, theme, palette) || '#FFFFFF' : '#FFFFFF';
  return mix(fg, bg, DENSITY[fill.pattern] ?? 0.5);
}

/** The CSS declarations for one font. */
function fontRules(font, theme, palette, normalFamily) {
  const family = fontFamily(font, theme);
  const vars = cellTextVars(font.size, family, normalFamily);
  const decoration = [font.u ? 'underline' : '', font.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  const color = resolveColor(font.color, theme, palette);
  return [
    `--oxe-font: ${fontStack(family)}`,
    ...Object.entries(vars).map(([k, v]) => `${k}: ${v}`),
    `font-weight: ${font.b ? 'bold' : 'normal'}`,
    `font-style: ${font.i ? 'italic' : 'normal'}`,
    `text-decoration: ${decoration || 'none'}`,
    ...(font.u === 'double' || font.u === 'doubleAccounting' ? ['text-decoration-style: double'] : []),
    ...(color ? [`color: ${color}`] : []),
  ].join('; ');
}

/** Write a workbook's cell formats into `root` (replacing the last workbook's) and set the Normal style's font on it.
 *  Returns { classOf(s): the cell class, deco(s): { fill, fillClass, sides: [{ name, style, cls }], align } }. */
export function applyCellStyles(root, workbook) {
  if (!root.dataset.oxeId) root.dataset.oxeId = String(nextId++);
  const scope = `.oxe-root[data-oxe-id="${root.dataset.oxeId}"]`;
  const { styles, theme } = workbook;
  const { palette } = styles;
  const normalFamily = fontFamily(styles.normal, theme);
  root.style.setProperty('--oxe-font', fontStack(normalFamily));
  setCellText(root, styles.normal.size, normalFamily, normalFamily);
  const css = [];
  const decos = styles.cellXfs.map((xf, i) => {
    css.push(`${scope} .oxe-s${i} { ${fontRules(styles.fonts[xf.fontId], theme, palette, normalFamily)}; }`);
    const fill = fillColor(styles.fills[xf.fillId], theme, palette);
    if (fill) css.push(`${scope} .oxe-f${i} { background-color: ${fill}; }`);
    const border = styles.borders[xf.borderId];
    const sides = [];
    for (const name of ['left', 'right', 'top', 'bottom']) {
      const side = border[name];
      if (!side.style) continue;
      const cls = `oxe-e${i}-${name}`;
      // (Automatic and missing border colours are black, as Excel draws them.)
      const color = resolveColor(side.color, theme, palette) || '#000000';
      css.push(`${scope} .${cls} { --oxe-es: ${LINE[side.style] || 'solid'}; --oxe-ec: ${color}; }`);
      sides.push({ name, style: side.style, cls });
    }
    return { fill, fillClass: `oxe-f${i}`, sides, align: xf.align };
  });
  let el = root.querySelector(':scope > style.oxe-styles');
  if (!el) {
    el = document.createElement('style');
    el.className = 'oxe-styles';
    root.prepend(el);
  }
  el.textContent = css.join('\n');
  const valid = (s) => (s < styles.cellXfs.length ? s : 0);
  return { classOf: (s) => `oxe-s${valid(s)}`, deco: (s) => decos[valid(s)] };
}
