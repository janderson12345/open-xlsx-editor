// The workbook's cell formats as CSS: one class per cellXfs index (.oxe-s3: font, size, colour), fill classes
// (.oxe-f3) and border-side classes (.oxe-e3-left: line style and colour, as --oxe-es and --oxe-ec), scoped to one editor in a style sheet of its
// own (sheets.js). Cells carry the classes of their `s`; nothing about a cell's look is written on the cell itself.
import { fontFamily, formatCodeOf } from '../ooxml/styles.js';
import { resolveColor } from '../ooxml/theme.js';
import { fontStack } from './fonts.js';
import { cellTextVars, setCellText } from './celltext.js';
import { setSheet } from './sheets.js';

let nextId = 1;

/** How much of a cell a pattern's foreground covers; a hatch or shade is drawn as the blend of its two colours for
 *  now (Excel prints the pattern itself). */
const DENSITY = { darkGray: 0.75, mediumGray: 0.5, lightGray: 0.25, gray125: 0.125, gray0625: 0.0625 };
/** CSS line styles for Excel's border styles (dash lengths are CSS's own, not yet Excel's). */
export const LINE = { dotted: 'dotted', hair: 'solid', dashed: 'dashed', mediumDashed: 'dashed', dashDot: 'dashed', mediumDashDot: 'dashed',
  dashDotDot: 'dotted', mediumDashDotDot: 'dotted', slantDashDot: 'dashed', double: 'double' };

const mix = (a, b, t) => '#' + [1, 3, 5].map((i) => Math.round(parseInt(a.slice(i, i + 2), 16) * t + parseInt(b.slice(i, i + 2), 16) * (1 - t)).toString(16).padStart(2, '0')).join('');

/** A gradient fill as CSS: Excel's degree 0 runs left to right, 90 top to bottom (CSS's 90deg and 180deg); a path
 *  gradient spreads out from its rectangle (drawn as a circle from the rectangle's corner, for now). */
function gradientCss({ type, degree, left, top, stops }, theme, palette) {
  const list = stops.map((x) => `${resolveColor(x.color, theme, palette) || '#FFFFFF'} ${(x.position * 100).toFixed(2)}%`).join(', ');
  return type === 'path' ? `radial-gradient(farthest-corner at ${left * 100}% ${top * 100}%, ${list})` : `linear-gradient(${degree + 90}deg, ${list})`;
}

/** A fill's colour (a gradient's CSS background), or null when there is none. */
function fillColor(fill, theme, palette) {
  if (!fill || fill.pattern === 'none') return null;
  if (fill.pattern === 'gradient') return fill.gradient.stops.length ? gradientCss(fill.gradient, theme, palette) : null;
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

/** A rich-text run's font (ooxml/strings.js runFont) as inline CSS: what it gives over the cell's own font. */
function runCss(font, theme, palette) {
  const out = [];
  if (font.name) out.push(`font-family: ${fontStack(fontFamily({ name: font.name, scheme: font.scheme }, theme))}`);
  if (font.size != null) out.push(`font-size: ${(font.size * 4) / 3}px`);
  out.push(`font-weight: ${font.b ? 'bold' : 'normal'}`, `font-style: ${font.i ? 'italic' : 'normal'}`);
  const decoration = [font.u ? 'underline' : '', font.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  out.push(`text-decoration: ${decoration || 'none'}`);
  if (font.u === 'double' || font.u === 'doubleAccounting') out.push('text-decoration-style: double');
  const color = font.color ? resolveColor(font.color, theme, palette) : null;
  if (color) out.push(`color: ${color}`);
  // (Automatic, set on part of a cell in another colour: the sheet's own text colour, not the cell's.)
  else if (font.color?.auto) out.push('color: var(--oxe-fg)');
  // (Raised or lowered as a cell's superscript and subscript are: editor.css .oxe-sup.)
  if (font.vertAlign === 'superscript') out.push('position: relative', 'top: -0.33em');
  if (font.vertAlign === 'subscript') out.push('position: relative', 'top: 0.25em');
  return out.join('; ');
}

/** Write a workbook's cell formats into `root` (replacing the last workbook's) and set the Normal style's font on it.
 *  Returns { classOf(s): the cell class, deco(s): { fill, fillClass, sides: [{ name, style, cls }], align, code },
 *  date1904, color(formatColour) }. */
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
    if (fill) css.push(`${scope} .oxe-f${i} { background: ${fill}; }`);
    const border = styles.borders[xf.borderId];
    const sides = [];
    let diag = null;
    if (border.diagonal?.style && (border.diagonalUp || border.diagonalDown)) {
      const cls = `oxe-e${i}-diag`;
      css.push(`${scope} .${cls} { --oxe-ec: ${resolveColor(border.diagonal.color, theme, palette) || '#000000'}; }`);
      diag = { up: border.diagonalUp, down: border.diagonalDown, style: border.diagonal.style, cls };
    }
    for (const name of ['left', 'right', 'top', 'bottom']) {
      const side = border[name];
      if (!side.style) continue;
      const cls = `oxe-e${i}-${name}`;
      // (Automatic and missing border colours are black, as Excel draws them.)
      const color = resolveColor(side.color, theme, palette) || '#000000';
      css.push(`${scope} .${cls} { --oxe-es: ${LINE[side.style] || 'solid'}; --oxe-ec: ${color}; }`);
      sides.push({ name, style: side.style, cls });
    }
    const f = styles.fonts[xf.fontId];
    const family = fontFamily(f, theme);
    const sizePx = (f.size * 4) / 3;
    // (The font as canvas text measures it, for counting wrapped lines: rowheights.js.)
    const fontCss = `${f.i ? 'italic ' : ''}${f.b ? 'bold ' : ''}${sizePx}px ${fontStack(family)}`;
    const sup = f.vertAlign === 'superscript' ? 'oxe-sup' : f.vertAlign === 'subscript' ? 'oxe-sub' : null;
    return { fill, fillClass: `oxe-f${i}`, sides, diag, sup, align: xf.align, code: formatCodeOf(styles, i), font: { family, sizePx, css: fontCss } };
  });
  setSheet(root, css.join('\n'));
  const valid = (s) => (s < styles.cellXfs.length ? s : 0);
  return {
    classOf: (s) => `oxe-s${valid(s)}`,
    deco: (s) => decos[valid(s)],
    date1904: workbook.date1904,
    // (A rich-text run's font as an inline style, over its cell's: each property the run gives, cells.js.)
    runStyle: (font) => runCss(font, theme, palette),
    // (A number format's colour: [Red] comes as '#FF0000', [ColorN] as an indexed colour of the workbook's palette.)
    color: (c) => (!c ? null : typeof c === 'string' ? c : resolveColor(c, theme, palette)),
  };
}
