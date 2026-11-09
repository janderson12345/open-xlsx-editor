// The workbook's cell formats as CSS: one class per cellXfs index (.oxe-s3), scoped to one editor, in a <style> of its
// own. Cells carry the class of their `s`; nothing about a cell's look is written on the cell itself.
import { fontFamily } from '../ooxml/styles.js';
import { fontStack } from './fonts.js';
import { cellTextVars, setCellText } from './celltext.js';

let nextId = 1;

/** The CSS declarations for one font. */
function fontRules(font, theme) {
  const family = fontFamily(font, theme);
  const vars = cellTextVars(font.size, family);
  const decoration = [font.u ? 'underline' : '', font.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  return [
    `--oxe-font: ${fontStack(family)}`,
    ...Object.entries(vars).map(([k, v]) => `${k}: ${v}`),
    `font-weight: ${font.b ? 'bold' : 'normal'}`,
    `font-style: ${font.i ? 'italic' : 'normal'}`,
    `text-decoration: ${decoration || 'none'}`,
    ...(font.u === 'double' || font.u === 'doubleAccounting' ? ['text-decoration-style: double'] : []),
  ].join('; ');
}

/** Write a workbook's cell formats into `root` (replacing the last workbook's) and set the Normal style's font on it.
 *  Returns the class for a cell format index. */
export function applyCellStyles(root, workbook) {
  if (!root.dataset.oxeId) root.dataset.oxeId = String(nextId++);
  const scope = `.oxe-root[data-oxe-id="${root.dataset.oxeId}"]`;
  const { styles, theme } = workbook;
  const normalFamily = fontFamily(styles.normal, theme);
  root.style.setProperty('--oxe-font', fontStack(normalFamily));
  setCellText(root, styles.normal.size, normalFamily);
  const css = styles.cellXfs.map((xf, i) => `${scope} .oxe-s${i} { ${fontRules(styles.fonts[xf.fontId], theme)}; }`).join('\n');
  let el = root.querySelector(':scope > style.oxe-styles');
  if (!el) {
    el = document.createElement('style');
    el.className = 'oxe-styles';
    root.prepend(el);
  }
  el.textContent = css;
  return (s) => `oxe-s${s < styles.cellXfs.length ? s : 0}`;
}
