// Home > Cell Styles' boxes (acts-more.js routes them here): New Cell Style... and a style's Modify..., Excel's Style
// box (stylebox.js) with Format... opening the Format Cells box (formatbox.js) on the style's own format; what they
// give run as the newCellStyle and modifyCellStyle commands (commands-styles.js).
import { styleBox } from './stylebox.js';
import { formatCellsBox } from './formatbox.js';
import { messageBox } from './dialogs.js';
import { fontFamily, xfOf } from '../ooxml/styles.js';
import { BUILTIN_FORMATS } from '../ooxml/format.js';
import { styleInfo, styleNameRefusal } from '../ooxml/cellstyledefine.js';
import { namedStyle } from '../ooxml/cellstyleedit.js';
import { formatsOf } from '../ooxml/stylesedit.js';
import { readXf } from '../ooxml/styles.js';

export const STYLE_ACTS = new Set(['newCellStyleBox', 'modifyCellStyleBox']);

/** "Style 1", "Style 2"...: the first one the workbook hasn't got, as Excel's box offers. */
const freeName = (wb) => { let k = 1; while (styleNameRefusal(wb, `Style ${k}`)) k++; return `Style ${k}`; };

export function createStyleActs(env, fontNames) {
  const { root, state, editing } = env;
  /** formatbox.js's look of format `xf` with `change` (the Style box's Format... so far) shown in it: its font, number
   *  format and alignment (its fill and border as they were). */
  const lookOf = (xf, change) => {
    const wb = state.workbook;
    const font = { ...wb.styles.fonts[xf.fontId], ...(change?.font || {}) };
    const code = change?.numFmt ?? wb.styles.numFmts[xf.numFmtId] ?? BUILTIN_FORMATS[xf.numFmtId] ?? 'General';
    return { xf: { ...xf, align: { ...xf.align, ...(change?.align || {}) }, protection: { ...xf.protection, ...(change?.protection || {}) } }, font, fontName: font.name && !change?.font?.name ? fontFamily(font, wb.theme) : font.name,
      fontNames: [...new Set([wb.theme.fonts.minor, wb.theme.fonts.major, ...fontNames])], fill: wb.styles.fills[xf.fillId], border: wb.styles.borders[xf.borderId],
      code, cell: null, date1904: wb.date1904, many: false, color: (c) => state.looks?.color(c) };
  };
  const back = () => state.grid.el.focus({ preventScroll: true });
  const refused = (out) => (out?.error ? messageBox(root, out.error).then(() => false) : !!out);
  return (name, args) => {
    const wb = state.workbook;
    if (name === 'newCellStyleBox') {
      // (By example: the active cell's format, as Excel's New Cell Style takes it.)
      const s = editing().selection;
      const cell = wb.sheet(state.sheetIndex).cells.get(s.ar)?.get(s.ac);
      const xf = xfOf(wb.styles, cell?.s || 0);
      return styleBox(root, { name: freeName(wb), includes: {}, byExample: true, look: (ch) => lookOf(xf, ch), format: (look) => formatCellsBox(root, look, 'Number'),
        refusal: (n) => styleNameRefusal(wb, n) })
        .then((got) => { back(); return got ? refused(editing().run('newCellStyle', got.name, got.includes, got.change)) : false; });
    }
    if (name === 'modifyCellStyleBox') {
      const styleName = args[0];
      const info = styleInfo(wb, styleName);
      if (!info) return false;
      // (A built-in style the workbook hasn't got yet: added, as applying it would, to be shown and changed.)
      const xf = info.xf || (() => { const f = formatsOf(wb); const { el } = namedStyle(f, styleName); return readXf(el, f.ns, wb.styles); })();
      return styleBox(root, { name: styleName, rename: !info.builtin, includes: info.includes, look: (ch) => lookOf(xf, ch), format: (look) => formatCellsBox(root, look, 'Number'),
        refusal: (n) => styleNameRefusal(wb, n, styleName) })
        .then((got) => {
          back();
          if (!got) return false;
          return refused(editing().run('modifyCellStyle', styleName, got.change, got.includes, got.name !== styleName ? got.name : null));
        });
    }
    return undefined;
  };
}
