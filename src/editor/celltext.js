// The one place cell text is sized. Line height is never set on cell content anywhere else: it goes through
// --oxe-cell-lh, set here (on the editor's root for the Normal style, in cellstyles.js's classes for the others), so
// the grid looks the same whatever the host page sets (LESSONS.md section 4: in standards mode every line box gets a
// strut of its parent's font, and an inherited line-height moves text).
import { ptToPx, printRowPx } from '../ooxml/units.js';
import { cellTextMetrics, fontRowPx } from './fonts.js';

const px3 = (v) => `${+v.toFixed(3)}px`;


/** The custom properties that size and place cell text (fonts.js cellTextMetrics): font size, line height, the shifts
 *  for top- and middle-aligned text, and wrapped text's pitch and shifts on screen (--oxe-wrap-*) and in print
 *  (--oxe-pwrap-*). `normalFamily` is the workbook's Normal font, which printed row heights depend on. */
export function cellTextVars(sizePt, family, normalFamily) {
  const px = ptToPx(sizePt);
  // (Wrapped lines step by the font's row height: fonts.js fontRowPx.)
  const rowPx = fontRowPx(family, px);
  const m = cellTextMetrics(family, px, { screen: rowPx, print: printRowPx(rowPx * 0.75, normalFamily) });
  return {
    '--oxe-cell-fs': `${px}px`, '--oxe-cell-lh': px3(m.lh), '--oxe-top-shift': px3(m.top), '--oxe-mid-shift': px3(m.mid),
    '--oxe-wrap-lh': px3(m.wrap.lh), '--oxe-wrap-bottom': px3(m.wrap.bottom), '--oxe-wrap-top': px3(m.wrap.top), '--oxe-wrap-mid': px3(m.wrap.mid),
    '--oxe-pwrap-lh': px3(m.printWrap.lh), '--oxe-pwrap-bottom': px3(m.printWrap.bottom), '--oxe-pwrap-top': px3(m.printWrap.top), '--oxe-pwrap-mid': px3(m.printWrap.mid),
  };
}

/** Set them on an element (the editor's root, for the Normal style's font). */
export function setCellText(el, sizePt, family, normalFamily = family) {
  for (const [k, v] of Object.entries(cellTextVars(sizePt, family, normalFamily))) el.style.setProperty(k, v);
}
