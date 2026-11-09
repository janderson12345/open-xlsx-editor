// The one place cell text is sized. Line height is never set on cell content anywhere else: it goes through
// --oxe-cell-lh, set here (on the editor's root for the Normal style, in cellstyles.js's classes for the others), so
// the grid looks the same whatever the host page sets (LESSONS.md section 4: in standards mode every line box gets a
// strut of its parent's font, and an inherited line-height moves text).
import { ptToPx } from '../ooxml/units.js';
import { cellLineHeight } from './fonts.js';

/** The custom properties that size cell text: font size, and the line height that puts the baseline where Excel's
 *  is (fonts.js cellLineHeight). */
export function cellTextVars(sizePt, family) {
  const px = ptToPx(sizePt);
  return { '--oxe-cell-fs': `${px}px`, '--oxe-cell-lh': `${+cellLineHeight(family, px).toFixed(3)}px` };
}

/** Set them on an element (the editor's root, for the Normal style's font). */
export function setCellText(el, sizePt, family) {
  for (const [k, v] of Object.entries(cellTextVars(sizePt, family))) el.style.setProperty(k, v);
}
