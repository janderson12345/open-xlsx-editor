// The one place cell text is sized. Line height is never set on cell content anywhere else: it goes through
// --oxe-cell-lh, set here, so the grid looks the same whatever the host page sets (LESSONS.md section 4: in standards
// mode every line box gets a strut of its parent's font, and an inherited line-height moves text).
import { ptToPx } from '../ooxml/units.js';

/** Line height as a fraction of the font size: Calibri's hhea ascent + descent + line gap (0.75 + 0.25 + 0.2207 em,
 *  from open-docx-editor's engine/metrics.json). Per-font values come with styles. */
const CALIBRI_LINE = 1.2207;

/** Font size and line height of the default cell text, on the grid's root element. */
export function setCellText(el, { sizePt = 11, line = CALIBRI_LINE } = {}) {
  const px = ptToPx(sizePt);
  el.style.setProperty('--oxe-cell-fs', `${px}px`);
  el.style.setProperty('--oxe-cell-lh', `${Math.round(px * line)}px`);
}
