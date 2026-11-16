// Row and column sizes as commands (commands.js shares its state and helpers as `env`): a heading edge dragged, Column
// Width and Row Height, Default Width, AutoFit, and hiding and showing lines; each one step of the history (sizeStep).
import { MAX_ROWS, MAX_COLS } from '../ooxml/refs.js';
import { typedWidth, pxWidth } from '../ooxml/sizes.js';

export function lineCommands(env) {
  const { ctx, state, linesEach, sizeStep } = env;
  // (Several areas: the lines of each, as Excel sized, hid and showed them; commands.js linesEach.)
  const setEach = (wb, si, axis, spans, change) => { for (const [a, b] of spans) wb.setLines(si, axis, a, b, change); };
  return {
    /** A heading edge dragged to `px` pixels: line `index`, or every selected whole line when it is one of them (as
     *  in Excel). 0 hides. */
    resizeLines(axis, index, px) {
      const spans = linesEach(axis, index);
      const size = Math.max(0, Math.round(px));
      const change = axis === 'col' ? { width: size ? pxWidth(size) : 0 } : { height: size ? size * 0.75 : 0 };
      return sizeStep((wb, si) => setEach(wb, si, axis, spans, change), `${axis === 'col' ? 'Column width' : 'Row height'} changed`, axis === 'col' ? 'formatColumns' : 'formatRows');
    },
    /** Column Width (characters) and Row Height (points) of the selection's columns or rows; 0 hides. */
    columnWidth(chars) {
      const spans = linesEach('col');
      return sizeStep((wb, si) => setEach(wb, si, 'col', spans, { width: chars > 0 ? typedWidth(chars) : 0 }), 'Column width changed');
    },
    rowHeight(points) {
      const spans = linesEach('row');
      // (Every row selected: the sheet's own row height, as Excel writes it, sizes.js setSheetRowHeight.)
      const s = state.sel;
      if (!s.more?.length && s.r0 === 0 && s.r1 >= MAX_ROWS - 1 && s.c0 === 0 && s.c1 >= MAX_COLS - 1 && points > 0) return sizeStep((wb, si) => wb.setSheetSizes(si, { rowHeight: points }), 'Row height changed', 'formatRows');
      return sizeStep((wb, si) => setEach(wb, si, 'row', spans, { height: points }), 'Row height changed', 'formatRows');
    },
    /** Format > Default Width...: every column without a width of its own `chars` wide (sizes.js setDefaultWidth). */
    defaultWidth(chars) { return sizeStep((wb, si) => wb.setSheetSizes(si, { defaultWidth: chars }), 'Default width changed'); },
    /** AutoFit: columns as wide as their widest text (an empty one left as it is); rows back to the height their text
     *  needs. A heading edge's double click fits line `index` (or the selected lines it is one of). */
    fitLines(axis, index = null) {
      const spans = linesEach(axis, index);
      if (axis === 'row') {
        // (Rows that wrapped text makes taller get the height found, as Excel writes it; the others none.)
        const fitted = [];
        for (const [a, b] of spans) for (let r = a; r <= Math.min(b, a + 5000); r++) { const pt = ctx.fitHeight?.(r); if (pt) fitted.push([r, pt]); }
        return sizeStep((wb, si) => { setEach(wb, si, 'row', spans, { height: null }); for (const [r, pt] of fitted) wb.setLines(si, 'row', r, r, { fitted: pt }); }, 'Row height fitted', 'formatRows');
      }
      const widths = [];
      for (const [a, b] of spans) for (let c = a; c <= b; c++) { const px = ctx.fitWidth?.(c); if (px) widths.push([c, px]); }
      if (!widths.length) return false;
      return sizeStep((wb, si) => { for (const [c, px] of widths) wb.setLines(si, 'col', c, c, { width: pxWidth(px), bestFit: true }); }, 'Column width fitted');
    },
    /** Hide (Ctrl+9 rows, Ctrl+0 columns) or show again (Ctrl+Shift+9, Ctrl+Shift+0) the selection's rows or columns. */
    hideLines(axis, hidden = true) {
      const spans = linesEach(axis);
      return sizeStep((wb, si) => setEach(wb, si, axis, spans, { hidden }), `${axis === 'col' ? 'Columns' : 'Rows'} ${hidden ? 'hidden' : 'shown'}`, axis === 'col' ? 'formatColumns' : 'formatRows');
    },
  };
}
