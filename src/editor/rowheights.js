// The row heights Excel works out itself. A row without customHeight is as tall as its tallest cell, whatever `ht`
// says (Excel's PDF of test/fixtures/rows.xlsx: a stored ht of 30 printed at 15pt, ht 10 with a 20pt font at 26.25pt):
// a cell's font's row height (fonts.js fontRowPx), times its lines when wrapped, never less than the sheet's default.
// Styled empty cells count; merged cells don't.
import { fontRowPx, fontDescentPx } from './fonts.js';

let ctx = null;
/** Lines a wrapped text takes in `width` px: words greedily, broken at spaces and at line breaks, measured in the
 *  cell's font. Without a canvas (Node) only line breaks count. */
export function countLines(text, font, width) {
  const paragraphs = String(text).replace(/\r/g, '').split('\n');
  if (typeof document === 'undefined') return paragraphs.length;
  ctx = ctx || document.createElement('canvas').getContext('2d');
  ctx.font = font;
  let lines = 0;
  for (const para of paragraphs) {
    let line = '';
    lines++;
    for (const word of para.split(' ')) {
      const next = line ? `${line} ${word}` : word;
      if (line && ctx.measureText(next).width > width) {
        lines++;
        line = word;
      } else line = next;
    }
  }
  return lines;
}

/** Each row's baseline: the Windows descent (px) of the tallest font among its cells (styled empty ones count), for
 *  every row with cells. Excel puts all the bottom-aligned text of a row on that one baseline (its PDF of rows.xlsx:
 *  11pt text beside a 20pt cell sits on the 20pt text's baseline, 8pt text beside 11pt on the 11pt one, and a styled
 *  empty 20pt cell lifts the row's text too), and a row of one smaller font keeps that font's own (fonts.xlsx: a row
 *  of Cambria 11 alone sits on Cambria's baseline, not the Normal style's). */
export function rowDescents(sheet, deco) {
  const out = new Map();
  if (!deco) return out;
  for (const [r, line] of sheet.cells) {
    let best = { row: -1, desc: 0 };
    for (const cell of line.values()) {
      const f = deco(cell.s || 0).font;
      const row = fontRowPx(f.family, f.sizePx);
      if (row > best.row) best = { row, desc: fontDescentPx(f.family, f.sizePx) };
    }
    out.set(r, best.desc);
  }
  return out;
}

/** Heights in points for the rows Excel sizes: Map(row -> pt), for rows with cells and no customHeight. `deco(s)`
 *  gives a cell format's font ({ family, sizePx, css }) and alignment (cellstyles.js); `cols` the screen columns,
 *  whose widths wrapped lines are counted in (less the cell's 2px padding each side). */
export function autoRowHeights(sheet, deco, cols) {
  const out = new Map();
  if (!deco) return out;
  const merged = (r, c) => sheet.merges.some((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1);
  const rowPt = new Map(); // per cell format: one line's height in points
  for (const [r, line] of sheet.cells) {
    if (sheet.rows.get(r)?.customHeight) continue;
    let best = sheet.format.defaultRowHeight;
    for (const [c, cell] of line) {
      if (merged(r, c)) continue;
      const s = cell.s || 0;
      const d = deco(s);
      if (!rowPt.has(s)) rowPt.set(s, fontRowPx(d.font.family, d.font.sizePx) * 0.75);
      let lines = 1;
      if (d.align.wrap && cell.t === 's' && cell.v) lines = countLines(cell.v, d.font.css, cols.size(c) - 4);
      best = Math.max(best, rowPt.get(s) * lines);
    }
    out.set(r, best);
  }
  return out;
}
