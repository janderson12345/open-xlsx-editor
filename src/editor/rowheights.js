// The row heights Excel works out itself. A row without customHeight is as tall as its tallest cell, whatever `ht`
// says (Excel's PDF of test/fixtures/rows.xlsx: a stored ht of 30 printed at 15pt, ht 10 with a 20pt font at 26.25pt):
// a cell's font's row height (fonts.js fontRowPx), times its lines when wrapped, never less than the sheet's default.
// Styled empty cells count; merged cells don't.
import { fontRowPx, fontDescentPx } from './fonts.js';

let ctx = null;
/** The lines a wrapped text breaks into in `width` px: words greedily, broken at spaces and at line breaks, measured
 *  in the cell's font. Without a canvas (Node) only line breaks count. */
export function wrapLines(text, font, width) {
  const paragraphs = String(text).replace(/\r/g, '').split('\n');
  if (typeof document === 'undefined') return paragraphs;
  ctx = ctx || document.createElement('canvas').getContext('2d');
  ctx.font = font;
  const out = [];
  for (const para of paragraphs) {
    let line = '';
    for (const word of para.split(' ')) {
      const next = line ? `${line} ${word}` : word;
      if (line && ctx.measureText(next).width > width) {
        out.push(line);
        line = word;
      } else line = next;
    }
    out.push(line);
  }
  return out;
}

/** How many lines a wrapped text takes in `width` px (wrapLines). */
export const countLines = (text, font, width) => wrapLines(text, font, width).length;

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

/** The height (pt) Excel gives a row for turned text (Format Cells > Alignment > Orientation), measured over COM with
 *  Calibri 11 (Rows.AutoFit after Range.Orientation): stacked letters (255) a line each and 1.5pt more ('ab' 31.5,
 *  'Stack' 76.5, 'Hello world text' 241.5); text at an angle its width up the angle and a line's height across it,
 *  the width 1.14 times the canvas's (Excel's GDI widths are wider): 'Hello world text' at 15, 30, 45, 60 and 90
 *  degrees 36, 52.5, 66, 76.5 and 81pt, worked out here within 2pt. Without a canvas (Node) only stacked text grows. */
export function turnedPt(text, rotation, font, linePt) {
  const plain = text.replace(/[\r\n]/g, '');
  if (rotation === 255) return linePt * [...plain].length + 1.5;
  if (typeof document === 'undefined') return linePt;
  ctx = ctx || document.createElement('canvas').getContext('2d');
  ctx.font = font;
  const a = (Math.abs(rotation > 90 ? 90 - rotation : rotation) * Math.PI) / 180;
  const px = ctx.measureText(plain).width * 1.14 * Math.sin(a) + (linePt / 0.75) * Math.cos(a);
  return Math.max(linePt, px * 0.75);
}

/** Heights in points for the rows Excel sizes: Map(row -> pt), for rows with cells and no customHeight. `deco(s)`
 *  gives a cell format's font ({ family, sizePx, css }) and alignment (cellstyles.js); `cols` the screen columns,
 *  whose widths wrapped lines are counted in (less the cell's 2px padding each side). */
export function autoRowHeights(sheet, deco, cols) {
  const out = new Map();
  if (!deco) return out;
  const merged = (r, c) => sheet.merges.some((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1);
  const rowPt = new Map(); // per cell format: one line's height in points
  const lineOf = (s) => {
    if (!rowPt.has(s)) { const d = deco(s); rowPt.set(s, fontRowPx(d.font.family, d.font.sizePx) * 0.75); }
    return rowPt.get(s);
  };
  // (A row with a format of its own, cells or not: its font's line, as Excel grows a row made 14pt to 18.75.)
  const rows = new Set(sheet.cells.keys());
  for (const [r, x] of sheet.rows) if (x.s) rows.add(r);
  for (const r of rows) {
    if (sheet.rows.get(r)?.customHeight) continue;
    const line = sheet.cells.get(r) || new Map();
    let best = sheet.format.defaultRowHeight;
    if (sheet.rows.get(r)?.s) best = Math.max(best, lineOf(sheet.rows.get(r).s));
    for (const [c, cell] of line) {
      if (merged(r, c)) continue;
      const s = cell.s || 0;
      const d = deco(s);
      if (!rowPt.has(s)) rowPt.set(s, fontRowPx(d.font.family, d.font.sizePx) * 0.75);
      let lines = 1;
      if (d.align.wrap && cell.t === 's' && cell.v) lines = countLines(cell.v, d.font.css, cols.size(c) - 4);
      best = Math.max(best, rowPt.get(s) * lines);
      // (Rich text: a run's larger font grows the row as Excel's does, a 16pt run in Aptos Narrow 11 to 21pt.)
      if (cell.runs) for (const run of cell.runs) if (run.font?.size) best = Math.max(best, fontRowPx(run.font.name && !run.font.scheme ? run.font.name : d.font.family, (run.font.size * 4) / 3) * 0.75);
      if (d.align.rotation && cell.v != null) best = Math.max(best, turnedPt(String(cell.v), d.align.rotation, d.font.css, rowPt.get(s)));
    }
    out.set(r, best);
  }
  return out;
}
