// The row heights Excel works out itself. A row without customHeight is as tall as its tallest cell, whatever `ht`
// says (Excel's PDF of test/fixtures/rows.xlsx: a stored ht of 30 printed at 15pt, ht 10 with a 20pt font at 26.25pt):
// a cell's font's row height (fonts.js fontRowPx), times its lines when wrapped, never less than the sheet's default.
// Styled empty cells count; merged cells don't.
import { fontRowPx, fontDescentPx } from './fonts.js';

let ctx = null;
/** The lines a wrapped text breaks into in `width` px: words greedily, broken at spaces and at line breaks, measured
 *  in the cell's font; a word wider than the cell broken where it reaches the edge, as Excel breaks it and the cell's
 *  CSS draws it (overflow-wrap: break-word): RSES.C2.10.00001.001 in a 19-character column takes two lines. Without
 *  a canvas (Node) only line breaks count. */
export function wrapLines(text, font, width) {
  const paragraphs = String(text).replace(/\r/g, '').split('\n');
  if (typeof document === 'undefined') return paragraphs;
  ctx = ctx || document.createElement('canvas').getContext('2d');
  ctx.font = font;
  const fits = (s) => ctx.measureText(s).width <= width;
  const out = [];
  for (const para of paragraphs) {
    let line = '';
    for (const word of para.split(' ')) {
      const next = line ? `${line} ${word}` : word;
      if (fits(next)) { line = next; continue; }
      if (line) out.push(line);
      line = word;
      // (A word too wide for a line of its own: as many characters as fit on each line, at least one.)
      while (!fits(line) && line.length > 1) {
        let n = 1;
        while (n < line.length && fits(line.slice(0, n + 1))) n++;
        out.push(line.slice(0, n));
        line = line.slice(n);
      }
    }
    out.push(line);
  }
  return out;
}

/** How much wider Excel's text is than the canvas's when it breaks wrapped lines for a row's height: measured over COM
 *  on a workbook of 58 wrapped rows in Calibri 11 (Range.RowHeight), every row Excel gave one more line needing 1.065 or
 *  more and every row it didn't less than 1.09; 1.075, the middle. turnedPt's 1.0645 for turned text agrees. */
export const EXCEL_WRAP_WIDTH = 1.075;

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
 *  the width 1.0645 times the canvas's and 7px more (Excel's GDI widths are wider): 'Hello world text' at 15, 30,
 *  45, 60 and 90 degrees 36, 52.5, 66, 76.5 and 81pt, worked out here within 2pt; 'Rotated row sized by Excel' at 90
 *  130.5pt (Excel's own height of test/fixtures/rotate.xlsx's row 46: the 1.14 times the canvas used before gave
 *  134.25), 'Row sized at 45' 63pt (66 here). Without a canvas (Node) only stacked text grows. */
export function turnedPt(text, rotation, font, linePt) {
  const plain = text.replace(/[\r\n]/g, '');
  if (rotation === 255) return linePt * [...plain].length + 1.5;
  if (typeof document === 'undefined') return linePt;
  ctx = ctx || document.createElement('canvas').getContext('2d');
  ctx.font = font;
  const a = (Math.abs(rotation > 90 ? 90 - rotation : rotation) * Math.PI) / 180;
  const px = (ctx.measureText(plain).width * 1.0645 + 7) * Math.sin(a) + (linePt / 0.75) * Math.cos(a);
  return Math.max(linePt, px * 0.75);
}

/** Heights in points for the rows Excel sizes: Map(row -> pt), for rows with cells and no customHeight. `deco(s)`
 *  gives a cell format's font ({ family, sizePx, css }) and alignment (cellstyles.js); `cols` the screen columns,
 *  whose widths wrapped lines are counted in (less the cell's 2px padding each side). */
export function autoRowHeights(sheet, deco, cols) {
  const out = new Map();
  if (!deco) return out;
  const mergeOf = (r, c) => sheet.merges.find((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1);
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
      const s = cell.s || 0;
      const d = deco(s);
      if (!rowPt.has(s)) rowPt.set(s, fontRowPx(d.font.family, d.font.sizePx) * 0.75);
      // (A cell in a merged range on this row alone counts with its font's line, as Excel's 16pt title merged across
      // A1:K1 makes its row 21pt (measured over COM); its wrapping doesn't. A range over several rows doesn't count.)
      const m = mergeOf(r, c);
      if (m) { if (m.r0 === m.r1) best = Math.max(best, rowPt.get(s)); continue; }
      let lines = 1;
      // (Any text the cell shows wraps: a string, a formula's text result (t "str", as =CONCATENATE gives) or an inline
      // string; before, only a string counted, so a formula's long text left its row a line tall.)
      // (Lines counted as Excel breaks them: its text measures wider than the canvas's, so a line that just fits here
      // goes on to the next in Excel; EXCEL_WRAP_WIDTH.)
      if (d.align.wrap && (cell.t === 's' || cell.t === 'str' || cell.t === 'inlineStr') && cell.v) lines = countLines(cell.v, d.font.css, (cols.size(c) - 4) / EXCEL_WRAP_WIDTH);
      best = Math.max(best, rowPt.get(s) * lines);
      // (Rich text: a run's larger font grows the row as Excel's does, a 16pt run in Aptos Narrow 11 to 21pt.)
      if (cell.runs) for (const run of cell.runs) if (run.font?.size) best = Math.max(best, fontRowPx(run.font.name && !run.font.scheme ? run.font.name : d.font.family, (run.font.size * 4) / 3) * 0.75);
      if (d.align.rotation && cell.v != null) best = Math.max(best, turnedPt(String(cell.v), d.align.rotation, d.font.css, rowPt.get(s)));
    }
    // (A thick bottom border, thickBot: 0.75pt more, as Excel gives the 21pt title row above 21.75pt, 21 without the
    // attribute (measured over COM). thickTop isn't measured.)
    if (sheet.rows.get(r)?.thickBot) best += 0.75;
    out.set(r, best);
  }
  return out;
}
