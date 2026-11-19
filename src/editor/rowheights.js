// The row heights Excel works out itself. A row without customHeight is as tall as its tallest cell, whatever `ht`
// says (Excel's PDF of test/fixtures/rows.xlsx: a stored ht of 30 printed at 15pt, ht 10 with a 20pt font at 26.25pt):
// a cell's font's row height (fonts.js fontRowPx), times its lines when wrapped. Styled empty cells count; merged cells
// don't. And every one of the row's 16,384 cells counts, the ones the file doesn't hold too, with the font of the row's
// own format or else their column's (the Normal style's in a column without one): standardRowHeight.
import { filledRows } from '../ooxml/sheetcells.js';
import { MAX_COLS } from '../ooxml/refs.js';
import { fontRowPx, fontDescentPx, windowsHasFont } from './fonts.js';
import { excelWidths } from '../ooxml/excelwidths.js';

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

/** The tallest a row can be in Excel, in points (409.5): a row whose wrapped text would take more stops there, as 18
 *  rows of 5 customer workbooks did in Excel where they grew to 3,765pt here. */
export const MAX_ROW_PT = 409.5;

/** How much wider Excel's text is than the canvas's when it breaks wrapped lines for a row's height, for a font
 *  excelwidths.js hasn't: measured over COM on a workbook of 58 wrapped rows in Calibri 11 (Range.RowHeight), every
 *  row Excel gave one more line needing 1.065 or more and every row it didn't less than 1.09; 1.075, the middle.
 *  turnedPt's 1.0645 for turned text agrees. */
export const EXCEL_WRAP_WIDTH = 1.075;

/** How many lines a wrapped text takes in `width` px (wrapLines). */
export const countLines = (text, font, width) => wrapLines(text, font, width).length;

/** The words Excel keeps whole when it wraps: a word ends after its spaces, or after a hyphen and the spaces that
 *  follow it. Nothing else is a place to break: not a slash, a comma, a full stop, a bracket, an en or an em dash. */
const WORDS = /[^ -]*(?:-? +|-)|[^ -]+$/g;

/** The lines Excel breaks a wrapped `text` into, as strings; `fits(a, b)` says whether characters a to b (b not
 *  included) of the text, carriage returns taken out, fit on a line. Measured over COM (tools/excel-compare/
 *  excel-wrapwidths.ps1: rows AutoFit at every column width from 12 to 420px, 64 sentences in 16 fonts and 58 strings
 *  of hyphens, spaces and punctuation, every line count Excel's from a column of 15px up):
 *  - words go on a line while they fit, and the line's width is counted without the spaces it ends with and without
 *    one hyphen before them: "TENTATIVE -" is a line where "TENTATIVE" fits, and of 20 hyphens 19 are counted;
 *  - spaces a line begins with count ("  indented" is wider), as do spaces between its words;
 *  - a word too wide for a line of its own is cut where it reaches the edge, a character at least, its hyphen counted
 *    like any character there, and the words after it go on along its last piece;
 *  - a line break makes a line, an empty one too ("abc\n\n" is three lines), but for one case: after a line that is
 *    one space alone, a last line that is empty or spaces isn't counted ("abc\n \n" and "abc\n \n  " are two lines
 *    and " \n" one; "abc\n  \n", with two spaces, is three). Text pasted from elsewhere often ends so. */
export function breakLines(text, fits) {
  const s = String(text).replace(/\r/g, '');
  return breakSpans(s, fits).map(([a, b]) => s.slice(a, b));
}

/** breakLines' lines as [from, to] in `s` (no carriage returns in it). */
function breakSpans(s, fits) {
  const hung = (a, b) => {
    while (b > a && s[b - 1] === ' ') b--;
    return b > a && s[b - 1] === '-' ? b - 1 : b;
  };
  const out = [];
  let from = 0;
  const paras = s.split('\n');
  if (paras.length > 1 && paras[paras.length - 2] === ' ' && /^ *$/.test(paras[paras.length - 1])) paras.pop();
  for (const para of paras) {
    let a = from;
    let b = from;
    for (const word of para.match(WORDS) || []) {
      const end = b + word.length;
      if (fits(a, hung(a, end))) { b = end; continue; }
      if (b > a) out.push([a, b]);
      a = b;
      b = end;
      while (!fits(a, hung(a, b))) {
        let e = b;
        while (e > a && s[e - 1] === ' ') e--;
        if (e - a <= 1) break;
        let k = 1;
        while (a + k < e && fits(a, a + k + 1)) k++;
        out.push([a, a + k]);
        a += k;
      }
    }
    out.push([a, b]);
    from += para.length + 1;
  }
  return out;
}

const otherWidths = new Map();
/** The width of a character excelwidths.js hasn't (outside printable ASCII), in whole pixels: the canvas's, rounded;
 *  without a canvas (Node) half the font's size, the whole of it from the CJK blocks on. */
function otherWidth(ch, css, px) {
  if (typeof document === 'undefined') return Math.round(ch.charCodeAt(0) >= 0x2e80 ? px : px / 2);
  const key = `${css}\n${ch}`;
  let w = otherWidths.get(key);
  if (w === undefined) {
    ctx = ctx || document.createElement('canvas').getContext('2d');
    ctx.font = css;
    w = Math.round(ctx.measureText(ch).width);
    if (otherWidths.size > 20000) otherWidths.clear();
    otherWidths.set(key, w);
  }
  return w;
}

/** The lines Excel breaks a wrapped cell's text into when it works out the row's height, in a column `colPx` wide
 *  (less `indentPx` of indent): { count, spans, starts }, `spans` the lines as [from, to] in the text and `starts`
 *  where each rich-text run begins in it (both null when the text isn't measured run by run). `font` is the cell format's ({ family, sizePx, bold, roman, css },
 *  cellstyles.js); `runs` a rich text's ([{ text, font }], strings.js), each measured in its own font.
 *  Excel measures in the whole pixels Windows gives each character (excelwidths.js; Small Fonts' under 8pt, and one
 *  pixel more a line in its made-up bold), and a line fits when
 *      width + 2 x ceil(digit / 4) + max(3, ceil(width / 28)) <= the column's pixels,
 *  `digit` being the width of 0 in the cell's font: 7px of margin up to 84px of text in Calibri 11, a pixel more
 *  for every 28px from there (20 digits, 140px, fit in 149; 100 in 729), 9px to begin with in Calibri 14 to 18, 11 in
 *  20 and 24. Measured over COM (excel-wrapwidths.ps1): the narrowest column that keeps 2 to 100 of one character on
 *  a line, for 95 characters in 9 fonts and 3 characters in some 50, each Excel's to the pixel; alignment left, centre
 *  or right changes nothing, and an indent level takes 9px (Calibri 11, Times New Roman 12 and Book Antiqua 12 bold).
 *  Before, lines were counted at the canvas's widths times EXCEL_WRAP_WIDTH in the column less 4px, right for most
 *  rows but a line out for long sentences in wide columns and for headings in narrow ones; a font the table hasn't
 *  (or a size over 36pt) is still counted so, and one Windows hasn't in Arial's widths. Italic is counted at the
 *  upright widths (not measured). Without a canvas (Node) a font the table hasn't counts line breaks only. */
function excelLines(text, font, colPx, indentPx, runs) {
  const s = String(text).replace(/\r/g, '');
  const width = colPx - indentPx;
  // (A font Windows hasn't is laid out in Arial, a roman one in Times New Roman, as its rows are: fontRowPx.)
  const face = (name, roman) => (windowsHasFont(name) ? name : roman ? 'Times New Roman' : 'Arial');
  const own = excelWidths(face(font.family, font.roman), font.sizePx, font.bold);
  if (!own) {
    if (typeof document === 'undefined') return { count: s.split('\n').length, spans: null, starts: null };
    ctx = ctx || document.createElement('canvas').getContext('2d');
    ctx.font = font.css;
    const limit = (width - 4) / EXCEL_WRAP_WIDTH;
    return { count: breakSpans(s, (a, b) => ctx.measureText(s.slice(a, b)).width <= limit).length, spans: null, starts: null };
  }
  // (Each character's width, in its run's font when the runs make up the text and the table has every font.)
  const cw = new Uint16Array(s.length);
  let extra = own.small && font.bold ? 1 : 0;
  const fill = (from, to, table) => {
    for (let i = from; i < to; i++) { const w = table.width(s.charCodeAt(i)); cw[i] = w < 0 ? otherWidth(s[i], font.css, font.sizePx) : w; }
  };
  const parts = runs && runs.map((r) => String(r.text).replace(/\r/g, ''));
  let starts = null;
  if (parts && parts.join('') === s) {
    const tables = runs.map((r) => (r.font ? excelWidths(r.font.name && !r.font.scheme ? face(r.font.name, false) : face(font.family, font.roman), r.font.size ? (r.font.size * 4) / 3 : font.sizePx, r.font.b ?? font.bold) : own));
    if (tables.every(Boolean)) {
      starts = [];
      let at = 0;
      parts.forEach((part, k) => { starts.push(at); fill(at, at + part.length, tables[k]); at += part.length; });
      extra = 0;
    }
  }
  if (!starts) fill(0, s.length, own);
  const pad = 2 * Math.ceil(own.digit / 4);
  const spans = breakSpans(s, (a, b) => {
    let sum = b > a ? extra : 0;
    for (let i = a; i < b; i++) { sum += cw[i]; if (sum + pad + 3 > width) return false; }
    return sum + pad + Math.max(3, Math.ceil(sum / 28)) <= width;
  });
  return { count: spans.length, spans: starts && spans, starts };
}

/** How many lines Excel breaks a wrapped cell's text into (excelLines). */
export const excelLineCount = (text, font, colPx, indentPx = 0, runs = null) => excelLines(text, font, colPx, indentPx, runs).count;

/** The height in points Excel gives a wrapped cell's text: its lines (excelLines) times `linePt`, the line of the
 *  cell's font; a rich text's, each line as tall as the tallest of the runs on it, `runPt(run)` giving a run's line
 *  (measured over COM, Rows.AutoFit at every column width from 40 to 400px: "Quarterly Review Totals" in Calibri 12
 *  bold followed by 36 characters at 10pt is 15.75pt for each line with a 12pt character on it and 12.75pt for each
 *  line of 10pt alone, and a 16pt run after 10pt text makes only its own lines 21pt). */
export function excelWrapPt(text, font, linePt, colPx, indentPx = 0, runs = null, runPt = null) {
  const { count, spans, starts } = excelLines(text, font, colPx, indentPx, runs);
  if (!spans || !runPt) return count * linePt;
  const lineOf = (r) => (runs[r].font ? runPt(runs[r]) : linePt);
  let pt = 0;
  let r = 0;
  for (const [a, b] of spans) {
    // (The run the line begins in, then every run that begins on it with a character of its own: one pass over the
    // runs for all the lines.)
    while (r + 1 < starts.length && starts[r + 1] <= a) r++;
    let tallest = lineOf(r);
    for (let q = r + 1; q < starts.length && starts[q] < b; q++) {
      if ((q + 1 < starts.length ? starts[q + 1] : Infinity) > starts[q]) tallest = Math.max(tallest, lineOf(q));
    }
    pt += tallest;
  }
  return pt;
}

/** Each row's baseline: the Windows descent (px) of the tallest font among its cells (styled empty ones count), for
 *  every row with cells. Excel puts all the bottom-aligned text of a row on that one baseline (its PDF of rows.xlsx:
 *  11pt text beside a 20pt cell sits on the 20pt text's baseline, 8pt text beside 11pt on the 11pt one, and a styled
 *  empty 20pt cell lifts the row's text too), and a row of one smaller font keeps that font's own (fonts.xlsx: a row
 *  of Cambria 11 alone sits on Cambria's baseline, not the Normal style's). */
export function rowDescents(sheet, deco) {
  const out = new Map();
  if (!deco) return out;
  for (const [r, line] of filledRows(sheet.cells)) {
    let best = { row: -1, desc: 0 };
    for (const cell of line.values()) {
      const f = deco(cell.s || 0).font;
      const row = fontRowPx(f.family, f.sizePx, f.bold, f.roman);
      if (row > best.row) best = { row, desc: fontDescentPx(f.family, f.sizePx) };
    }
    out.set(r, best.desc);
  }
  // (Rows of formatted empty cells kept as bands, sheetcells.js: one answer for a band's rows, from its runs.)
  for (const { r0, r1, runs } of sheet.cells.bands || []) {
    let best = { row: -1, desc: 0 };
    for (let k = 0; k < runs.length; k += 3) {
      const f = deco(runs[k + 2] || 0).font;
      const row = fontRowPx(f.family, f.sizePx, f.bold, f.roman);
      if (row > best.row) best = { row, desc: fontDescentPx(f.family, f.sizePx) };
    }
    for (let r = r0; r <= r1; r++) out.set(r, best.desc);
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

/** One line's height in points for each cell format, from `deco` (cellstyles.js), kept. */
function linesOf(deco) {
  const rowPt = new Map();
  return (s) => {
    if (!rowPt.has(s)) { const d = deco(s); rowPt.set(s, fontRowPx(d.font.family, d.font.sizePx, d.font.bold, d.font.roman) * 0.75); }
    return rowPt.get(s);
  };
}

/** The sheet's 16,384 columns as stretches of one format each, tallest font first: [{ min, max, pt }], a `<col>`
 *  with its style's line, the columns between with the Normal style's (cell format 0). */
function columnLines(sheet, lineOf) {
  const out = [];
  let next = 0;
  for (const col of [...(sheet.cols || [])].sort((a, b) => a.min - b.min)) {
    if (col.max < next) continue;
    const min = Math.max(col.min, next);
    if (min > next) out.push({ min: next, max: min - 1, pt: lineOf(0) });
    out.push({ min, max: col.max, pt: lineOf(col.style || 0) });
    next = col.max + 1;
  }
  if (next < MAX_COLS) out.push({ min: next, max: MAX_COLS - 1, pt: lineOf(0) });
  return out.sort((a, b) => b.pt - a.pt);
}

/** The tallest line among the columns a row holds no cell in (`count` cells, `has(c)`): what the cells the file
 *  doesn't hold make the row, 0 when it holds all 16,384. */
function emptyCellsPt(columns, count, has) {
  for (const seg of columns) {
    if (seg.max - seg.min + 1 > count) return seg.pt;
    for (let c = seg.min; c <= seg.max; c++) if (!has(c)) return seg.pt;
  }
  return 0;
}

/** The sheet's standard row height in points (Excel's Worksheet.StandardHeight): the height of a row the file says
 *  nothing about, and the least a row of cells without a format of its own takes. Measured over COM on 240 workbooks
 *  storing 10 to 21.9pt under Calibri 11 and Arial 10, and 17 more with column styles (tools/excel-compare/
 *  stdheight-book.mjs holds the cases): sheetFormatPr's defaultRowHeight counts only with customHeight="1" (20 stays
 *  20, 10 stays 10). Without it Excel never reads the number (a file saved on a screen at 125% stores 14.4 for
 *  Calibri 11's 15, at 150% 14.5) and works the height out as any row's: the tallest font among the row's 16,384
 *  cells, each empty one with its column's format, the Normal style's where the column has none. So it is the Normal
 *  font's line, unless every column has a style of smaller fonts (all Calibri 6: 8.25pt; all but one: 15) or some
 *  column, hidden or not, a taller one (five columns of Calibri 20: 26.25). Without `deco` the stored number. */
export function standardRowHeight(sheet, deco) {
  if (!deco || sheet.format.customHeight) return sheet.format.defaultRowHeight;
  return columnLines(sheet, linesOf(deco))[0].pt;
}

/** Heights in points for the rows Excel sizes: Map(row -> pt), for rows with cells and no customHeight, and the
 *  sheet's standard height (standardRowHeight) as the map's `standard`, for every other row. `deco(s)` gives a cell
 *  format's font ({ family, sizePx, css }) and alignment (cellstyles.js); `cols` the screen columns, whose widths
 *  wrapped lines are counted in (less the cell's 2px padding each side).
 *  A row's least height is its empty cells': the font of its own format when it has one (a row made Calibri 6 with
 *  nothing in it is 8.25pt under a Normal style of Calibri 11, measured over COM), else the tallest font among the
 *  columns it holds no cell in (emptyCellsPt); sheetFormatPr's defaultRowHeight is no part of it, with customHeight
 *  or without (a row element without `ht` in a sheet stored at 20pt with customHeight is 15pt). */
export function autoRowHeights(sheet, deco, cols) {
  const out = new Map();
  if (!deco) return out;
  const mergeOf = (r, c) => sheet.merges.find((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1);
  const lineOf = linesOf(deco);
  const columns = columnLines(sheet, lineOf);
  const stored = !!sheet.format.customHeight;
  out.standard = stored ? sheet.format.defaultRowHeight : columns[0].pt;
  // (A row with a format of its own, cells or not: its font's line, as Excel grows a row made 14pt to 18.75.)
  // (Rows of formatted empty cells kept as bands, sheetcells.js, are worked out from their runs below, not made.)
  // (With a stored standard height, customHeight, every row the file has an element for is worked out: Excel gives
  // such a row its fonts' height, not the stored one.)
  const rows = new Set();
  for (const [r] of filledRows(sheet.cells)) rows.add(r);
  for (const [r, x] of sheet.rows) if ((x.s || stored) && !sheet.cells.blankRuns?.(r)) rows.add(r);
  for (const { r0, r1, runs } of sheet.cells.bands || []) {
    let count = 0;
    let tallest = 0;
    for (let k = 0; k < runs.length; k += 3) { count += runs[k + 1] - runs[k] + 1; tallest = Math.max(tallest, lineOf(runs[k + 2] || 0)); }
    // (The runs are in column order, as sheetscan.js makes them, so a binary search: a crafted band of thousands of
    // runs, met column by column for all 16,384, would otherwise cost runs x columns for every band.)
    const inRuns = (c) => {
      let lo = 0;
      let hi = runs.length / 3 - 1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        if (c < runs[mid * 3]) hi = mid - 1;
        else if (c > runs[mid * 3 + 1]) lo = mid + 1;
        else return true;
      }
      return false;
    };
    const least = emptyCellsPt(columns, count, inRuns);
    const merged = sheet.merges.filter((m) => m.r1 >= r0 && m.r0 <= r1);
    for (let r = r0; r <= r1; r++) {
      const own = sheet.rows.size ? sheet.rows.get(r) : null;
      if (own?.customHeight) continue;
      let cells = tallest;
      if (merged.some((m) => r >= m.r0 && r <= m.r1)) {
        // (An empty cell in a merged range counts as any cell does: with its font's line on a range of one row.)
        cells = 0;
        for (let k = 0; k < runs.length; k += 3) {
          for (let c = runs[k]; c <= runs[k + 1]; c++) {
            const m = mergeOf(r, c);
            if (!m || m.r0 === m.r1) cells = Math.max(cells, lineOf(runs[k + 2] || 0));
          }
        }
      }
      let best = Math.max(cells, own?.s ? lineOf(own.s) : least);
      if (own?.thickBot) best += 0.75;
      out.set(r, Math.min(best, MAX_ROW_PT));
    }
  }
  for (const r of rows) {
    const own = sheet.rows.get(r);
    if (own?.customHeight) continue;
    const line = sheet.cells.get(r) || new Map();
    let best = own?.s ? lineOf(own.s) : emptyCellsPt(columns, line.size, (c) => line.has(c));
    for (const [c, cell] of line) {
      const s = cell.s || 0;
      const d = deco(s);
      const pt = lineOf(s);
      // (A cell in a merged range on this row alone counts with its font's line, as Excel's 16pt title merged across
      // A1:K1 makes its row 21pt (measured over COM); its wrapping doesn't. A range over several rows doesn't count.)
      const m = mergeOf(r, c);
      if (m) { if (m.r0 === m.r1) best = Math.max(best, pt); continue; }
      let wrapped = pt;
      // (Any text the cell shows wraps: a string, a formula's text result (t "str", as =CONCATENATE gives) or an inline
      // string; before, only a string counted, so a formula's long text left its row a line tall.)
      // (Lines counted as Excel breaks them, at Windows' whole-pixel widths and Excel's margin: excelLineCount. An
      // indent level takes 9px of the column, none when the text is centred.)
      // (A hidden column's text doesn't wrap: Excel gives it its font's one line, where counting it at no width made
      // a line of each character, 390pt rows of 5 customer workbooks that Excel has at 30 to 72.)
      // (A rich text's run in a larger font grows the lines it is on, as Excel's does: a 16pt run in Aptos Narrow 11
      // makes a row of one line 21pt.)
      const runPt = (run) => (run.font.size || run.font.name || run.font.b != null ? fontRowPx(run.font.name && !run.font.scheme ? run.font.name : d.font.family, run.font.size ? (run.font.size * 4) / 3 : d.font.sizePx, run.font.b ?? d.font.bold) * 0.75 : pt);
      if (d.align.wrap && cols.size(c) > 0 && (cell.t === 's' || cell.t === 'str' || cell.t === 'inlineStr') && cell.v) wrapped = excelWrapPt(cell.v, d.font, pt, cols.size(c), d.align.indent && d.align.horizontal !== 'center' ? d.align.indent * 9 : 0, cell.runs, runPt);
      best = Math.max(best, wrapped);
      // (Not wrapped, or not measured run by run: a run's larger font still counts, with its one line at least.)
      if (cell.runs) for (const run of cell.runs) if (run.font?.size) best = Math.max(best, fontRowPx(run.font.name && !run.font.scheme ? run.font.name : d.font.family, (run.font.size * 4) / 3, run.font.b ?? d.font.bold) * 0.75);
      if (d.align.rotation && cell.v != null) best = Math.max(best, turnedPt(String(cell.v), d.align.rotation, d.font.css, pt));
    }
    // (A thick bottom border, thickBot: 0.75pt more, as Excel gives the 21pt title row above 21.75pt, 21 without the
    // attribute (measured over COM). thickTop isn't measured.)
    if (own?.thickBot) best += 0.75;
    out.set(r, Math.min(best, MAX_ROW_PT));
  }
  return out;
}
