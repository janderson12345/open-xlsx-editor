// What a sheet prints, from the workbook's defined names (Excel's Page Setup > Sheet): its print area
// (_xlnm.Print_Area, the used range without one) and the rows and columns printed at the top and left of every page
// (_xlnm.Print_Titles: "Sheet1!$1:$2", "Sheet1!$A:$A", or both, comma-separated). A print area of several ranges is
// printed by its first for now.
import { parseRange, colIndex, MAX_ROWS, MAX_COLS } from './refs.js';

/** The ranges a defined name's text holds on its sheet ("Probe!$B$3:$H$70,Probe!$1:$2"): [{ r0, c0, r1, c1 }], whole
 *  rows and columns as whole lines. */
export function rangesOf(text) {
  const out = [];
  for (const part of String(text || '').split(',')) {
    const ref = part.trim().replace(/^(?:'(?:[^']|'')+'|[^!]+)!/, '').replace(/\$/g, '');
    let m = /^(\d+):(\d+)$/.exec(ref);
    if (m) { out.push({ r0: Number(m[1]) - 1, r1: Number(m[2]) - 1, c0: 0, c1: MAX_COLS - 1, rows: true }); continue; }
    m = /^([A-Z]+):([A-Z]+)$/i.exec(ref);
    if (m) { out.push({ r0: 0, r1: MAX_ROWS - 1, c0: colIndex(m[1].toUpperCase()), c1: colIndex(m[2].toUpperCase()), cols: true }); continue; }
    const a = parseRange(ref);
    if (a) out.push(a);
  }
  return out;
}

/** Sheet si's print area ({ r0, c0, r1, c1 } or null: its used range) and print titles ({ rows: [r0, r1] or null,
 *  cols: [c0, c1] or null }). */
export function printSetup(wb, si) {
  const name = (n) => wb.definedNames.find((d) => d.name.toLowerCase() === n && d.sheet === si);
  const area = rangesOf(name('_xlnm.print_area')?.text)[0] || null;
  const titles = { rows: null, cols: null };
  for (const r of rangesOf(name('_xlnm.print_titles')?.text)) {
    if (r.rows) titles.rows = [r.r0, r.r1];
    if (r.cols) titles.cols = [r.c0, r.c1];
  }
  return { area, titles };
}
