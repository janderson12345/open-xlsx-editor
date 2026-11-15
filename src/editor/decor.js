// Cell fills and borders: a rectangle under the text for each filled cell in view, an edge over it for each side with a
// border, for empty cells as much as for cells with values. Their colours and line styles come from cellstyles.js's
// classes; their geometry from here, on screen in whole pixels on the gridlines, in print as Excel's PDF has them.
import { mergeAt } from './cells.js';
import { setVars, div } from './dom.js';
import { lineStyle } from '../ooxml/sheet.js';
import { LINE } from './cellstyles.js';
import { sparkElement } from './sparkview.js';

/** Border line widths in units of a thin line: thin and its dash patterns 1, medium and its patterns 2, thick and
 *  double 3; hair is thinner than a unit. */
const UNITS = { thin: 1, dotted: 1, dashed: 1, dashDot: 1, dashDotDot: 1, hair: 1, medium: 2, mediumDashed: 2, mediumDashDot: 2,
  mediumDashDotDot: 2, slantDashDot: 2, thick: 3, double: 3 };
/** Printed widths: a unit is 1.28px (8 device pixels of a 600 dpi printer) and hair 0.16px (1 device pixel), centred on
 *  the cell boundary (Excel's PDF of test/fixtures/fills.xlsx: thin 1.28, medium 2.56, thick and double 3.84, hair
 *  0.16). On screen a unit is a pixel. */
const PRINT_UNIT = 1.28;
const PRINT_HAIR = 0.16;
/** Where Excel's printed grid starts beyond margin + column widths and row heights: half a point, across and down
 *  (fills.xlsx: every fill starts 0.67px right and 0.7px below that, every border is centred there). Text, fitted
 *  before this was known, carries it in its own padding and baseline. */
export const PRINT_GRID_ORIGIN = { x: 0.67, y: 0.7 };

/** A border side's width in pixels, on screen or in print. */
export function edgeWidth(style, print) {
  if (!style) return 0;
  if (print) return style === 'hair' ? PRINT_HAIR : (UNITS[style] || 1) * PRINT_UNIT;
  return UNITS[style] || 1;
}

/** Where a line of width w on boundary b starts: on screen on the gridline pixel (b - 1) and centred on it for wider
 *  lines; in print centred on Excel's boundary. */
const lineStart = (b, w, print, origin) => (print ? b + origin - w / 2 : b - 1 - Math.floor((w - 1) / 2));

/** Fill rectangles and border edges for the cells of `win` ({ r0, r1, c0, c1 }), as elements: { fills, edges }.
 *  `st` carries the sheet, its axes, `deco` (cellstyles.js: per cell format, its fill and its sides) and `print`. */
export function decorations(st, win) {
  const fills = [];
  const edges = [];
  const { rows, cols, sheet, deco, print } = st;
  if (!deco) return { fills, edges };
  const ox = print ? PRINT_GRID_ORIGIN.x : 0;
  const oy = print ? PRINT_GRID_ORIGIN.y : 0;
  const drawnMerges = new Set();
  const sparks = [];
  for (let r = win.r0; r <= win.r1; r++) {
    if (!rows.size(r)) continue;
    const line = sheet.cells.get(r);
    for (let c = win.c0; c <= win.c1; c++) {
      if (!cols.size(c)) continue;
      const merge = mergeAt(st, r, c);
      if (merge && !drawnMerges.has(merge)) {
        // A merged range is filled as its top-left cell is, or with the sheet's own colour, which hides the gridlines
        // inside it (on screen; print has no gridlines).
        drawnMerges.add(merge);
        const d = deco(sheet.cells.get(merge.r0)?.get(merge.c0)?.s || 0);
        if (d.fill || !print) fills.push(rect(st, merge.r0, merge.c0, merge.r1, merge.c1, d.fill ? d.fillClass : 'oxe-blank', ox, oy));
      }
      // (A sparkline, over the cell's fills: ooxml/sparklines.js, sparkview.js.)
      const sp = st.sparks?.get(`${r},${c}`);
      if (sp && (!merge || (merge.r0 === r && merge.c0 === c))) sparks.push(sparkElement(st, merge || { r0: r, c0: c, r1: r, c1: c }, sp, ox, oy));
      // (A conditional format's fill, over the cell's own: ooxml/condformat.js.)
      const cf = st.cond?.get(`${r},${c}`);
      if (cf?.fill) {
        const m = merge || { r0: r, c0: c, r1: r, c1: c };
        if (m.r0 === r && m.c0 === c) { const el = rect(st, m.r0, m.c0, m.r1, m.c1, 'oxe-cf-fill', ox, oy); el.style.setProperty('--oxe-cf', cf.fill); fills.push(el); }
      }
      // (A data bar: from its start to its end across the cell, a pixel in from the gridlines, and the axis dashed.)
      if (cf?.bar) fills.push(...dataBar(st, r, c, cf.bar, ox, oy));
      const cell = line?.get(c);
      // (An empty cell: its row's or column's own format, ooxml/sheet.js lineStyle.)
      const s = cell ? cell.s : lineStyle(sheet, r, c);
      if (!s) continue;
      const d = deco(s);
      if (d.fill && !merge && !cf?.fill) fills.push(rect(st, r, c, r, c, d.fillClass, ox, oy));
      // (A conditional format's border sides over the cell's own: ooxml/condformat.js.)
      for (const side of d.sides) if (!cf?.border?.[side.name]) edges.push(...edge(st, r, c, side, ox, oy));
      if (cf?.border) {
        for (const [name, b] of Object.entries(cf.border)) {
          for (const el of edge(st, r, c, { name, style: b.style, cls: 'oxe-cf-edge' }, ox, oy)) { el.style.setProperty('--oxe-es', LINE[b.style] || 'solid'); el.style.setProperty('--oxe-ec', b.color); edges.push(el); }
        }
      }
      if (d.diag) {
        const m = merge || { r0: r, c0: c, r1: r, c1: c };
        if (m.r0 === r && m.c0 === c) for (const dir of ['up', 'down']) if (d.diag[dir]) {
          const el = rect(st, m.r0, m.c0, m.r1, m.c1, `oxe-diag oxe-diag-${dir} ${d.diag.cls}`, ox, oy);
          el.classList.remove('oxe-fill');
          el.style.setProperty('--oxe-bw', `${edgeWidth(d.diag.style, st.print)}px`);
          edges.push(el);
        }
      }
    }
  }
  fills.push(...sparks);
  return { fills, edges };
}

/** A data bar's elements in cell (r, c): the bar (gradient: its colour fading to white towards its end, as Excel draws
 *  it; a border in its border colour) and, when it has one, the axis. */
function dataBar(st, r, c, bar, ox, oy) {
  const x = st.cols.start(c) + ox;
  const w = st.cols.size(c);
  const y = st.rows.start(r) + oy;
  const h = st.rows.size(r);
  if (w < 4 || h < 4) return [];
  const inner = w - 3;
  const from = x + 1 + bar.from * inner;
  const width = Math.max(0, (bar.to - bar.from) * inner);
  const el = div('oxe-fill oxe-cf-bar');
  setVars(el, { '--oxe-x': from, '--oxe-y': y + 1, '--oxe-w': width, '--oxe-h': h - 3 });
  // (Pointing away from the axis: a negative bar's gradient fades leftwards.)
  el.style.background = bar.gradient ? `linear-gradient(to ${bar.negative ? 'left' : 'right'}, ${bar.color}, #ffffff)` : bar.color;
  if (bar.border) el.style.boxShadow = `inset 0 0 0 1px ${bar.border}`;
  const out = [el];
  if (bar.axis != null) {
    const axis = div('oxe-fill oxe-cf-axis');
    setVars(axis, { '--oxe-x': x + 1 + bar.axis * inner, '--oxe-y': y + 1, '--oxe-w': 1, '--oxe-h': h - 3 });
    axis.style.borderLeft = `1px dashed ${bar.axisColor}`;
    out.push(axis);
  }
  return out;
}

/** A fill from the start of (r0, c0) to the end of (r1, c1); on screen it covers the gridline pixels inside. */
function rect(st, r0, c0, r1, c1, cls, ox, oy) {
  const el = div(`oxe-fill ${cls}`);
  setVars(el, { '--oxe-x': st.cols.start(c0) + ox, '--oxe-y': st.rows.start(r0) + oy, '--oxe-w': st.cols.start(c1 + 1) - st.cols.start(c0), '--oxe-h': st.rows.start(r1 + 1) - st.rows.start(r0) });
  return el;
}

/** Excel's printed dash patterns, in pixels: [on, off, on, off, ...] and where along the page an "on" run starts
 *  (its phase, modulo the period). Excel's PDF of fills.xlsx: dashed 3.84 on 1.28 off; medium dashed 11.52 on 3.84
 *  off; dash-dot 11.52, 3.84, 3.84, 3.84; dotted and hair print solid. The phases are where the dashes fell in that
 *  PDF, anchored to the page (horizontal and vertical edges agree within 0.1 to 0.4px); the patterns not in the probe
 *  are guesses of the same shape. */
const PRINT_DASH = {
  dashed: { runs: [3.84, 1.28], phase: 4.3 },
  mediumDashed: { runs: [11.52, 3.84], phase: 3.8 },
  dashDot: { runs: [11.52, 3.84, 3.84, 3.84], phase: 18.94 },
  mediumDashDot: { runs: [11.52, 3.84, 3.84, 3.84], phase: 18.94 },
  slantDashDot: { runs: [11.52, 3.84, 3.84, 3.84], phase: 18.94 },
  dashDotDot: { runs: [11.52, 3.84, 3.84, 3.84, 3.84, 3.84], phase: 0 },
  mediumDashDotDot: { runs: [11.52, 3.84, 3.84, 3.84, 3.84, 3.84], phase: 0 },
};

/** One side's edge, the length of the cell, with the corners closed: one element, or in print the dashes of a dashed
 *  style as elements of their own. */
function edge(st, r, c, side, ox, oy) {
  const w = edgeWidth(side.style, st.print);
  const vertical = side.name === 'left' || side.name === 'right';
  const x0 = st.cols.start(c);
  const x1 = st.cols.start(c + 1);
  const y0 = st.rows.start(r);
  const y1 = st.rows.start(r + 1);
  // In print the corners are shared as Excel's PDF shares them (fills.xlsx): the left edge runs the cell's full
  // height, the right edge starts below the top boundary's line, and the top and bottom edges start after the left
  // boundary's line and run past the right one. On screen every edge covers its gridline pixel at both ends.
  let x, y, len;
  if (vertical) {
    x = lineStart(side.name === 'left' ? x0 : x1, w, st.print, ox);
    y = st.print ? y0 + oy + (side.name === 'left' ? -w / 2 : w / 2) : y0 - 1;
    len = (st.print ? y1 + oy + w / 2 : y1) - y;
  } else {
    y = lineStart(side.name === 'top' ? y0 : y1, w, st.print, oy);
    x = st.print ? x0 + ox + w / 2 : x0 - 1;
    len = (st.print ? x1 + ox + w / 2 : x1) - x;
  }
  const piece = (at, size) => {
    const el = div(`oxe-edge ${vertical ? 'oxe-edge-v' : 'oxe-edge-h'} ${side.cls}${st.print && side.style !== 'double' ? ' oxe-edge-solid' : ''}`);
    setVars(el, vertical ? { '--oxe-x': x, '--oxe-y': at, '--oxe-h': size, '--oxe-bw': w } : { '--oxe-x': at, '--oxe-y': y, '--oxe-w': size, '--oxe-bw': w });
    return el;
  };
  const start = vertical ? y : x;
  const dash = st.print && PRINT_DASH[side.style];
  if (!dash) return [piece(start, len)];
  // (Dashes are laid along the page: `st.pageOrigin` is where the sheet's 0 falls on the page.)
  const origin = vertical ? st.pageOrigin.y : st.pageOrigin.x;
  const period = dash.runs.reduce((a, b) => a + b, 0);
  const pieces = [];
  let at = start - ((((start + origin - dash.phase) % period) + period) % period);
  for (let k = 0; at < start + len; k = (k + 1) % dash.runs.length) {
    const run = dash.runs[k];
    if (k % 2 === 0) {
      const a = Math.max(at, start);
      const b = Math.min(at + run, start + len);
      if (b > a) pieces.push(piece(a, b - a));
    }
    at += run;
  }
  return pieces;
}
