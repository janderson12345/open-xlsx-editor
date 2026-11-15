// A sparkline's shapes inside its cell, in pixels from the cell's top left, as Excel draws them. No DOM: the grid,
// print and PDF draw these (sparkview.js), and test/sparklines.test.js compares them with Excel's PDF.
//
// Measured on Excel's PDFs (tools/excel-compare/excel-sparklines.ps1, docs/SPARKLINES.md), where each sparkline is a
// 200 dpi picture of its cell: the plot sits 4 picture pixels (1.92px) in from the cell's sides, 8 (3.84px) below its
// top and 8.5 (4.08px) above its bottom, in cells 8 to 45 characters wide and 12 to 96 points tall. Its width is cut
// into one equal slot per value (with a date axis, one per day from the first date to the last); a line's points are
// at the slots' middles, a column fills the middle 1/1.4 of its slot.
const PAD_X = 1.92;
const PAD_TOP = 3.84;
const PAD_BOTTOM = 4.08;
/** A column's width in its slot (4 values: bars 60.64 of 84.96px slots; 12 values: 28 of 39.36 picture pixels). */
const BAR = 1 / 1.4;
/** A bar at the axis's own value is drawn a picture pixel tall (the least value of a column sparkline over positive
 *  numbers, sparklines.xlsx "Column, points"). */
const HAIRLINE = 0.48;
/** Markers are diamonds; their width across by line weight in points (Excel's PDFs, sheet Weights: 0.25 and 0.75pt
 *  3.36px, 1.5pt 4.8, 2.25pt 7.68, 3pt 9.12, 6pt 12.96), straight between, on beyond 6pt as from 3 to 6. */
const MARKER = [[0, 3.36], [0.75, 3.36], [1.5, 4.8], [2.25, 7.68], [3, 9.12], [6, 12.96]];

export function markerSize(weight) {
  for (let k = 1; k < MARKER.length; k++) {
    const [w0, s0] = MARKER[k - 1];
    const [w1, s1] = MARKER[k];
    if (weight <= w1) return s0 + ((weight - w0) / (w1 - w0)) * (s1 - s0);
  }
  return 12.96 + (weight - 6) * 1.28;
}

/** Which colour a point takes: the series' (or for a line, the markers'), then negative, first, last, high and low,
 *  each over the ones before (sparklines.xlsx: a first point that is also the high one shows the high colour, a last
 *  that is the low one the low colour, a negative low point the low colour; every point at the high or the low). */
function pointColor(sp, v, i, ends, base) {
  const { colors } = sp;
  let color = base;
  if (sp.negative && v < 0) color = colors.negative;
  if (sp.first && i === ends.first) color = colors.first;
  if (sp.last && i === ends.last) color = colors.last;
  if (sp.high && v === ends.high) color = colors.high;
  if (sp.low && v === ends.low) color = colors.low;
  return color;
}

/**
 * The shapes of sparkline `sp` (ooxml/sparklines.js sparklineLooks) in a cell w by h: { bars: [{ x, y, w, h, color }],
 * line: { color, width, runs: [[[x, y], ...], ...] } or null, markers: [{ x, y, size, color }], axis: { x0, x1, y,
 * color } or null }.
 */
export function sparkShapes(sp, w, h) {
  const out = { bars: [], line: null, markers: [], axis: null };
  const n = sp.values.length;
  const top = PAD_TOP;
  const bottom = h - PAD_BOTTOM;
  const left = PAD_X;
  const width = w - 2 * PAD_X;
  if (!n || width <= 0 || bottom <= top) return out;
  // Across: slots by index, or by day on a date axis; right to left mirrors them.
  let slot = width / n;
  let at = (i) => left + (i + 0.5) * slot;
  if (sp.dates?.length) {
    const d0 = Math.min(...sp.dates);
    slot = width / (Math.max(...sp.dates) - d0 + 1);
    at = (i) => left + (sp.dates[i] - d0 + 0.5) * slot;
  }
  const xOf = sp.rightToLeft ? (i) => 2 * left + width - at(i) : at;
  // Down: hi at the plot's top, lo at its bottom; all one value on the top line (sparklines.xlsx "All the same").
  const { lo, hi } = sp;
  const yOf = (v) => (hi === lo ? top : top + ((hi - v) / (hi - lo)) * (bottom - top));
  const clampY = (y) => Math.min(bottom, Math.max(top, y));
  const shown = sp.values.map((v, i) => [v, i]).filter(([v]) => v != null);
  const ends = { first: shown[0]?.[1], last: shown.at(-1)?.[1], high: Math.max(...shown.map(([v]) => v)), low: Math.min(...shown.map(([v]) => v)) };
  if (sp.type === 'line') {
    // Empty cells break the line (gap), or are passed over (span); a point on its own draws nothing.
    const runs = [];
    let run = [];
    sp.values.forEach((v, i) => {
      if (v == null) { if (sp.empty !== 'span' && run.length) { runs.push(run); run = []; } return; }
      run.push([xOf(i), yOf(v)]);
    });
    if (run.length) runs.push(run);
    out.line = { color: sp.colors.series, width: (sp.lineWeight * 4) / 3, runs: runs.filter((r) => r.length > 1) };
    // Markers: every point with markers on, else only the points an option colours.
    const size = markerSize(sp.lineWeight);
    for (const [v, i] of shown) {
      const color = pointColor(sp, v, i, ends, sp.markers ? sp.colors.markers : null);
      if (color) out.markers.push({ x: xOf(i), y: yOf(v), size, color });
    }
  } else {
    const bw = slot * BAR;
    // (A column runs from the axis's zero, or from the end of the axis nearer zero when it doesn't cross it: from the
    // least value over positive numbers, the greatest over negative ones (sparklines.xlsx "Column, negative"); bars
    // past a custom end stop there ("Custom max 8"). Win/loss: up from the middle for a win, down for a loss, nothing
    // for 0.)
    const mid = (top + bottom) / 2;
    const base = sp.type === 'stacked' ? mid : clampY(yOf(Math.min(hi, Math.max(lo, 0))));
    for (const [v, i] of shown) {
      if (sp.type === 'stacked' && v === 0) continue;
      const end = sp.type === 'stacked' ? (v > 0 ? top : bottom) : clampY(yOf(v));
      let y0 = Math.min(base, end);
      let y1 = Math.max(base, end);
      // (On the axis: a hairline on the plot's side of it.)
      if (y1 - y0 < HAIRLINE) [y0, y1] = base > mid ? [base - HAIRLINE, base] : [base, base + HAIRLINE];
      out.bars.push({ x: xOf(i) - bw / 2, y: y0, w: bw, h: y1 - y0, color: pointColor(sp, v, i, ends, sp.colors.series) });
    }
  }
  // The horizontal axis, when asked for and zero is on the vertical one (none over positive numbers only,
  // sparklines.xlsx "Line, axis, positive"), across the plot.
  if (sp.axis && (sp.type === 'stacked' || (lo <= 0 && hi >= 0 && hi > lo))) {
    out.axis = { x0: left, x1: left + width, y: sp.type === 'stacked' ? (top + bottom) / 2 : yOf(0), color: sp.colors.axis };
  }
  return out;
}

const r2 = (n) => Math.round(n * 100) / 100;

/** The shapes as SVG paths in the order they're painted (bars, the line, the axis over it, the markers): [{ d, fill }
 *  or { d, stroke, width }], for sparkview.js and the PDF export. */
export function sparkPaths(s) {
  const out = [];
  for (const b of s.bars) out.push({ d: `M${r2(b.x)} ${r2(b.y)}h${r2(b.w)}v${r2(b.h)}h${r2(-b.w)}z`, fill: b.color });
  if (s.line?.runs.length) out.push({ d: s.line.runs.map((run) => run.map(([x, y], k) => `${k ? 'L' : 'M'}${r2(x)} ${r2(y)}`).join('')).join(''), stroke: s.line.color, width: r2(s.line.width) });
  if (s.axis) out.push({ d: `M${r2(s.axis.x0)} ${r2(s.axis.y)}H${r2(s.axis.x1)}`, stroke: s.axis.color, width: 1 });
  for (const m of s.markers) {
    const k = m.size / 2;
    out.push({ d: `M${r2(m.x)} ${r2(m.y - k)}l${r2(k)} ${r2(k)}l${r2(-k)} ${r2(k)}l${r2(-k)} ${r2(-k)}z`, fill: m.color });
  }
  return out;
}
