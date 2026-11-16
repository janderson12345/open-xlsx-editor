// Sparklines as Excel keeps them: a worksheet extLst ext ({05C60535-...}) holding x14:sparklineGroups, each group with
// its type, options and colours and the sparklines that share them (each a source range, xm:f, drawn in one cell,
// xm:sqref). Measured with tools/excel-compare/excel-sparklines.ps1 (docs/SPARKLINES.md): Excel writes the type only
// when it isn't a line (column; win/loss is "stacked"), each option only when it's on or not its default, and every
// colour, its default style's in theme colours when inserted from the ribbon. Read here into sheet.sparklines; what
// each cell shows is worked out by sparklineLooks for the view (editor/sparkgeom.js, sparkview.js).
import { kids, kid, attr } from './xml.js';
import { readColor } from './styles.js';
import { resolveColor } from './theme.js';
import { parseRange, parseRef } from './refs.js';
import { X14, XM } from './cfext.js';
import { maxOf, minOf } from './listmath.js';

export const EXT_SPARKLINES = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}';
const COLORS = ['Series', 'Negative', 'Axis', 'Markers', 'First', 'Last', 'High', 'Low'];
const on = (el, name) => { const v = attr(el, name); return v === '1' || v === 'true'; };
const num = (el, name) => (attr(el, name) == null ? null : Number(attr(el, name)));

/** A reference as Excel writes it in a sparkline (Sheet1!B2:M2, 'My sheet'!$B$2:$M$2): { sheet, range } or null. */
export function parseSheetRange(f) {
  const m = f && /^(?:'((?:[^']|'')+)'|([^!]+))!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/.exec(f.trim());
  if (!m) return null;
  const range = parseRange(m[3].replace(/\$/g, ''));
  return range && { sheet: (m[1] ?? m[2]).replace(/''/g, "'"), range };
}

/**
 * The worksheet's sparkline groups: [{ type ('line', 'column' or 'stacked': win/loss), lineWeight (points, 0.75 by
 * default), markers, high, low, first, last, negative, axis (displayXAxis), empty ('gap', 'zero' or 'span'; the
 * schema's default is zero, which Excel then leaves unwritten), hidden (displayHidden), rightToLeft, dates (the date
 * axis's range, or null), minType and maxType ('individual', 'group' or 'custom'), min and max (manualMin, manualMax),
 * colors: { series, negative, axis, markers, first, last, high, low } (styles.js readColor), sparklines: [{ source
 * (its reference), r, c (the cell it's drawn in) }] }].
 */
export function readSparklines(root, ns) {
  return groupElements(root, ns).map(readGroup);
}

/** The worksheet's x14:sparklineGroup elements, in the file's order. */
export function groupElements(root, ns) {
  const ext = kids(kid(root, 'extLst', ns), 'ext', ns).find((e) => attr(e, 'uri') === EXT_SPARKLINES);
  return kids(kid(ext, 'sparklineGroups', X14), 'sparklineGroup', X14);
}

/** One x14:sparklineGroup, as readSparklines gives it. */
export function readGroup(g) {
  return {
    type: attr(g, 'type') || 'line',
    lineWeight: num(g, 'lineWeight') ?? 0.75,
    markers: on(g, 'markers'), high: on(g, 'high'), low: on(g, 'low'), first: on(g, 'first'), last: on(g, 'last'), negative: on(g, 'negative'),
    axis: on(g, 'displayXAxis'),
    empty: attr(g, 'displayEmptyCellsAs') || 'zero',
    hidden: on(g, 'displayHidden'),
    rightToLeft: on(g, 'rightToLeft'),
    dates: on(g, 'dateAxis') ? kid(g, 'f', XM)?.textContent || null : null,
    minType: attr(g, 'minAxisType') || 'individual', maxType: attr(g, 'maxAxisType') || 'individual',
    min: num(g, 'manualMin') ?? 0, max: num(g, 'manualMax') ?? 0,
    colors: Object.fromEntries(COLORS.map((n) => [n.toLowerCase(), readColor(kid(g, `color${n}`, X14))])),
    sparklines: kids(kid(g, 'sparklines', X14), 'sparkline', X14).map((s) => {
      const at = parseRef((kid(s, 'sqref', XM)?.textContent || '').trim());
      return at && { source: kid(s, 'f', XM)?.textContent || '', r: at.r, c: at.c };
    }).filter(Boolean),
  };
}

/** Whether row r or column c of a sheet is hidden. */
const hiddenAt = (sheet, r, c) => !!sheet.rows.get(r)?.hidden || sheet.cols.some((k) => k.hidden && c >= k.min && c <= k.max);

/** A range's cells in order (along its row or down its column), each { r, c, v } with v a number, null (empty) or
 *  0 for text: Excel plots a text cell as 0 (sparklines.xlsx, "Text in the data": its points on the zero line). */
function cellsOf(wb, ref) {
  const p = parseSheetRange(ref);
  const si = p ? wb.sheetIndexOf(p.sheet) : -1;
  if (si < 0 || wb.sheets[si].kind !== 'worksheet') return null;
  const sheet = wb.sheet(si);
  const out = [];
  for (let r = p.range.r0; r <= p.range.r1; r++) {
    for (let c = p.range.c0; c <= p.range.c1; c++) {
      const v = sheet.cells.get(r)?.get(c)?.v;
      out.push({ r, c, hidden: hiddenAt(sheet, r, c), v: typeof v === 'number' ? v : typeof v === 'string' && v !== '' ? 0 : null });
    }
  }
  return out;
}

/**
 * What each sparkline of sheet si shows, by "r,c": { type, values ([number or null], in the source's order, cells in
 * hidden rows and columns left out unless the group shows them, empty cells 0 when shown as zero), dates (each
 * value's date for a date axis, else null), lo, hi (the vertical axis: the sparkline's own least and greatest, the
 * group's, or the group's custom ones), empty, rightToLeft, lineWeight, the point options and colours as '#RRGGBB' }.
 */
export function sparklineLooks(wb, si) {
  const out = new Map();
  const groups = wb.sheet(si).sparklines || [];
  const color = (c) => resolveColor(c, wb.theme, wb.styles.palette);
  for (const g of groups) {
    const dates = g.dates ? cellsOf(wb, g.dates) : null;
    const lines = [];
    for (const s of g.sparklines) {
      const cells = cellsOf(wb, s.source);
      if (!cells) continue;
      const values = [];
      const when = [];
      cells.forEach((cell, k) => {
        if (cell.hidden && !g.hidden) return;
        // (A date axis plots each value at its date; one without a date isn't plotted.)
        if (dates && dates[k]?.v == null) return;
        values.push(cell.v == null && g.empty === 'zero' ? 0 : cell.v);
        if (dates) when.push(dates[k].v);
      });
      const shown = values.filter((v) => v != null);
      lines.push({ s, values, dates: dates ? when : null, lo: shown.length ? minOf(shown) : 0, hi: shown.length ? maxOf(shown) : 0 });
    }
    const groupLo = minOf(lines.map((l) => l.lo));
    const groupHi = maxOf(lines.map((l) => l.hi));
    const colors = Object.fromEntries(Object.entries(g.colors).map(([k, c]) => [k, color(c)]));
    for (const l of lines) {
      const lo = g.minType === 'custom' ? g.min : g.minType === 'group' ? groupLo : l.lo;
      const hi = g.maxType === 'custom' ? g.max : g.maxType === 'group' ? groupHi : l.hi;
      out.set(`${l.s.r},${l.s.c}`, {
        type: g.type, values: l.values, dates: l.dates, lo, hi, empty: g.empty, rightToLeft: g.rightToLeft, lineWeight: g.lineWeight,
        markers: g.markers, high: g.high, low: g.low, first: g.first, last: g.last, negative: g.negative, axis: g.axis, colors,
      });
    }
  }
  return out;
}
