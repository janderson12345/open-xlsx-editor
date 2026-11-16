// Insert > Charts: a chart of the selection made as Excel makes one (docs/CHARTS.md). The series come from the cells
// as Excel takes them (measured over COM on ten shapes of selection): the first row holds the series' names (or the
// categories) when none of its cells but the corner is a number, the first column likewise; a blank corner makes both
// labels. Series run down the columns when the numbers have more rows than columns, else along the rows. The chart
// part (chartxml.js), a drawing (the sheet's own, or a new one) with the chart anchored beside the data, their
// relationships and content types, and the sheet's <drawing> are written as Excel writes them.
import { chartXml, CHART_KINDS } from './chartxml.js';
import { sheetPrefix } from '../calc/tokens.js';
import { colName } from './refs.js';
import { formatCodeOf } from './styles.js';
import { colWidthToPx, defaultColPx, rowHeightToPx, EMU_PER_PX } from './units.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml';
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const XDR = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
/** Sheet elements after <drawing> (CT_Worksheet's order). */
const AFTER_DRAWING = ['legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
/** A new chart's size: Excel's Insert > Charts gives 5 by 3 inches. */
export const CHART_SIZE = { w: 480, h: 288 };

const MDW = 7;
const cellAt = (sheet, r, c) => sheet.cells.get(r)?.get(c) || null;
const isNumber = (cell) => cell?.t === 'n' && cell.v != null;
const shown = (cell) => (cell?.v == null ? null : cell.t === 'b' ? (cell.v ? 'TRUE' : 'FALSE') : cell.v);

/** A range's reference as a series formula has it: Sheet1!$A$2:$A$5. */
const rangeRef = (name, r0, c0, r1, c1) => `${sheetPrefix(name)}$${colName(c0)}$${r0 + 1}${r0 === r1 && c0 === c1 ? '' : `:$${colName(c1)}$${r1 + 1}`}`;

/**
 * The series a chart of `sel` ({ r0, c0, r1, c1 }) on sheet si gets, as Excel takes them: { byRows, series: [{ name,
 * cat, val }] }, each part { f, values, numeric, format } or null. Null when the range has no numbers to chart.
 * `byRows` true or false takes the series along the rows or down the columns (Excel's Select Data keeps a chart's own
 * way and Switch Row/Column turns it: excel-chart-data.ps1); null, Excel's rule for a new chart.
 */
export function chartSeries(wb, si, sel, byRows = null) {
  const sheet = wb.sheet(si);
  const name = wb.sheets[si].name;
  const { r0, c0, r1, c1 } = sel;
  const corner = cellAt(sheet, r0, c0);
  const many = r1 > r0 && c1 > c0;
  const noNumbers = (cells) => cells.every((cell) => !isNumber(cell));
  const row = [];
  for (let c = c0 + 1; c <= c1; c++) row.push(cellAt(sheet, r0, c));
  const col = [];
  for (let r = r0 + 1; r <= r1; r++) col.push(cellAt(sheet, r, c0));
  const blankCorner = corner?.v == null || corner.v === '';
  // (One column or row: its first cell a name when it isn't a number.)
  const rowLabels = r1 > r0 && ((many && blankCorner) || (noNumbers(c1 > c0 ? row : []) && !isNumber(corner)));
  const colLabels = c1 > c0 && ((many && blankCorner) || (r1 > r0 && noNumbers(col) && !isNumber(cellAt(sheet, r0, c0))));
  const dr0 = r0 + (rowLabels ? 1 : 0);
  const dc0 = c0 + (colLabels ? 1 : 0);
  if (dr0 > r1 || dc0 > c1) return null;
  byRows ??= !(r1 - dr0 > c1 - dc0);
  const data = (a0, b0, a1, b1, numeric) => {
    const values = [];
    for (let r = a0; r <= a1; r++) for (let c = b0; c <= b1; c++) values.push(shown(cellAt(sheet, r, c)));
    const first = (() => { for (let r = a0; r <= a1; r++) for (let c = b0; c <= b1; c++) if (cellAt(sheet, r, c)?.v != null) return cellAt(sheet, r, c); return null; })();
    const num = numeric ?? (first ? isNumber(first) : false);
    return { f: rangeRef(name, a0, b0, a1, b1), values: num ? values.map((v) => (typeof v === 'number' ? v : null)) : values.map((v) => (v == null ? null : String(v))), numeric: num, format: num && first ? formatCodeOf(wb.styles, first.s || 0) : 'General' };
  };
  const series = [];
  if (!byRows) {
    const cat = colLabels ? data(dr0, c0, r1, c0) : null;
    for (let c = dc0; c <= c1; c++) series.push({ name: rowLabels ? data(r0, c, r0, c, false) : null, cat, val: data(dr0, c, r1, c, true) });
  } else {
    const cat = rowLabels ? data(r0, dc0, r0, c1) : null;
    for (let r = dr0; r <= r1; r++) series.push({ name: colLabels ? data(r, c0, r, c0, false) : null, cat, val: data(r, dc0, r, c1, true) });
  }
  return { byRows, series };
}

/** A column's and a row's size in pixels on sheet `sheet` (as the grid has them). */
function sizes(sheet) {
  const colPx = (c) => {
    const col = sheet.cols.find((x) => c >= x.min && c <= x.max);
    if (col?.hidden) return 0;
    const mdw = sheet.format?.mdw ?? MDW;
    return col?.width != null ? colWidthToPx(col.width, mdw) : defaultColPx(sheet.format, mdw);
  };
  const rowPx = (r) => {
    const info = sheet.rows.get(r);
    if (info?.hidden) return 0;
    return rowHeightToPx(info?.ht ?? sheet.format.defaultRowHeight);
  };
  return { colPx, rowPx };
}

/** A two-cell anchor from cell (r, c) for w by h pixels: its end's cell and offset (EMU), as Excel works it out. */
export function anchorFor(sheet, r, c, w, h) {
  const { colPx, rowPx } = sizes(sheet);
  const walk = (start, size, px) => {
    let at = start;
    let left = px;
    for (let guard = 0; guard < 20000 && left >= size(at); guard++) { left -= size(at); at++; }
    return { line: at, off: Math.round(left * EMU_PER_PX) };
  };
  const to = { col: walk(c, colPx, w), row: walk(r, rowPx, h) };
  return { from: { col: c, colOff: 0, row: r, rowOff: 0 }, to: { col: to.col.line, colOff: to.col.off, row: to.row.line, rowOff: to.row.off } };
}

const hex = (n) => Array.from({ length: n }, () => Math.floor(Math.random() * 16).toString(16).toUpperCase()).join('');
/** Fresh ids, as Excel makes them: two axis ids and a series GUID tail. */
export const freshIds = () => ({ axes: [100000000 + Math.floor(Math.random() * 2000000000), 100000000 + Math.floor(Math.random() * 2000000000)], guid: `${hex(4)}-${hex(4)}-${hex(4)}-${hex(12)}` });

const freePath = (pkg, base, ext) => { for (let n = 1; ; n++) if (!pkg.has(`${base}${n}.${ext}`)) return `${base}${n}.${ext}`; };
const relOf = (pkg, path, type) => pkg.readRels(path).find((x) => x.type === type);

/** Sheet si's drawing part, made (empty) with its relationship and the sheet's <drawing> when it has none. */
function drawingOf(wb, si) {
  const pkg = wb.pkg;
  const info = wb.sheets[si];
  const have = relOf(pkg, info.path, `${REL}/drawing`);
  if (have) return have.target;
  const path = freePath(pkg, 'xl/drawings/drawing', 'xml');
  pkg.setText(path, `${DECL}<xdr:wsDr ${XDR}></xdr:wsDr>`);
  pkg.addOverride(path, CT_DRAWING);
  const id = pkg.addRel(`${REL}/drawing`, `../${path.slice(3)}`, false, info.path);
  const { doc } = wb._part(si);
  const root = doc.documentElement;
  const el = doc.createElementNS(root.namespaceURI, root.prefix ? `${root.prefix}:drawing` : 'drawing');
  if (!root.getAttribute('xmlns:r')) root.setAttribute('xmlns:r', REL);
  el.setAttributeNS(REL, 'r:id', id);
  const before = [...root.childNodes].find((n) => n.nodeType === 1 && AFTER_DRAWING.includes(n.localName));
  root.insertBefore(el, before || null);
  wb.dirty.add(info.path);
  return path;
}

/**
 * Insert a chart of `kind` (chartxml.js CHART_KINDS) for the selection `sel` of sheet si, anchored with its top-left
 * at cell `at` ({ r, c }: by default the second row of the selection, two columns past it) and CHART_SIZE big.
 * Returns { chart: its part's path, drawing, name } or null when the selection has nothing to chart.
 */
export function insertChart(wb, si, sel, kind, { at = null, ids = freshIds(), size = CHART_SIZE } = {}) {
  if (!CHART_KINDS[kind]) throw new Error(`No chart kind ${kind}`);
  const found = chartSeries(wb, si, sel);
  if (!found || !found.series.some((s) => s.val.values.some((v) => v != null))) return null;
  const pkg = wb.pkg;
  const chart = freePath(pkg, 'xl/charts/chart', 'xml');
  pkg.setText(chart, chartXml(kind, found.series, { date1904: wb.date1904, ids }));
  pkg.addOverride(chart, CT_CHART);
  const drawing = drawingOf(wb, si);
  const rid = pkg.addRel(`${REL}/chart`, `../charts/${chart.slice('xl/charts/'.length)}`, false, drawing);
  const text = pkg.text(drawing);
  const used = [...text.matchAll(/<(?:\w+:)?cNvPr\b[^>]*\sid="(\d+)"/g)].map((m) => Number(m[1]));
  const id = Math.max(1, ...used) + 1;
  const name = `Chart ${id - 1}`;
  const pos = at || { r: sel.r0 + 1, c: sel.c1 + 2 };
  const a = anchorFor(wb.sheet(si), pos.r, pos.c, size.w, size.h);
  const point = (tag, p) => `<xdr:${tag}><xdr:col>${p.col}</xdr:col><xdr:colOff>${p.colOff}</xdr:colOff><xdr:row>${p.row}</xdr:row><xdr:rowOff>${p.rowOff}</xdr:rowOff></xdr:${tag}>`;
  const anchor = `<xdr:twoCellAnchor>${point('from', a.from)}${point('to', a.to)}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${REL}" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
  pkg.setText(drawing, text.replace(/<\/(\w+:)?wsDr>\s*$/, (end) => anchor + end));
  return { chart, drawing, name, anchor: a, series: found.series.length };
}
