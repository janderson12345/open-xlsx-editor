// A worksheet part (xl/worksheets/sheetN.xml) read into a sparse model: cells by row and column, column widths, row
// heights, merged ranges and the sheet's defaults. The part's XML document is kept with the model so that an edit
// changes only the elements it has to and everything not modelled (conditional formats, validations, extLst...) stays
// where it was.
import { kids, kid, attr } from './xml.js';
import { parseRef, parseRange } from './refs.js';
import { stringItemText } from './strings.js';
import { shiftFormula } from '../calc/tokens.js';
import { readColor } from './styles.js';
import { readExtRules, readDataBar, readIconSet, extIdOf } from './cfext.js';
import { readFilterColumn } from './filtermore.js';

/** A cell as stored: `t` its type (n number, s string, b boolean, e error, str formula text, d ISO date), `v` its
 *  value (null for a cell that only carries a style), `s` its style (an index into cellXfs), `f` its formula text.
 *  readSheet adds, where they apply: `shared` (the index of the shared formula the cell belongs to; `f` is then its
 *  own text, worked out from the group's first cell), `array` (an array formula's range, on its first cell) and
 *  `dataTable`. */
const makeCell = (t, v, s, f) => ({ t, v, s, f });

/** Read one `c` element's value, resolving shared strings. */
export function readCell(c, ns, sst) {
  const s = Number(attr(c, 's') || 0);
  const t = attr(c, 't') || 'n';
  const fEl = kid(c, 'f', ns);
  const f = fEl ? fEl.textContent : null;
  const vEl = kid(c, 'v', ns);
  const v = vEl ? vEl.textContent : null;
  switch (t) {
    case 's': return makeCell('s', v == null ? null : (sst[Number(v)] ?? ''), s, f);
    case 'inlineStr': return makeCell('s', stringItemText(kid(c, 'is', ns), ns), s, f);
    case 'b': return makeCell('b', v == null ? null : v === '1' || v === 'true', s, f);
    case 'e': case 'str': case 'd': return makeCell(t, v, s, f);
    default: return makeCell('n', v == null || v === '' ? null : Number(v), s, f);
  }
}

/** The sheet's model from its parsed XML and the workbook's shared strings. */
export function readSheet(doc, sst = []) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  /** @type {Map<number, Map<number, ReturnType<typeof makeCell>>>} */
  const cells = new Map();
  const rows = new Map(); // row -> { ht (points), hidden, customHeight }
  let maxRow = -1;
  let maxCol = -1;
  let r = -1;
  const shared = new Map(); // si -> the group's first cell { r, c, text }
  const arrays = [];
  for (const rowEl of kids(kid(root, 'sheetData', ns), 'row', ns)) {
    // (`r` may be left out: rows and cells then follow on from the one before.)
    r = attr(rowEl, 'r') ? Number(attr(rowEl, 'r')) - 1 : r + 1;
    const ht = attr(rowEl, 'ht');
    const hidden = attr(rowEl, 'hidden') === '1' || attr(rowEl, 'hidden') === 'true';
    if (ht != null || hidden) rows.set(r, { ht: ht == null ? null : Number(ht), hidden, customHeight: attr(rowEl, 'customHeight') === '1' });
    let c = -1;
    let line = null;
    for (const cEl of kids(rowEl, 'c', ns)) {
      const at = parseRef(attr(cEl, 'r'));
      c = at ? at.c : c + 1;
      if (!line) cells.set(r, (line = new Map()));
      const cell = readCell(cEl, ns, sst);
      if (cell.f != null) formulaKind(cell, kid(cEl, 'f', ns), r, c, shared, arrays, attr(cEl, 'cm') != null);
      line.set(c, cell);
      if (c > maxCol) maxCol = c;
    }
    if (line && r > maxRow) maxRow = r;
  }
  const cols = [];
  for (const colsEl of kids(root, 'cols', ns)) {
    for (const col of kids(colsEl, 'col', ns)) {
      cols.push({
        min: Number(attr(col, 'min')) - 1,
        max: Number(attr(col, 'max')) - 1,
        width: attr(col, 'width') == null ? null : Number(attr(col, 'width')),
        hidden: attr(col, 'hidden') === '1' || attr(col, 'hidden') === 'true',
        style: Number(attr(col, 'style') || 0),
      });
    }
  }
  const merges = kids(kid(root, 'mergeCells', ns), 'mergeCell', ns).map((m) => parseRange(attr(m, 'ref'))).filter(Boolean);
  for (const m of merges) {
    if (m.r1 > maxRow) maxRow = m.r1;
    if (m.c1 > maxCol) maxCol = m.c1;
  }
  const fp = kid(root, 'sheetFormatPr', ns);
  const num = (name) => (fp && attr(fp, name) != null ? Number(attr(fp, name)) : null);
  const view = kid(kid(root, 'sheetViews', ns), 'sheetView', ns);
  return {
    cells,
    rows,
    cols,
    merges,
    /** The ranges of array formulas (their cells can't be changed one at a time). */
    arrays,
    format: { defaultRowHeight: num('defaultRowHeight') ?? 15, defaultColWidth: num('defaultColWidth'), baseColWidth: num('baseColWidth') },
    showGridLines: !view || attr(view, 'showGridLines') !== '0',
    showHeadings: !view || attr(view, 'showRowColHeaders') !== '0',
    ...readView(view, ns),
    page: readPageSetup(root, ns),
    autoFilter: readAutoFilter(kid(root, 'autoFilter', ns), ns),
    condFormats: readCondFormats(root, ns),
    validations: readValidations(root, ns),
    maxRow,
    maxCol,
  };
}

/** The sheet view's frozen rows and columns (a frozen pane's ySplit and xSplit; a split that isn't frozen is
 *  ignored) and its zoom in percent. */
/**
 * The sheet's AutoFilter: { r0, c0, r1, c1 (its range, the header row first), columns: Map(column index => { values:
 * [shown texts], blank (blanks shown too), other (a kind of filter not read here: custom, top 10, colour, dates) }) },
 * or null.
 */
function readAutoFilter(el, ns) {
  const area = el && parseRange(attr(el, 'ref'));
  if (!area) return null;
  const columns = new Map();
  // (Each column's filter: a list of values, or filtermore.js's text, number, top 10, average, date or colour filter.)
  for (const fc of kids(el, 'filterColumn', ns)) columns.set(area.c0 + Number(attr(fc, 'colId')), readFilterColumn(fc, ns));
  return { ...area, columns };
}

/**
 * The sheet's conditional formats: [{ ranges: [{ r0, c0, r1, c1 }], rules: [{ type, priority, dxfId, stopIfTrue,
 * operator, formulas, text, rank, percent, bottom, aboveAverage, equalAverage, dupe, colorScale: { cfvo: [{ type, val }],
 * colors } }] }], as the file has them (formulas without "=", relative to the first range's top-left cell).
 */
function readCondFormats(root, ns) {
  const flag = (el, name, dflt = false) => { const v = attr(el, name); return v == null ? dflt : v === '1' || v === 'true'; };
  // (Data bars' settings and Excel 2010's icon sets are in the worksheet's extLst: cfext.js.)
  const ext = readExtRules(root, ns);
  return kids(root, 'conditionalFormatting', ns).map((cf) => ({
    ranges: (attr(cf, 'sqref') || '').split(/\s+/).map(parseRange).filter(Boolean),
    rules: kids(cf, 'cfRule', ns).map((r) => {
      const scale = kid(r, 'colorScale', ns);
      return {
        type: attr(r, 'type'), priority: Number(attr(r, 'priority') || 0), dxfId: attr(r, 'dxfId') == null ? null : Number(attr(r, 'dxfId')),
        stopIfTrue: flag(r, 'stopIfTrue'), operator: attr(r, 'operator'), formulas: kids(r, 'formula', ns).map((f) => f.textContent),
        text: attr(r, 'text'), rank: Number(attr(r, 'rank') || 10), percent: flag(r, 'percent'), bottom: flag(r, 'bottom'),
        aboveAverage: flag(r, 'aboveAverage', true), equalAverage: flag(r, 'equalAverage'),
        colorScale: scale ? { cfvo: kids(scale, 'cfvo', ns).map((v) => ({ type: attr(v, 'type'), val: attr(v, 'val') })), colors: kids(scale, 'color', ns).map(readColor) } : null,
        dataBar: kid(r, 'dataBar', ns) ? { ...readDataBar(kid(r, 'dataBar', ns), ns), ...(ext.bars.get(extIdOf(r, ns)) || {}), fill: readDataBar(kid(r, 'dataBar', ns), ns).fill, showValue: readDataBar(kid(r, 'dataBar', ns), ns).showValue } : null,
        iconSet: kid(r, 'iconSet', ns) ? readIconSet(kid(r, 'iconSet', ns), ns) : null,
        ext: extIdOf(r, ns),
      };
    }),
  })).concat(ext.extra);
}

/**
 * The sheet's data validation rules: [{ ranges, type ('whole', 'decimal', 'list', 'date', 'time', 'textLength',
 * 'custom' or 'none'), operator, formulas (formula1, formula2, as the file has them, relative to the first range's
 * top-left cell), allowBlank, dropDown (a list's in-cell dropdown shown: the file's showDropDown means hidden),
 * errorStyle, errorTitle, error, showError, promptTitle, prompt, showPrompt }].
 */
function readValidations(root, ns) {
  const flag = (el, name) => attr(el, name) === '1' || attr(el, name) === 'true';
  return kids(kid(root, 'dataValidations', ns), 'dataValidation', ns).map((v) => ({
    ranges: (attr(v, 'sqref') || '').split(/\s+/).map(parseRange).filter(Boolean),
    type: attr(v, 'type') || 'none', operator: attr(v, 'operator') || 'between',
    formulas: [kid(v, 'formula1', ns), kid(v, 'formula2', ns)].filter(Boolean).map((f) => f.textContent),
    allowBlank: flag(v, 'allowBlank'), dropDown: !flag(v, 'showDropDown'),
    errorStyle: attr(v, 'errorStyle') || 'stop', errorTitle: attr(v, 'errorTitle') || '', error: attr(v, 'error') || '', showError: flag(v, 'showErrorMessage'),
    promptTitle: attr(v, 'promptTitle') || '', prompt: attr(v, 'prompt') || '', showPrompt: flag(v, 'showInputMessage'),
  }));
}

function readView(view, ns) {
  const pane = view && kid(view, 'pane', ns);
  const frozen = pane && /^frozen/.test(attr(pane, 'state') || '');
  const zoom = Number(view && attr(view, 'zoomScale')) || 100;
  return {
    freeze: { rows: frozen ? Math.trunc(Number(attr(pane, 'ySplit') || 0)) : 0, cols: frozen ? Math.trunc(Number(attr(pane, 'xSplit') || 0)) : 0 },
    zoom: Math.min(400, Math.max(10, zoom)),
    /** Formulas > Show Formulas (Ctrl+`): the sheet shows what its cells hold, not their values. */
    showFormulas: attr(view, 'showFormulas') === '1',
  };
}

/** What kind of formula a cell holds. A shared formula is written out on its group's first cell only (the one with
 *  `ref`); the others read it as moved to where they are. */
function formulaKind(cell, fEl, r, c, shared, arrays, dynamic = false) {
  const t = attr(fEl, 't');
  if (t === 'shared') {
    const si = attr(fEl, 'si');
    cell.shared = Number(si);
    if (cell.f !== '') shared.set(si, { r, c, text: cell.f });
    else if (shared.has(si)) {
      const first = shared.get(si);
      try { cell.f = shiftFormula(first.text, r - first.r, c - first.c); } catch { /* (left '': not worked out) */ }
    }
  } else if (t === 'array' && dynamic) {
    // (Excel 365's dynamic array formula, its cell metadata marking it: the range it spilled into, not a fixed array.)
    cell.dynamic = true;
    cell.spill = parseRange(attr(fEl, 'ref')) || { r0: r, c0: c, r1: r, c1: c };
  } else if (t === 'array') {
    cell.array = parseRange(attr(fEl, 'ref')) || { r0: r, c0: c, r1: r, c1: c };
    arrays.push(cell.array);
  } else if (t === 'dataTable') cell.dataTable = true;
}

/** Paper sizes by `paperSize` code, in inches (ECMA-376 Part 1, 18.3.1.63): only the common ones so far. */
const PAPER = { 1: [8.5, 11], 5: [8.5, 14], 8: [11.69, 16.54], 9: [8.27, 11.69], 11: [5.83, 8.27] };

/** What the sheet says about printing: paper, orientation, margins (inches), gridlines and headings, its scale or
 *  fit to pages (fitToWidth, fitToHeight: 0 as many as it takes), its manual page breaks (before rows and columns,
 *  0-based) and headers and footers. Without `pageMargins` Excel uses its Normal margins, 0.7in at the sides and
 *  0.75in top and bottom. */
function readPageSetup(root, ns) {
  const m = kid(root, 'pageMargins', ns);
  const ps = kid(root, 'pageSetup', ns);
  const po = kid(root, 'printOptions', ns);
  const sheetPr = kid(root, 'sheetPr', ns);
  const fitToPage = attr(kid(sheetPr, 'pageSetUpPr', ns), 'fitToPage') === '1';
  const margin = (name, d) => (m && attr(m, name) != null ? Number(attr(m, name)) : d);
  const on = (el, name) => !!el && (attr(el, name) === '1' || attr(el, name) === 'true');
  const code = Number(attr(ps, 'paperSize') || 1);
  return {
    paperSize: code,
    paper: PAPER[code] || PAPER[1],
    landscape: attr(ps, 'orientation') === 'landscape',
    scale: Number(attr(ps, 'scale') || 100),
    fitToPage,
    margins: { left: margin('left', 0.7), right: margin('right', 0.7), top: margin('top', 0.75), bottom: margin('bottom', 0.75), header: margin('header', 0.3), footer: margin('footer', 0.3) },
    gridLines: on(po, 'gridLines'),
    headings: on(po, 'headings'),
    pageOrder: attr(ps, 'pageOrder') || 'downThenOver',
    fitToWidth: Number(attr(ps, 'fitToWidth') ?? 1),
    fitToHeight: Number(attr(ps, 'fitToHeight') ?? 1),
    headerFooter: (() => {
      const hf = kid(root, 'headerFooter', ns);
      const text = (name) => kid(hf, name, ns)?.textContent || '';
      return {
        differentFirst: on(hf, 'differentFirst'), differentOddEven: on(hf, 'differentOddEven'),
        scaleWithDoc: !hf || attr(hf, 'scaleWithDoc') !== '0', alignWithMargins: !hf || attr(hf, 'alignWithMargins') !== '0',
        odd: { header: text('oddHeader'), footer: text('oddFooter') }, even: { header: text('evenHeader'), footer: text('evenFooter') },
        first: { header: text('firstHeader'), footer: text('firstFooter') },
      };
    })(),
    breaks: {
      rows: kids(kid(root, 'rowBreaks', ns), 'brk', ns).filter((b) => attr(b, 'man') === '1' || attr(b, 'man') === 'true').map((b) => Number(attr(b, 'id'))),
      cols: kids(kid(root, 'colBreaks', ns), 'brk', ns).filter((b) => attr(b, 'man') === '1' || attr(b, 'man') === 'true').map((b) => Number(attr(b, 'id'))),
    },
  };
}

/** A cell of the model, or null. */
export const cellAt = (sheet, r, c) => sheet.cells.get(r)?.get(c) || null;

