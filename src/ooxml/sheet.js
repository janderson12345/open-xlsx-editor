// A worksheet part (xl/worksheets/sheetN.xml) read into a sparse model: cells by row and column, column widths, row
// heights, merged ranges and the sheet's defaults. The part's XML document is kept with the model so that an edit
// changes only the elements it has to and everything not modelled (conditional formats, validations, extLst...) stays
// where it was.
import { readOutlinePr } from './outline.js';
import { kids, kid, attr, parseXml } from './xml.js';
import { parseRef, parseRange, MAX_COLS } from './refs.js';
import { stringItemText, stringItemRuns } from './strings.js';
import { makeCell, withRuns, outlineLevel, rowEntry, formulaKind } from './sheetcells.js';
import { findSheetData, scanSheetData } from './sheetscan.js';
import { readColor } from './styles.js';
import { readExtRules, readDataBar, readIconSet, extIdOf } from './cfext.js';
import { readSparklines } from './sparklines.js';
import { readFilterColumn } from './filtermore.js';

/** Read one `c` element's value, resolving shared strings. */
export function readCell(c, ns, sst) {
  const s = Number(attr(c, 's') || 0);
  const t = attr(c, 't') || 'n';
  const fEl = kid(c, 'f', ns);
  const f = fEl ? fEl.textContent : null;
  const vEl = kid(c, 'v', ns);
  const v = vEl ? vEl.textContent : null;
  switch (t) {
    case 's': return withRuns(makeCell('s', v == null ? null : (sst[Number(v)] ?? ''), s, f), v == null ? null : sst.runs?.get(Number(v)));
    case 'inlineStr': return withRuns(makeCell('s', stringItemText(kid(c, 'is', ns), ns), s, f), stringItemRuns(kid(c, 'is', ns), ns));
    case 'b': return makeCell('b', v == null ? null : v === '1' || v === 'true', s, f);
    case 'e': case 'str': case 'd': return makeCell(t, v, s, f);
    default: return makeCell('n', v == null || v === '' ? null : Number(v), s, f);
  }
}

/** The format an empty cell (r, c) of a sheet's model has: its row's own, else its column's (0, none). */
export function lineStyle(sheet, r, c) {
  return sheet.rows.get(r)?.s || sheet.cols.find((x) => c >= x.min && c <= x.max)?.style || 0;
}

/** A col element's min or max as a 0-based column within the sheet. */
const colIndex = (v) => Math.max(0, Math.min(MAX_COLS - 1, (Math.trunc(Number(v)) || 1) - 1));

/** The cells of a parsed sheet's sheetData: { cells, rows, arrays, maxRow, maxCol } (sheetscan.js reads the same
 *  from the part's bytes; test/sheetscan.test.js holds the two to the same answer). */
function readSheetData(root, ns, sst) {
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
    const entry = rowEntry(attr(rowEl, 'ht'), attr(rowEl, 'hidden'), attr(rowEl, 'customFormat'), attr(rowEl, 's'), attr(rowEl, 'outlineLevel'), attr(rowEl, 'collapsed'), attr(rowEl, 'thickBot'), attr(rowEl, 'customHeight'));
    if (entry) rows.set(r, entry);
    let c = -1;
    let line = null;
    for (const cEl of kids(rowEl, 'c', ns)) {
      const at = parseRef(attr(cEl, 'r'));
      c = at ? at.c : c + 1;
      if (!line) cells.set(r, (line = new Map()));
      const cell = readCell(cEl, ns, sst);
      if (cell.f != null) { const fEl = kid(cEl, 'f', ns); formulaKind(cell, attr(fEl, 't'), attr(fEl, 'si'), attr(fEl, 'ref'), r, c, shared, arrays, attr(cEl, 'cm') != null); }
      line.set(c, cell);
      if (c > maxCol) maxCol = c;
    }
    if (line && r > maxRow) maxRow = r;
  }
  return { cells, rows, arrays, maxRow, maxCol };
}

/** The sheet's model from its parsed XML and the workbook's shared strings. `data`: its cells already read from the
 *  part's bytes (readSheetBytes), `doc` then being the part around them. */
export function readSheet(doc, sst = [], data = null) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const { cells, rows, arrays } = data || (data = readSheetData(root, ns, sst));
  let { maxRow, maxCol } = data;
  const cols = [];
  for (const colsEl of kids(root, 'cols', ns)) {
    for (const col of kids(colsEl, 'col', ns)) {
      cols.push({
        // (Within the sheet's 16,384 columns, whatever the file says: a crafted max of two billion isn't walked.)
        min: colIndex(attr(col, 'min')),
        max: colIndex(attr(col, 'max')),
        width: attr(col, 'width') == null ? null : Number(attr(col, 'width')),
        hidden: attr(col, 'hidden') === '1' || attr(col, 'hidden') === 'true',
        style: Number(attr(col, 'style') || 0),
        level: outlineLevel(attr(col, 'outlineLevel')),
        collapsed: attr(col, 'collapsed') === '1' || attr(col, 'collapsed') === 'true',
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
    outline: readOutlinePr(root, ns, view),
    page: readPageSetup(root, ns),
    autoFilter: readAutoFilter(kid(root, 'autoFilter', ns), ns),
    condFormats: readCondFormats(root, ns),
    validations: readValidations(root, ns),
    sparklines: readSparklines(root, ns),
    maxRow,
    maxCol,
  };
}

/**
 * The sheet's AutoFilter: { r0, c0, r1, c1 (its range, the header row first), columns: Map(column index => { values:
 * [shown texts], blank (blanks shown too), other (a kind of filter not read here: custom, top 10, colour, dates) }) },
 * or null.
 */
/** An autoFilter element (the sheet's, or a table part's) as { r0, c0, r1, c1, columns: Map(c -> filter) }. */
export function readAutoFilter(el, ns) {
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
        text: attr(r, 'text'), timePeriod: attr(r, 'timePeriod'), rank: Number(attr(r, 'rank') || 10), percent: flag(r, 'percent'), bottom: flag(r, 'bottom'),
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
 * errorStyle, errorTitle, error, showError, promptTitle, prompt, showPrompt, ext }]. The main rules, then those in
 * the extLst in Excel's 2010 form (`ext`: x14:dataValidation, its formulas in xm:f and its ranges in xm:sqref), which
 * Excel writes for a rule whose formulas read another sheet (tools/excel-compare/excel-dv-x14.ps1).
 */
function readValidations(root, ns) {
  const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
  const XM = 'http://schemas.microsoft.com/office/excel/2006/main';
  const ext = [...kids(kid(root, 'extLst', ns), 'ext', ns)].flatMap((e) => kids(kid(e, 'dataValidations', X14), 'dataValidation', X14));
  const formulaOf = (v, name) => kid(kid(v, name, X14), 'f', XM);
  return [...kids(kid(root, 'dataValidations', ns), 'dataValidation', ns), ...ext].map((v) => rule(v, v.namespaceURI === X14 ? {
    sqref: kid(v, 'sqref', XM)?.textContent, formulas: [formulaOf(v, 'formula1'), formulaOf(v, 'formula2')], ext: true,
  } : { sqref: attr(v, 'sqref'), formulas: [kid(v, 'formula1', ns), kid(v, 'formula2', ns)], ext: false }));
}

/** One rule of readValidations, from its element and where its ranges and formulas are. */
function rule(v, { sqref, formulas, ext }) {
  const flag = (el, name) => attr(el, name) === '1' || attr(el, name) === 'true';
  return {
    ranges: (sqref || '').split(/\s+/).map(parseRange).filter(Boolean),
    type: attr(v, 'type') || 'none', operator: attr(v, 'operator') || 'between',
    formulas: formulas.filter(Boolean).map((f) => f.textContent),
    allowBlank: flag(v, 'allowBlank'), dropDown: !flag(v, 'showDropDown'),
    errorStyle: attr(v, 'errorStyle') || 'stop', errorTitle: attr(v, 'errorTitle') || '', error: attr(v, 'error') || '', showError: flag(v, 'showErrorMessage'),
    promptTitle: attr(v, 'promptTitle') || '', prompt: attr(v, 'prompt') || '', showPrompt: flag(v, 'showInputMessage'),
    ext,
  };
}

function readView(view, ns) {
  const pane = view && kid(view, 'pane', ns);
  const frozen = pane && /^frozen/.test(attr(pane, 'state') || '');
  const zoom = Number(view && attr(view, 'zoomScale')) || 100;
  // (The sheet view's topLeftCell: where the window starts, so where a freeze made while scrolled starts; the pane's:
  // where the scrolled part was left. Excel opens a sheet at them, docs/VIEW.md.)
  const at = (el) => { const p = el && attr(el, 'topLeftCell') ? parseRef(attr(el, 'topLeftCell')) : null; return p ? { r: p.r, c: p.c } : null; };
  const start = at(view) || { r: 0, c: 0 };
  const rows = frozen ? Math.trunc(Number(attr(pane, 'ySplit') || 0)) : 0;
  const cols = frozen ? Math.trunc(Number(attr(pane, 'xSplit') || 0)) : 0;
  // (A split that isn't frozen: its sizes are in twentieths of a point, the headings included; shown unsplit.)
  const split = pane && !frozen ? { x: Number(attr(pane, 'xSplit') || 0), y: Number(attr(pane, 'ySplit') || 0) } : null;
  return {
    freeze: { rows, cols, top: rows ? start.r : 0, left: cols ? start.c : 0 },
    viewTopLeft: start,
    scrolledTo: (frozen && at(pane)) || start,
    split,
    zoom: Math.min(400, Math.max(10, zoom)),
    /** Formulas > Show Formulas (Ctrl+`): the sheet shows what its cells hold, not their values. */
    showFormulas: attr(view, 'showFormulas') === '1',
  };
}

/** Paper sizes by `paperSize` code, in inches (ECMA-376 Part 1, 18.3.1.63): only the common ones so far. */
const PAPER = { 1: [8.5, 11], 3: [11, 17], 5: [8.5, 14], 7: [7.25, 10.5], 8: [11.69, 16.54], 9: [8.27, 11.69], 11: [5.83, 8.27], 13: [7.17, 10.12] };

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
    centerH: on(po, 'horizontalCentered'),
    centerV: on(po, 'verticalCentered'),
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


/**
 * A worksheet part read from its bytes: its cells by scanning them (sheetscan.js), the rest by parsing the part
 * without its sheetData, which is small whatever the sheet holds. { model, frame (that parsed rest: the part's
 * document with an empty sheetData) }, or null when the part has to be parsed whole (readSheet on its document).
 * `compact`: how many formatted empty cells it takes to keep them as bands (sheetcells.js), Infinity for never.
 */
export function readSheetBytes(u8, sst = [], { compact } = {}) {
  const at = findSheetData(u8);
  if (!at) return null;
  const parseIs = (xml) => {
    const frag = parseXml(at.rootOpen + xml + at.rootClose).documentElement;
    const ns = frag.namespaceURI;
    return [stringItemText(kid(frag, 'is', ns), ns), stringItemRuns(kid(frag, 'is', ns), ns)];
  };
  const data = scanSheetData(u8, at, sst, { ...(compact == null ? {} : { compact }), parseIs });
  if (!data) return null;
  const decoder = new TextDecoder();
  const frame = parseXml(`${decoder.decode(u8.subarray(0, at.start))}<${at.prefix}sheetData/>${decoder.decode(u8.subarray(at.end))}`);
  return { model: readSheet(frame, sst, data), frame };
}
