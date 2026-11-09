// A worksheet part (xl/worksheets/sheetN.xml) read into a sparse model: cells by row and column, column widths, row
// heights, merged ranges and the sheet's defaults. The part's XML document is kept with the model so that an edit
// changes only the elements it has to and everything not modelled (conditional formats, validations, extLst...) stays
// where it was.
import { kids, kid, attr } from './xml.js';
import { parseRef, parseRange } from './refs.js';
import { stringItemText } from './strings.js';

/** A cell as stored: `t` its type (n number, s string, b boolean, e error, str formula text, d ISO date), `v` its
 *  value (null for a cell that only carries a style), `s` its style (an index into cellXfs), `f` its formula text. */
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
      line.set(c, readCell(cEl, ns, sst));
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
    format: { defaultRowHeight: num('defaultRowHeight') ?? 15, defaultColWidth: num('defaultColWidth'), baseColWidth: num('baseColWidth') },
    showGridLines: !view || attr(view, 'showGridLines') !== '0',
    maxRow,
    maxCol,
  };
}

/** A cell of the model, or null. */
export const cellAt = (sheet, r, c) => sheet.cells.get(r)?.get(c) || null;

