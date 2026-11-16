// The sheet view as Excel writes it (measured over COM and on its ribbon's commands, docs/VIEW.md): frozen panes as
// a pane element (xSplit and ySplit, the scrolled pane's top-left cell, the active pane, state frozen) with a
// selection for each pane, the zoom as zoomScale and zoomScaleNormal, Show Formulas as showFormulas="1", and the
// gridlines and headings hidden as showGridLines="0" and showRowColHeaders="0".
import { kids, kid, attr } from './xml.js';
import { refName } from './refs.js';

/** Excel's zoom limits, in percent. */
export const MIN_ZOOM = 10;
export const MAX_ZOOM = 400;

const make = (doc, ns, name) => doc.createElementNS(ns, doc.documentElement.prefix ? `${doc.documentElement.prefix}:${name}` : name);

/** The sheet's first sheetView, made if there is none. */
function sheetView(doc) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  let views = kid(root, 'sheetViews', ns);
  if (!views) {
    views = make(doc, ns, 'sheetViews');
    // (sheetViews comes after sheetPr and dimension.)
    const after = [...root.childNodes].find((n) => n.nodeType === 1 && !['sheetPr', 'dimension'].includes(n.localName));
    root.insertBefore(views, after || null);
  }
  let view = kid(views, 'sheetView', ns);
  if (!view) {
    view = make(doc, ns, 'sheetView');
    view.setAttribute('workbookViewId', '0');
    views.appendChild(view);
  }
  return { view, ns };
}

/** CT_SheetView's attributes in the schema's order, as Excel writes them. */
const VIEW_ATTRS = ['windowProtection', 'showFormulas', 'showGridLines', 'showRowColHeaders', 'showZeros', 'rightToLeft', 'tabSelected', 'showRuler',
  'showOutlineSymbols', 'defaultGridColor', 'showWhiteSpace', 'view', 'topLeftCell', 'colorId', 'zoomScale', 'zoomScaleNormal', 'zoomScaleSheetLayoutView',
  'zoomScalePageLayoutView', 'workbookViewId'];

/** The sheet view's topLeftCell set (null: A1, the attribute left out), its attributes kept in Excel's order. */
function setViewTopLeft(view, at) {
  const ref = at && (at.r || at.c) ? refName(at.r, at.c) : null;
  if ((view.getAttribute('topLeftCell') || null) === ref) return;
  const attrs = [...view.attributes].map((a) => [a.name, a.value]).filter(([n]) => n !== 'topLeftCell');
  if (ref) attrs.push(['topLeftCell', ref]);
  const rank = (n) => { const k = VIEW_ATTRS.indexOf(n); return k < 0 ? VIEW_ATTRS.length : k; };
  for (const [n] of [...view.attributes].map((a) => [a.name])) view.removeAttribute(n);
  for (const [n, v] of attrs.sort((a, b) => rank(a[0]) - rank(b[0]))) view.setAttribute(n, v);
}

/**
 * Freeze { rows, cols } (how many) counted from `top` and `left`, the first row and column in view when it was
 * frozen, or unfreeze ({ rows: 0, cols: 0 }); `active` is the active cell ({ r, c }). As Excel writes it (measured over
 * COM, tools/excel-compare/excel-freeze.ps1, test/freeze-excel.json): the cell first in view as the sheet view's
 * topLeftCell, with a selection of its own; the pane's topLeftCell the first cell past the freeze (column A when only
 * rows are frozen, row 1 when only columns); rows and columns, activePane bottomRight with a selection in each of the
 * other panes (the top-right one in the first row in view of the first scrolled column, the bottom-left one in the first
 * column in view of the first scrolled row); only rows, bottomLeft; only columns, topRight; the active cell's selection
 * in the active pane, A1 written as a selection with nothing in it. `split`: a split turned into a freeze, written
 * frozenSplit with the split's own pane selections kept, as Excel's Freeze Panes on a split window writes it.
 * Unfreezing takes the pane and its selections away and leaves the sheet view's topLeftCell as it was.
 */
export function setFreeze(doc, { rows, cols, top = 0, left = 0, split = false }, active = { r: 0, c: 0 }) {
  const { view, ns } = sheetView(doc);
  const old = kids(view, 'selection', ns).find((s) => !attr(s, 'pane') || attr(s, 'pane') === attr(kid(view, 'pane', ns), 'activePane')) || kids(view, 'selection', ns)[0];
  const sqref = old && attr(old, 'activeCell') === refName(active.r, active.c) ? attr(old, 'sqref') : null;
  const kept = split ? kids(view, 'selection', ns).filter((s) => ['topRight', 'bottomLeft'].includes(attr(s, 'pane'))) : [];
  for (const n of [...view.childNodes]) if (n.nodeType === 1 && (n.localName === 'pane' || n.localName === 'selection')) view.removeChild(n);
  const first = [...view.childNodes].find((n) => n.nodeType === 1) || null;
  const select = (pane, r, c, ref = null) => {
    const s = make(doc, ns, 'selection');
    if (pane) s.setAttribute('pane', pane);
    if (r || c || ref) {
      s.setAttribute('activeCell', refName(r, c));
      s.setAttribute('sqref', ref || refName(r, c));
    }
    view.insertBefore(s, first);
  };
  if (!rows && !cols) {
    if (active.r || active.c || sqref) select(null, active.r, active.c, sqref);
    return;
  }
  setViewTopLeft(view, { r: top, c: left });
  const pane = make(doc, ns, 'pane');
  if (cols) pane.setAttribute('xSplit', String(cols));
  if (rows) pane.setAttribute('ySplit', String(rows));
  pane.setAttribute('topLeftCell', refName(rows ? top + rows : 0, cols ? left + cols : 0));
  const activePane = rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight';
  pane.setAttribute('activePane', activePane);
  pane.setAttribute('state', split ? 'frozenSplit' : 'frozen');
  view.insertBefore(pane, first);
  // (The top-left pane's selection: only when the window doesn't start at A1, as Excel leaves it out then.)
  if (top || left) select(null, top, left);
  if (kept.length) for (const s of kept) view.insertBefore(s, first);
  else if (rows && cols) {
    select('topRight', top, left + cols);
    select('bottomLeft', top + rows, left);
  }
  select(activePane, active.r, active.c, sqref);
}

/** Show Formulas on (showFormulas="1", as Excel writes it) or off (no attribute). */
export function setShowFormulas(doc, on) {
  const { view } = sheetView(doc);
  if (on) view.setAttribute('showFormulas', '1'); else view.removeAttribute('showFormulas');
}

/** View > Show: the gridlines (showGridLines) or the row and column headings (showRowColHeaders) on or off. As Excel
 *  writes them: "0" when off, the attribute gone when on again (measured over COM, Window.DisplayGridlines and
 *  DisplayHeadings, and the files Excel saved). */
export function setShown(doc, name, on) {
  const { view } = sheetView(doc);
  if (on) view.removeAttribute(name); else view.setAttribute(name, '0');
}

/** The zoom in percent (Excel's 10 to 400; 100 takes the attributes away). */
export function setZoom(doc, percent) {
  const { view } = sheetView(doc);
  const z = Math.round(Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, percent)));
  for (const name of ['zoomScale', 'zoomScaleNormal']) {
    if (z === 100) view.removeAttribute(name); else view.setAttribute(name, String(z));
  }
  // (Excel writes them before workbookViewId; attribute order doesn't matter to a reader.)
}
