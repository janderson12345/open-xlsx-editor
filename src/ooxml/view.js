// The sheet view as Excel writes it (measured over COM and on its ribbon's commands, docs/VIEW.md): frozen panes as
// a pane element (xSplit and ySplit, the scrolled pane's top-left cell, the active pane, state frozen) with a
// selection for each pane, and the zoom as zoomScale and zoomScaleNormal.
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

/**
 * Freeze the rows above and the columns left of a cell ({ rows, cols }: how many), or unfreeze ({ rows: 0, cols: 0 }).
 * `active` is the active cell ({ r, c }), `topLeft` the first cell of the scrolled pane. As Excel writes it: rows and
 * columns, activePane bottomRight with a selection in each of the three panes (the top-right one in row 1 of the
 * first scrolled column, the bottom-left one in column A of the first scrolled row); only rows, bottomLeft; only
 * columns, topRight; the active cell's selection in the active pane.
 */
export function setFreeze(doc, { rows, cols }, active = { r: 0, c: 0 }, topLeft = null) {
  const { view, ns } = sheetView(doc);
  const old = kids(view, 'selection', ns).find((s) => !attr(s, 'pane') || attr(s, 'pane') === attr(kid(view, 'pane', ns), 'activePane')) || kids(view, 'selection', ns)[0];
  const sqref = old && attr(old, 'activeCell') === refName(active.r, active.c) ? attr(old, 'sqref') : null;
  for (const n of [...view.childNodes]) if (n.nodeType === 1 && (n.localName === 'pane' || n.localName === 'selection')) view.removeChild(n);
  const first = [...view.childNodes].find((n) => n.nodeType === 1) || null;
  const select = (pane, r, c, ref = null) => {
    const s = make(doc, ns, 'selection');
    if (pane) s.setAttribute('pane', pane);
    s.setAttribute('activeCell', refName(r, c));
    s.setAttribute('sqref', ref || refName(r, c));
    view.insertBefore(s, first);
  };
  if (!rows && !cols) {
    if (active.r || active.c || sqref) select(null, active.r, active.c, sqref);
    return;
  }
  const pane = make(doc, ns, 'pane');
  if (cols) pane.setAttribute('xSplit', String(cols));
  if (rows) pane.setAttribute('ySplit', String(rows));
  const tl = topLeft || { r: rows, c: cols };
  pane.setAttribute('topLeftCell', refName(Math.max(rows, tl.r), Math.max(cols, tl.c)));
  const activePane = rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight';
  pane.setAttribute('activePane', activePane);
  pane.setAttribute('state', 'frozen');
  view.insertBefore(pane, first);
  if (rows && cols) {
    select('topRight', 0, cols);
    select('bottomLeft', rows, 0);
  }
  select(activePane, active.r, active.c, sqref);
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
