// How a table (ListObject) looks in its table style (docs/TABLES.md): Excel's built-in styles (src/engine/tablestyles.js,
// measured) or the workbook's own (styles.xml's tableStyles, each element a dxf), with the table's Table Style Options.
// Each element covers part of the table, and later ones are laid over earlier ones in Excel's order: the whole table,
// the column stripes, the row stripes, the last column, the first column, the header row, the total row, then the
// header's and total's first and last cells (measured: every cell of all 60 built-in styles with every option on, and
// with only the header and total rows, against Excel's DisplayFormat, test/tablelook-*-excel.json). An element's border
// sides are its region's outline, its vertical and horizontal ones the lines inside; a stripe's region is each band.
// The cell's own format is laid over the table's (tablelook-excel.json's own cases): a fill, a border side, and a font
// colour, bold or italic other than its cell style's.
import { parseXml, NS } from './xml.js';
import { readDxf, xfOf } from './styles.js';
import { lineStyle } from './sheet.js';
import { resolveColor } from './theme.js';
import { TABLE_DXFS, TABLE_STYLES } from '../engine/tablestyles.js';

/** The built-in styles' names, in the gallery's order (Light, Medium, Dark). */
export const BUILTIN_TABLE_STYLES = Object.keys(TABLE_STYLES);

const parsed = new Map();
/** A built-in dxf, read once. */
function builtinDxf(i) {
  if (!parsed.has(i)) {
    const doc = parseXml(`<dxf xmlns="${NS.x}">${TABLE_DXFS[i]}</dxf>`);
    parsed.set(i, readDxf(doc.documentElement, NS.x));
  }
  return parsed.get(i);
}

/** Style `name`'s elements: { type: { dxf, size } }, the workbook's own style first (as Excel looks it up), else a
 *  built-in one; null for no style or one that isn't there. */
export function tableStyleElements(styles, name) {
  if (!name) return null;
  const own = styles?.tableStyles?.[name];
  if (own) {
    const out = {};
    for (const [type, e] of Object.entries(own)) if (styles.dxfs[e.dxfId]) out[type] = { dxf: styles.dxfs[e.dxfId], size: e.size };
    return out;
  }
  const b = TABLE_STYLES[name];
  if (!b) return null;
  return Object.fromEntries(Object.entries(b).map(([type, i]) => [type, { dxf: builtinDxf(i), size: 1 }]));
}

/** The regions each element covers in `t` ({ r0, c0, r1, c1, headers, totals, style: { first, last, rows, cols } }),
 *  in the order they are laid on: [{ type, areas: [{ r0, c0, r1, c1 }] }]. Stripe sizes from `els`. */
function regions(t, els, win = null) {
  const o = t.style;
  const d0 = t.r0 + t.headers;
  const d1 = t.r1 - t.totals;
  const out = [{ type: 'wholeTable', areas: [{ r0: t.r0, c0: t.c0, r1: t.r1, c1: t.c1 }] }];
  const bands = (from, to, a, b) => {
    // (First and second stripes, each of its size, taking turns from the first data row or column.)
    const first = [];
    const second = [];
    const s1 = Math.max(1, (els.firstRowStripe && a === 'row' ? els.firstRowStripe.size : els.firstColumnStripe && a === 'col' ? els.firstColumnStripe.size : 1) || 1);
    const s2 = Math.max(1, (a === 'row' ? els.secondRowStripe?.size : els.secondColumnStripe?.size) || 1);
    // (Only the bands round the window: from the cycle it starts in, a line before it, to a line after it.)
    const w0 = win ? (a === 'row' ? win.r0 : win.c0) - 1 : from;
    const w1 = win ? (a === 'row' ? win.r1 : win.c1) + 1 : to;
    const start = from + Math.floor(Math.max(0, w0 - from) / (s1 + s2)) * (s1 + s2);
    for (let i = start, k = 0; i <= Math.min(to, w1); k ^= 1) {
      const n = k ? s2 : s1;
      (k ? second : first).push(b(i, Math.min(to, i + n - 1)));
      i += n;
    }
    return [first, second];
  };
  if (o.cols && d1 >= d0) {
    const [a, b] = bands(t.c0, t.c1, 'col', (x, y) => ({ r0: d0, c0: x, r1: d1, c1: y }));
    out.push({ type: 'firstColumnStripe', areas: a }, { type: 'secondColumnStripe', areas: b });
  }
  if (o.rows && d1 >= d0) {
    const [a, b] = bands(d0, d1, 'row', (x, y) => ({ r0: x, c0: t.c0, r1: y, c1: t.c1 }));
    out.push({ type: 'firstRowStripe', areas: a }, { type: 'secondRowStripe', areas: b });
  }
  if (o.last) out.push({ type: 'lastColumn', areas: [{ r0: t.r0, c0: t.c1, r1: t.r1, c1: t.c1 }] });
  if (o.first) out.push({ type: 'firstColumn', areas: [{ r0: t.r0, c0: t.c0, r1: t.r1, c1: t.c0 }] });
  if (t.headers) out.push({ type: 'headerRow', areas: [{ r0: t.r0, c0: t.c0, r1: t.r0, c1: t.c1 }] });
  if (t.totals) out.push({ type: 'totalRow', areas: [{ r0: t.r1, c0: t.c0, r1: t.r1, c1: t.c1 }] });
  const cell = (r, c) => [{ r0: r, c0: c, r1: r, c1: c }];
  if (t.headers && o.first) out.push({ type: 'firstHeaderCell', areas: cell(t.r0, t.c0) });
  if (t.headers && o.last) out.push({ type: 'lastHeaderCell', areas: cell(t.r0, t.c1) });
  if (t.totals && o.first) out.push({ type: 'firstTotalCell', areas: cell(t.r1, t.c0) });
  if (t.totals && o.last) out.push({ type: 'lastTotalCell', areas: cell(t.r1, t.c1) });
  return out;
}

/**
 * Table `t`'s look in style elements `els` (tableStyleElements): Map "r,c" -> { fill, color, bold, italic, border:
 * { top, bottom, left, right: { style, color } } } for each cell of the table, colours as '#RRGGBB' through `theme`
 * and `palette`. Each edge is one line, the last element's to set it. With `raw`, each look also has `raw`: the
 * colours as the dxfs give them ({ fill, color, border: { side: colour } }), for formats written from it.
 */
export function tableLook(t, els, theme, palette, { raw = false, win = null } = {}) {
  const out = new Map();
  if (!els) return out;
  // (With `win` ({ r0, c0, r1, c1 }), only its cells: a table's look worked out where it is drawn, not all of a huge one.)
  const R0 = win ? Math.max(t.r0, win.r0) : t.r0;
  const R1 = win ? Math.min(t.r1, win.r1) : t.r1;
  const C0 = win ? Math.max(t.c0, win.c0) : t.c0;
  const C1 = win ? Math.min(t.c1, win.c1) : t.c1;
  if (R0 > R1 || C0 > C1) return out;
  const color = (c) => (c ? resolveColor(c, theme, palette) : null);
  const edges = new Map();
  const setEdge = (key, side) => { if (side?.style) edges.set(key, { style: side.style, color: side.color?.auto ? '#000000' : color(side.color) || '#000000', raw: side.color }); };
  for (const { type, areas } of regions(t, els, win)) {
    const dxf = els[type]?.dxf;
    if (!dxf) continue;
    for (const a of areas) {
      for (let r = Math.max(a.r0, R0); r <= Math.min(a.r1, R1); r++) {
        for (let c = Math.max(a.c0, C0); c <= Math.min(a.c1, C1); c++) {
          const look = out.get(`${r},${c}`) || {};
          if (dxf.fill) { const f = color(dxf.fill); if (f) { look.fill = f; if (raw) look.raw = { ...look.raw, fill: dxf.fill }; } }
          const f = dxf.font;
          if (f?.color) { const fc = color(f.color); if (fc) { look.color = fc; if (raw) look.raw = { ...look.raw, color: f.color }; } }
          if (f?.b !== undefined) look.bold = f.b;
          if (f?.i !== undefined) look.italic = f.i;
          out.set(`${r},${c}`, look);
        }
      }
      const b = dxf.border;
      if (!b) continue;
      // (Edges: "h r,c" the line above cell (r, c), "v r,c" the line left of it.)
      for (let c = Math.max(a.c0, C0); c <= Math.min(a.c1, C1); c++) {
        setEdge(`h${a.r0},${c}`, b.top);
        setEdge(`h${a.r1 + 1},${c}`, b.bottom);
        for (let r = Math.max(a.r0 + 1, R0); r <= Math.min(a.r1, R1 + 1); r++) setEdge(`h${r},${c}`, b.horizontal);
      }
      for (let r = Math.max(a.r0, R0); r <= Math.min(a.r1, R1); r++) {
        setEdge(`v${r},${a.c0}`, b.left);
        setEdge(`v${r},${a.c1 + 1}`, b.right);
        for (let c = Math.max(a.c0 + 1, C0); c <= Math.min(a.c1, C1 + 1); c++) setEdge(`v${r},${c}`, b.vertical);
      }
    }
  }
  for (let r = R0; r <= R1; r++) {
    for (let c = C0; c <= C1; c++) {
      const sides = { top: edges.get(`h${r},${c}`), bottom: edges.get(`h${r + 1},${c}`), left: edges.get(`v${r},${c}`), right: edges.get(`v${r},${c + 1}`) };
      const border = Object.fromEntries(Object.entries(sides).filter(([, v]) => v).map(([k, v]) => [k, { style: v.style, color: v.color }]));
      const look = out.get(`${r},${c}`) || {};
      if (raw && Object.keys(border).length) look.raw = { ...look.raw, border: Object.fromEntries(Object.entries(sides).filter(([, v]) => v).map(([k, v]) => [k, v.raw])) };
      if (Object.keys(border).length) look.border = border;
      if (Object.keys(look).length) out.set(`${r},${c}`, look);
    }
  }
  return out;
}

/**
 * The table's look of a cell under its own format (tablelook-excel.json's own cases): `own` { fill, font, border } is
 * the cell's (styles.js), `styleFont` its cell style's font. Its fill (any pattern) and border sides take the table's
 * place; its font colour, bold and italic only where they differ from the cell style's (bold turned off on a header
 * cell stays bold, an rgb black on a white header shows black). Returns the look left for the table to draw.
 */
export function withOwnFormat(look, own, styleFont, theme, palette) {
  if (!look) return look;
  const out = { ...look };
  if (own.fill && own.fill.pattern !== 'none') delete out.fill;
  const f = own.font;
  const same = (a, b) => JSON.stringify(a || null) === JSON.stringify(b || null);
  const key = (c) => (c ? { rgb: c.rgb ? resolveColor(c, theme, palette) : null, theme: c.rgb ? null : c.theme ?? null, indexed: c.indexed ?? null, tint: c.tint || 0, auto: !!c.auto } : null);
  if (f && !same(key(f.color), key(styleFont?.color))) out.color = resolveColor(f.color, theme, palette) || '#000000';
  if (f && !!f.b !== !!styleFont?.b) out.bold = !!f.b;
  if (f && !!f.i !== !!styleFont?.i) out.italic = !!f.i;
  if (out.border && own.border) {
    out.border = { ...out.border };
    for (const side of ['top', 'bottom', 'left', 'right']) {
      const s = own.border[side];
      if (s?.style) out.border[side] = { style: s.style, color: s.color?.auto || !s.color ? '#000000' : resolveColor(s.color, theme, palette) || '#000000' };
    }
  }
  return out;
}

// (About a screen of cells: 64 rows by 32 columns.)
const BLOCK_ROWS = 64;
const BLOCK_COLS = 32;
/** What the tables on sheet `si` of `wb` draw (readTables' tables): { get("r,c") } -> look (tableLook), less what each
 *  cell's own format draws itself (withOwnFormat; its own border sides left out, drawn as the cell's). Worked out a
 *  block of 64 rows by 32 columns at a time as the grid or print asks for them, and kept: a table of a whole sheet costs only
 *  what is drawn. */
export function tableLooks(wb, tables, si) {
  const mine = tables.filter((t) => t.si === si);
  const blocks = new Map();
  return {
    get(key) {
      const [r, c] = key.split(',').map(Number);
      const id = `${Math.floor(r / BLOCK_ROWS)},${Math.floor(c / BLOCK_COLS)}`;
      if (!blocks.has(id)) {
        const r0 = Math.floor(r / BLOCK_ROWS) * BLOCK_ROWS;
        const c0 = Math.floor(c / BLOCK_COLS) * BLOCK_COLS;
        blocks.set(id, blockLooks(wb, mine, si, { r0, c0, r1: r0 + BLOCK_ROWS - 1, c1: c0 + BLOCK_COLS - 1 }));
      }
      return blocks.get(id).get(key);
    },
  };
}

/** The looks of tables `tables` in window `win` of sheet si (tableLooks). */
function blockLooks(wb, tables, si, win) {
  const out = new Map();
  const styles = wb.styles;
  const sheet = wb.sheet(si);
  for (const t of tables) {
    if (t.r1 < win.r0 || t.r0 > win.r1 || t.c1 < win.c0 || t.c0 > win.c1) continue;
    for (const [k, look] of tableLook(t, tableStyleElements(styles, t.style.name), wb.theme, styles.palette, { win })) {
      const [r, c] = k.split(',').map(Number);
      const xf = xfOf(styles, sheet.cells.get(r)?.get(c)?.s ?? lineStyle(sheet, r, c) ?? 0);
      const border = styles.borders[xf.borderId];
      const own = { fill: styles.fills[xf.fillId], font: styles.fonts[xf.fontId], border: null };
      const styleFont = styles.fonts[styles.styleXfs?.[xf.xfId]?.fontId] || styles.normal;
      const left = withOwnFormat(look, own, styleFont, wb.theme, styles.palette);
      if (left.border && border) {
        left.border = Object.fromEntries(Object.entries(left.border).filter(([side]) => !border[side]?.style));
        if (!Object.keys(left.border).length) delete left.border;
      }
      out.set(k, left);
    }
  }
  return out;
}
