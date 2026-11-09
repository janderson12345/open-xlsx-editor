// What the Home tab's buttons do to a selection, as Excel does it (measured over COM, docs/FORMATS.md). Each function
// returns the cells' new formats, [{ r, c, s }], for the editor to apply and keep in its undo history; nothing here
// changes a cell. A toggle (Bold, Italic, Underline, Strikethrough, Wrap Text) follows the active cell: on if it is
// off there, for every cell. Grow and Shrink Font step from the active cell's size through Excel's list.
import { xfOf, formatCodeOf } from './styles.js';
import { isDateFormat } from './format.js';

/** Excel's font sizes, which Grow and Shrink Font step through (no further than 8 and 72). */
export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
/** Excel's border lines: thin, the "thick" of the menu (stored as medium) and double. */
const THIN = { style: 'thin' };
const THICK = { style: 'medium' };
const DOUBLE = { style: 'double' };
/** At most this many cells are formatted at once; a selection of whole rows or columns is cut to the used range. */
export const MAX_CELLS = 1 << 20;

/** The cells of a selection, cut to the sheet's used range when it runs past it (whole rows or columns). */
export function cellsOf(sheet, sel) {
  const r1 = sel.r1 - sel.r0 > 10000 ? Math.max(sel.r0, Math.min(sel.r1, sheet.maxRow)) : sel.r1;
  const c1 = sel.c1 - sel.c0 > 1000 ? Math.max(sel.c0, Math.min(sel.c1, sheet.maxCol)) : sel.c1;
  const out = [];
  for (let r = sel.r0; r <= r1; r++) for (let c = sel.c0; c <= c1 && out.length < MAX_CELLS; c++) out.push({ r, c });
  return out;
}

const styleAt = (sheet, r, c) => sheet.cells.get(r)?.get(c)?.s || 0;

/** Every cell of the selection with the same change (stylesedit.js's form). */
export function formatEach(wb, si, sel, change) {
  const sheet = wb.sheet(si);
  const memo = new Map(); // (cells sharing a format share the result)
  return cellsOf(sheet, sel).map(({ r, c }) => {
    const s = styleAt(sheet, r, c);
    if (!memo.has(s)) memo.set(s, wb.styleWith(s, typeof change === 'function' ? change(s) : change));
    return { r, c, s: memo.get(s) };
  });
}

/** The active cell's font and cell format. */
export function activeLook(wb, si, sel) {
  const s = styleAt(wb.sheet(si), sel.ar, sel.ac);
  const xf = xfOf(wb.styles, s);
  return { s, xf, font: wb.styles.fonts[xf.fontId], code: formatCodeOf(wb.styles, s) };
}

/** Bold, Italic, Strikethrough (`key` b, i, strike) or Underline (u): on for all unless the active cell has it. */
export function toggleFont(wb, si, sel, key) {
  const { font } = activeLook(wb, si, sel);
  const value = key === 'u' ? (font.u ? null : 'single') : !font[key];
  return formatEach(wb, si, sel, { font: { [key]: value } });
}

/** The font size one step up or down Excel's list from `size` (a size between steps goes to the next one). */
export function stepSize(size, dir) {
  if (dir > 0) return FONT_SIZES.find((n) => n > size) ?? FONT_SIZES.at(-1);
  return [...FONT_SIZES].reverse().find((n) => n < size) ?? FONT_SIZES[0];
}

/** Grow Font (1) or Shrink Font (-1): every cell gets the size a step from the active cell's. */
export function stepFont(wb, si, sel, dir) {
  return formatEach(wb, si, sel, { font: { size: stepSize(activeLook(wb, si, sel).font.size, dir) } });
}

/** Wrap Text: on for all unless the active cell wraps. */
export function toggleWrap(wb, si, sel) {
  return formatEach(wb, si, sel, { align: { wrap: !activeLook(wb, si, sel).xf.align.wrap } });
}

/** Increase (1) or Decrease (-1) Indent, cell by cell: up to 250 levels. Indenting makes a cell left-aligned unless it
 *  is right-aligned (Excel indents those from the right). */
export function stepIndent(wb, si, sel, dir) {
  return formatEach(wb, si, sel, (s) => {
    const a = xfOf(wb.styles, s).align;
    const indent = Math.max(0, Math.min(250, (a.indent || 0) + dir));
    if (dir < 0) return { align: { indent } };
    return { align: { indent, horizontal: a.horizontal === 'right' ? 'right' : 'left' } };
  });
}

/** The sides each border preset sets on a cell at (r, c) of the selection; null for none, undefined to leave a side. */
const PRESETS = {
  bottom: (at) => ({ bottom: at.bottom ? THIN : undefined }),
  top: (at) => ({ top: at.top ? THIN : undefined }),
  left: (at) => ({ left: at.left ? THIN : undefined }),
  right: (at) => ({ right: at.right ? THIN : undefined }),
  none: () => ({ left: null, right: null, top: null, bottom: null }),
  all: () => ({ left: THIN, right: THIN, top: THIN, bottom: THIN }),
  outside: (at) => edges(at, THIN),
  thickOutside: (at) => edges(at, THICK),
  doubleBottom: (at) => ({ bottom: at.bottom ? DOUBLE : undefined }),
  thickBottom: (at) => ({ bottom: at.bottom ? THICK : undefined }),
  topBottom: (at) => ({ top: at.top ? THIN : undefined, bottom: at.bottom ? THIN : undefined }),
  topThickBottom: (at) => ({ top: at.top ? THIN : undefined, bottom: at.bottom ? THICK : undefined }),
  topDoubleBottom: (at) => ({ top: at.top ? THIN : undefined, bottom: at.bottom ? DOUBLE : undefined }),
};
const edges = (at, line) => ({ left: at.left ? line : undefined, right: at.right ? line : undefined, top: at.top ? line : undefined, bottom: at.bottom ? line : undefined });
export const BORDER_PRESETS = Object.keys(PRESETS);
const OPPOSITE = { left: 'right', right: 'left', top: 'bottom', bottom: 'top' };
const STEP = { left: [0, -1], right: [0, 1], top: [-1, 0], bottom: [1, 0] };
/** An automatic line colour: none given, indexed 64 (the system's) or auto. The menu's lines are automatic. */
const isAuto = (c) => !c || c.auto || c.indexed === 64 || (!c.rgb && c.theme == null && c.indexed == null);
const sameLine = (a, b) => !!a && !!b && a.style === b.style && isAuto(a.color) && isAuto(b.color);

/** A border preset of the Borders menu over the selection. A side drawn on the selection's edge takes the
 *  neighbour's facing side away, unless it is the same line (Excel keeps one line between two cells). */
export function setBorders(wb, si, sel, preset) {
  const sheet = wb.sheet(si);
  const out = new Map();
  const put = (r, c, sides) => {
    const s = styleAt(sheet, r, c);
    out.set(`${r}:${c}`, { r, c, s: wb.styleWith(s, { border: sides }) });
  };
  const cells = cellsOf(sheet, sel);
  const r1 = Math.max(...cells.map((x) => x.r));
  const c1 = Math.max(...cells.map((x) => x.c));
  for (const { r, c } of cells) {
    const at = { top: r === sel.r0, bottom: r === r1, left: c === sel.c0, right: c === c1 };
    const sides = Object.fromEntries(Object.entries(PRESETS[preset](at)).filter(([, v]) => v !== undefined));
    if (!Object.keys(sides).length) continue;
    put(r, c, sides);
    for (const [side, line] of Object.entries(sides)) {
      if (!at[side]) continue;
      const nr = r + STEP[side][0];
      const nc = c + STEP[side][1];
      if (nr < 0 || nc < 0) continue;
      const ns = styleAt(sheet, nr, nc);
      const facing = wb.styles.borders[xfOf(wb.styles, ns).borderId][OPPOSITE[side]];
      if (facing.style && !sameLine(facing, line)) {
        put(nr, nc, { [OPPOSITE[side]]: null });
      }
    }
  }
  return [...out.values()];
}

/** Horizontal or vertical alignment for every cell (Excel's buttons set it; pressing one again leaves it so). */
export const setAlign = (wb, si, sel, key, value) => formatEach(wb, si, sel, { align: { [key]: value } });

/** Where each section of a format code has literal text (quoted, escaped, in brackets, or after _ or *): a mask. */
function literalMask(code) {
  const lit = new Array(code.length).fill(false);
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') { const j = code.indexOf('"', i + 1); const end = j < 0 ? code.length - 1 : j; for (let k = i; k <= end; k++) lit[k] = true; i = end; }
    else if (ch === '\\' || ch === '_' || ch === '*') { lit[i] = true; if (i + 1 < code.length) lit[i + 1] = true; i++; }
    else if (ch === '[') { const j = code.indexOf(']', i); const end = j < 0 ? code.length - 1 : j; for (let k = i; k <= end; k++) lit[k] = true; i = end; }
  }
  return lit;
}

/** One section of a format code with a decimal place more (dir 1) or fewer (-1); unchanged when it shows no digits,
 *  or is a fraction. */
function stepSection(sec, dir) {
  const lit = literalMask(sec);
  const real = (i) => !lit[i];
  const at = [...sec].map((ch, i) => (real(i) ? ch : ''));
  const ePos = at.findIndex((ch) => ch === 'E' || ch === 'e');
  const end = ePos < 0 ? sec.length : ePos;
  const digits = [];
  for (let i = 0; i < end; i++) if (at[i] && '0#?'.includes(at[i])) digits.push(i);
  if (!digits.length || at.includes('/')) return sec;
  const dot = at.slice(0, end).indexOf('.');
  const last = digits.at(-1);
  if (dir > 0) {
    if (dot >= 0) return sec.slice(0, last + 1) + '0' + sec.slice(last + 1);
    return sec.slice(0, last + 1) + '.0' + sec.slice(last + 1);
  }
  if (dot < 0) return sec;
  if (last < dot) return sec.slice(0, dot) + sec.slice(dot + 1); // ("0." loses its point)
  const out = sec.slice(0, last) + sec.slice(last + 1);
  return digits.filter((i) => i > dot).length === 1 ? out.slice(0, dot) + out.slice(dot + 1) : out;
}

/** The format code Increase (1) or Decrease (-1) Decimal gives from `code`. General takes the decimals the cell
 *  shows (`shown`, its text; null when it isn't a number: then nothing changes). Dates and times don't change. */
export function stepDecimals(code, dir, shown) {
  if (!code || code === 'General') {
    if (shown == null || !/^-?\d*\.?\d+(E[+-]\d+)?$/i.test(shown)) return null;
    if (/E/i.test(shown)) return null;
    const d = Math.max(0, (shown.split('.')[1] || '').length + dir);
    return d ? `0.${'0'.repeat(d)}` : '0';
  }
  if (code === '@' || isDateFormat(code)) return null;
  const sections = [];
  const lit = literalMask(code);
  let from = 0;
  for (let i = 0; i <= code.length; i++) if (i === code.length || (code[i] === ';' && !lit[i])) { sections.push(code.slice(from, i)); from = i + 1; }
  const next = sections.map((s) => stepSection(s, dir)).join(';');
  return next === code ? null : next;
}

/** Increase or Decrease Decimal: every cell gets the code worked out from the active cell's (`shown`: the active
 *  cell's text as the grid shows it). [] when it doesn't change. */
export function stepDecimal(wb, si, sel, dir, shown) {
  const next = stepDecimals(activeLook(wb, si, sel).code, dir, shown);
  return next == null ? [] : formatEach(wb, si, sel, { numFmt: next });
}
