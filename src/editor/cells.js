// The cells as the grid and print draw them: a cell's element (its text as shown, aligned, wrapped, running on into
// empty neighbours, indented, in its format's classes), merged ranges, and numbers too wide for their column as #s.
import { fontDescentPx } from './fonts.js';
import { inSelection } from './selection.js';
import { cellDisplay } from '../ooxml/numfmt.js';
import { editText } from '../ooxml/input.js';
import { colChars } from '../ooxml/units.js';
import { MDW } from './geometry.js';
import { setVars, div } from './dom.js';
import { iconElement } from './cficons.js';

/** The merged range a cell is inside, if any. */
export const mergeAt = (st, r, c) => st.sheet.merges.find((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1) || null;

/** What each error indicator says, as Excel's error button does. */
const ERROR_TEXT = {
  evalError: "This cell's formula results in an error.",
  twoDigitTextYear: 'This cell contains a date string represented with only 2 digits for the year.',
  numberStoredAsText: 'The number in this cell is formatted as text or preceded by an apostrophe.',
  formula: 'The formula in this cell differs from the formulas in this area of the spreadsheet.',
  formulaRange: 'The formula in this cell refers to a range that has additional numbers adjacent to it.',
  unlockedFormula: 'This cell contains a formula and is not locked to protect it from being changed inadvertently.',
};

const NO_ALIGN = { horizontal: 'general', vertical: 'bottom', wrap: false, indent: 0, rotation: 0, shrink: false, readingOrder: 0 };
/** Excel's textRotation as degrees counter-clockwise: 1 to 90 up, 91 to 180 down (91 is -1); 255, stacked, is none. */
export const rotationDegrees = (r) => (!r || r === 255 ? 0 : r <= 90 ? r : 90 - r);
/** An indent level in px: 10.08 in print (Excel's PDF of test/fixtures/align.xlsx: one, two and three levels from
 *  either side), 9 on screen (not yet measured). */
const INDENT_PX = { screen: 9, print: 10.08 };

/** Where a cell's text sits across: general alignment puts numbers right, booleans and errors in the middle and text
 *  left; fill, justify and distributed are drawn left for now, centre-across-selection as centre. */
function horizontalOf(align, cell) {
  const h = align.horizontal;
  if (h === 'general') return cell.t === 'n' ? 'right' : cell.t === 'b' || cell.t === 'e' ? 'center' : 'left';
  if (h === 'center' || h === 'centerContinuous') return 'center';
  return h === 'right' ? 'right' : 'left';
}

/** The columns either side of c that text may run on into: up to the nearest cell with something in it, a merged
 *  range or the edge of what is drawn. Returns [first, last + 1). */
function runOn(st, r, c, line, win) {
  let end = c + 1;
  while (end <= win.c1 + 1 && end < st.cols.count && !(line.get(end)?.v != null) && !mergeAt(st, r, end)) end++;
  let begin = c;
  while (begin - 1 >= Math.max(0, win.c0 - 1) && !(line.get(begin - 1)?.v != null) && !mergeAt(st, r, begin - 1)) begin--;
  return [begin, end];
}

/** One cell's element: its text in an inner span (which carries the line height and the baseline shifts), placed by
 *  the cell format's alignment and indent; text runs on over empty neighbours (left text to the right, right text to
 *  the left, centred text both ways, as far as the nearer side allows), wrapped text and numbers never do. */
export function cellElement(st, r, c, cell, line, win, rowTop) {
  const merge = mergeAt(st, r, c);
  const look = st.deco ? st.deco(cell.s || 0) : null;
  const align = (look && look.align) || NO_ALIGN;
  // (Show Formulas: a formula is shown as text, and runs on as text does.)
  const formulas = st.sheet.showFormulas;
  const isText = cell.t === 's' || cell.t === 'str' || (formulas && cell.f != null);
  const wrap = align.wrap && isText;
  // (A General number gets the characters that fit its cell on screen, in print too: st.screenCols when printing.)
  const screen = st.screenCols || st.cols;
  const screenW = merge ? screen.start(merge.c1 + 1) - screen.start(c) : screen.size(c);
  // (Show Formulas: what the cell holds, unformatted, as Excel shows it: =A1*2, 1234.5 for 1,234.50, 45292 for a date,
  // 0.25 for 25%, 007 for '007.)
  const held = formulas ? editText(cell, 'General') : null;
  const shown = formulas ? { text: held, parts: [held], color: null, fit: false } : cellDisplay(cell, colChars(screenW, MDW), look ? look.code : 'General', st.date1904);
  // (A cell full of #s is centred, whatever its alignment: Excel's PDF of numfmt.xlsx, negative dates.)
  const h = shown.hashes ? 'center' : (shown.textFormat || (formulas && cell.f != null)) && align.horizontal === 'general' ? 'left' : horizontalOf(align, cell);
  // (Without wrapping, a line break shows as nothing; wrapped, it breaks the line.)
  const clean = (t) => (wrap ? t.replace(/\r/g, '') : t.replace(/[\r\n]/g, ''));
  // (Rotated, stacked or shrunk text stays in its own cell, as wrapped text does.)
  const turn = rotationDegrees(align.rotation);
  const stacked = align.rotation === 255;
  const shrink = align.shrink && !wrap && cell.v != null;
  const el = div(`oxe-cell oxe-h-${h} oxe-v-${align.vertical === 'top' ? 'top' : align.vertical === 'center' ? 'center' : 'bottom'}${wrap ? ' oxe-wrap' : ''}${turn ? ' oxe-rot' : ''}${stacked ? ' oxe-stack' : ''}`, 'gridcell');
  if (align.readingOrder === 2) el.dir = 'rtl';
  if (st.classOf) el.classList.add(st.classOf(cell.s || 0));
  // (Superscript or subscript: raised or lowered, cellstyles.js.)
  if (look?.sup) el.classList.add(look.sup);
  el.setAttribute('aria-colindex', String(c + 2));
  if (st.idBase) el.id = `${st.idBase}-${r}-${c}`;
  // (A conditional format's font: its colour and weight over the cell's own, ooxml/condformat.js.)
  const cf = st.cond?.get(`${r},${c}`);
  if (cf) {
    if (cf.color) { el.classList.add('oxe-cf-color'); el.style.setProperty('--oxe-cf-color', cf.color); }
    if (cf.bold !== undefined) el.classList.add(cf.bold ? 'oxe-cf-bold' : 'oxe-cf-normal');
    if (cf.italic !== undefined) el.classList.add(cf.italic ? 'oxe-cf-italic' : 'oxe-cf-upright');
    if (cf.strike) el.classList.add('oxe-cf-strike');
    if (cf.hideValue) el.classList.add('oxe-cf-novalue');
  }
  // (An icon set's icon at the cell's left: put in after the text, which stays the cell's first child.)
  const cfIcon = cf?.icon ? iconElement(cf.icon.set, cf.icon.index) : null;
  if (cfIcon) el.classList.add('oxe-cf-iconed');
  // (A note or comment thread: heard as the cell's description, from notemarks.js's texts.)
  if (st.idBase && st.notes?.has(`${r},${c}`)) el.setAttribute('aria-describedby', `${st.idBase}-note-${r}-${c}`);
  if (st.sel && inSelection(st.sel, r, c)) el.setAttribute('aria-selected', 'true');
  // (Excel's error indicator, the green triangle in the top-left corner, and what it says: ooxml/errorcheck.js.)
  const flags = st.errors?.get(`${r},${c}`);
  if (flags) { el.classList.add('oxe-errmark'); el.setAttribute('aria-description', flags.map((f) => ERROR_TEXT[f]).join(' ')); }
  let x = st.cols.start(c);
  let w = st.cols.size(c);
  let hgt = st.rows.size(r);
  if (merge) {
    el.classList.add('oxe-merged');
    el.setAttribute('aria-colspan', String(merge.c1 - merge.c0 + 1));
    el.setAttribute('aria-rowspan', String(merge.r1 - merge.r0 + 1));
    w = st.cols.start(merge.c1 + 1) - x;
    hgt = st.rows.start(merge.r1 + 1) - st.rows.start(r);
  } else if (isText && !wrap && !turn && !stacked && !shrink) {
    const [begin, end] = runOn(st, r, c, line, win);
    const left = st.cols.start(begin);
    const right = st.cols.start(end);
    if (h === 'left') w = right - x;
    else if (h === 'right') { w = x + w - left; x = left; }
    else {
      const centre = x + w / 2;
      const half = Math.min(centre - left, right - centre);
      x = centre - half;
      w = 2 * half;
    }
  }
  const vars = { '--oxe-x': x, '--oxe-y': st.rows.start(r) - rowTop, '--oxe-w': w, '--oxe-h': hgt };
  // Bottom-aligned text sits on its row's baseline, which the row's tallest font sets (rowheights.js rowDescents).
  if (!wrap && align.vertical !== 'top' && align.vertical !== 'center' && look && st.rowDesc) {
    const lift = (st.rowDesc.get(r) ?? 0) - fontDescentPx(look.font.family, look.font.sizePx);
    if (lift > 0.01) vars['--oxe-row-shift'] = -lift;
  }
  if (align.indent && h !== 'center') {
    el.classList.add(h === 'right' ? 'oxe-indent-r' : 'oxe-indent-l');
    vars['--oxe-indent'] = align.indent * (st.print ? INDENT_PX.print : INDENT_PX.screen);
  }
  el.appendChild(textSpan(shown, clean));
  if (cfIcon) el.appendChild(cfIcon);
  const color = shown.color && st.color ? st.color(shown.color) : null;
  if (color) {
    el.classList.add('oxe-nf-color');
    el.style.setProperty('--oxe-nf-color', color);
  }
  // (A formatted number wider than its cell shows #s instead; fitNumbers checks once the cell is in the page.)
  // (Shrink to fit makes the font smaller instead: fitNumbers again.)
  if (shrink) el.dataset.oxeShrink = '1';
  else if (shown.fit) el.dataset.oxeFit = '1';
  setVars(el, vars);
  // (An angle, not a length: set as one, setVars giving px.)
  if (turn) el.style.setProperty('--oxe-rot', `${turn}deg`);
  return el;
}

/** The text span: plain text, or for formats with padding (_x: a space as wide as x, drawn as a hidden x) and fill
 *  (*x: x repeated across the room left in the cell) the pieces in a row. */
function textSpan(shown, clean) {
  const span = document.createElement('span');
  span.className = 'oxe-tx';
  if (shown.parts.every((p) => typeof p === 'string')) {
    span.textContent = clean(shown.text);
    return span;
  }
  span.classList.add('oxe-parts');
  if (shown.parts.some((p) => p && p.fill)) span.classList.add('oxe-parts-fill');
  for (const p of shown.parts) {
    if (typeof p === 'string') span.append(clean(p));
    else {
      const piece = document.createElement('span');
      piece.className = p.pad ? 'oxe-pad' : 'oxe-repeat';
      piece.textContent = p.pad || p.fill.repeat(200);
      span.appendChild(piece);
    }
  }
  return span;
}

let measure = null;
/** Turned text kept inside its cell: the span turns about its middle, so it is moved by half the difference between
 *  its turned box and its own, towards the cell's inside (up for bottom-aligned text, right for left-aligned). All the
 *  sizes are read before any is set, so the page is laid out once. */
function placeTurned(container) {
  const moves = [...container.querySelectorAll('.oxe-cell.oxe-rot')].map((el) => {
    const span = el.firstElementChild;
    const a = (parseFloat(el.style.getPropertyValue('--oxe-rot')) * Math.PI) / 180;
    const w = span.offsetWidth;
    const h = span.offsetHeight;
    const dw = (w * Math.abs(Math.cos(a)) + h * Math.abs(Math.sin(a)) - w) / 2;
    const dh = (w * Math.abs(Math.sin(a)) + h * Math.abs(Math.cos(a)) - h) / 2;
    const cl = el.classList;
    return [span, cl.contains('oxe-h-left') ? dw : cl.contains('oxe-h-right') ? -dw : 0, cl.contains('oxe-v-top') ? dh : cl.contains('oxe-v-center') ? 0 : -dh];
  });
  for (const [span, dx, dy] of moves) {
    span.style.setProperty('--oxe-rot-dx', `${dx.toFixed(2)}px`);
    span.style.setProperty('--oxe-rot-dy', `${dy.toFixed(2)}px`);
  }
}

/** The canvas AutoFit measures text with. */
let fitCtx = null;
/** Replace formatted numbers wider than their cells with #s, as Excel does (as many as fit), and make shrink-to-fit
 *  text smaller until it fits. Run once the cells are in
 *  the page, so one layout serves them all; the width is measured in the cell's own font. */
export function fitNumbers(container) {
  placeTurned(container);
  const cells = container.querySelectorAll('.oxe-cell[data-oxe-fit], .oxe-cell[data-oxe-shrink]');
  if (!cells.length) return;
  measure = measure || document.createElement('canvas').getContext('2d');
  const jobs = [];
  const shrinks = [];
  for (const el of cells) {
    const span = el.firstElementChild;
    const cs = getComputedStyle(span);
    const box = getComputedStyle(el);
    const room = parseFloat(el.style.getPropertyValue('--oxe-w')) - parseFloat(box.paddingLeft) - parseFloat(box.paddingRight);
    measure.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
    // (Padding pieces count; a fill shrinks to nothing.)
    const text = [...span.childNodes].map((n) => (n.classList && n.classList.contains('oxe-repeat') ? '' : n.textContent)).join('');
    const wide = measure.measureText(text).width;
    if (wide <= room + 0.5) continue;
    if (el.dataset.oxeShrink) shrinks.push([span, (parseFloat(cs.fontSize) * Math.max(room, 1)) / wide]);
    else jobs.push([span, Math.max(1, Math.floor(room / measure.measureText('#').width))]);
  }
  // (Shrink to fit: the font made small enough for the text to fit the cell's width. Not yet measured against the
  // sizes Excel picks.)
  for (const [span, px] of shrinks) span.style.fontSize = `${px.toFixed(2)}px`;
  for (const [span, n] of jobs) {
    span.classList.remove('oxe-parts', 'oxe-parts-fill');
    span.textContent = '#'.repeat(n);
    span.parentElement.classList.remove('oxe-h-left', 'oxe-h-right');
    span.parentElement.classList.add('oxe-h-center');
  }
}

/** The pixels AutoFit gives column c of the grid's sheet (`st`): its widest text as shown (a General number in full,
 *  the longest line of wrapped text), measured in each cell's font, plus 9 of padding (docs/SIZES.md); null when the
 *  column is empty. */
export function columnFitWidth(st, c) {
    if (!fitCtx) fitCtx = document.createElement('canvas').getContext('2d');
    let widest = 0;
    for (const [r, line] of st.sheet.cells) {
      const cell = line.get(c);
      if (!cell || cell.v == null || mergeAt(st, r, c)) continue;
      const look = st.deco ? st.deco(cell.s || 0) : null;
      const text = cellDisplay(cell, Infinity, look ? look.code : 'General', st.date1904).text;
      if (!text) continue;
      fitCtx.font = look ? look.font.css : '14.6667px sans-serif';
      for (const part of String(text).split('\n')) widest = Math.max(widest, fitCtx.measureText(part.replace(/\r/g, '')).width);
    }
    return widest ? Math.ceil(widest) + 9 : null;
}
