// The cells as the grid and print draw them: a cell's element (its text as shown, aligned, wrapped, running on into
// empty neighbours, indented, in its format's classes), merged ranges, and numbers too wide for their column as #s.
import { fontDescentPx } from './fonts.js';
import { inSelection } from './selection.js';
import { cellDisplay } from '../ooxml/numfmt.js';
import { editText } from '../ooxml/input.js';
import { ERROR_TEXT } from '../ooxml/errorcheck.js';
import { colChars } from '../ooxml/units.js';
import { MDW } from './geometry.js';
import { wrapLines } from './rowheights.js';
import { setVars, div } from './dom.js';
import { iconElement } from './cficons.js';

/** The merged range a cell is inside, if any. */
export const mergeAt = (st, r, c) => st.sheet.merges.find((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1) || null;

const NO_ALIGN = { horizontal: 'general', vertical: 'bottom', wrap: false, indent: 0, rotation: 0, shrink: false, readingOrder: 0 };
/** Excel's textRotation as degrees counter-clockwise: 1 to 90 up, 91 to 180 down (91 is -1); 255, stacked, is none. */
export const rotationDegrees = (r) => (!r || r === 255 ? 0 : r <= 90 ? r : 90 - r);
/** An indent level in px: 10.08 in print (Excel's PDF of test/fixtures/align.xlsx: one, two and three levels from
 *  either side), 9 on screen (not yet measured). */
const INDENT_PX = { screen: 9, print: 10.08 };

/** Where a cell's text sits across: general alignment puts numbers right, booleans and errors in the middle and text
 *  left; fill, justify and distributed are drawn left for now, centre-across-selection as centre. Turned or stacked,
 *  general goes by the angle whatever the value (Excel's PDF of test/fixtures/rotate.xlsx): 1 to 89 degrees up left,
 *  90 up and every angle down right, 90 down left; stacked letters centred. */
function horizontalOf(align, cell) {
  const h = align.horizontal;
  const r = align.rotation;
  if (h === 'general' && r === 255) return 'center';
  if (h === 'general' && r) return r >= 90 && r < 180 ? 'right' : 'left';
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
  // (A hidden formula of a protected sheet shows as nothing, a hidden constant as itself: st.hiddenAt, measured.)
  const held = formulas ? (cell.f != null && st.hiddenAt?.(r, c) ? '' : editText(cell, 'General')) : null;
  // (A conditional format's number format over the cell's own: ooxml/condformat.js.)
  const cfCode = st.cond?.get(`${r},${c}`)?.numFmt;
  const shown = formulas ? { text: held, parts: [held], color: null, fit: false } : cellDisplay(cell, colChars(screenW, MDW), cfCode || (look ? look.code : 'General'), st.date1904);
  // (A cell full of #s is centred, whatever its alignment: Excel's PDF of numfmt.xlsx, negative dates.)
  const h = shown.hashes ? 'center' : (shown.textFormat || (formulas && cell.f != null)) && align.horizontal === 'general' ? 'left' : horizontalOf(align, cell);
  // (Rotated, stacked or shrunk text stays in its own cell, as wrapped text does.)
  const turn = rotationDegrees(align.rotation);
  const stacked = align.rotation === 255;
  // (Without wrapping, a line break shows as nothing; wrapped, it breaks the line. Stacked, each character is a line of
  // its own, stepping by the font's default row height as wrapped lines do: Excel's PDF of rotate.xlsx, 19.36px a
  // letter in Calibri 11.)
  const clean = (t) => (stacked ? [...t.replace(/[\r\n]/g, '')].join('\n') : wrap ? t.replace(/\r/g, '') : t.replace(/[\r\n]/g, ''));
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
  // (A table style's font, under the cell's own and a conditional format's: ooxml/tablestyle.js tableLooks.)
  const tl = st.tables?.get(`${r},${c}`);
  if (tl) {
    if (tl.color && !cf?.color) { el.classList.add('oxe-tbl', 'oxe-tbl-color'); el.style.setProperty('--oxe-tbl-color', tl.color); }
    if (tl.bold !== undefined && cf?.bold === undefined) el.classList.add('oxe-tbl', tl.bold ? 'oxe-tbl-bold' : 'oxe-tbl-normal');
    if (tl.italic !== undefined && cf?.italic === undefined) el.classList.add('oxe-tbl', tl.italic ? 'oxe-tbl-italic' : 'oxe-tbl-upright');
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
  // (Rich text: each run in its own font, over the cell's, when the cell shows its text as it is.)
  const runs = !formulas && cell.runs && st.runStyle && shown.text === cell.v ? cell.runs : null;
  el.appendChild(runs ? runsSpan(runs, clean, st.runStyle, stacked) : textSpan(shown, clean));
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
function runsSpan(runs, clean, runStyle, stacked = false) {
  const span = document.createElement('span');
  span.className = 'oxe-tx oxe-runs';
  for (const [k, run] of runs.entries()) {
    const piece = document.createElement('span');
    // (Stacked: a line between one run's last letter and the next one's first, as between its own.)
    piece.textContent = (stacked && k ? '\n' : '') + clean(run.text);
    if (run.font) piece.style.cssText = runStyle(run.font);
    span.appendChild(piece);
  }
  return span;
}

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
/** How much taller than the line here Excel's turned text is, in em: its turned text keeps ~1.7px more from the
 *  cell's sides at 90 degrees, ~1.1px at 45 and ~0.5px at 15 (Excel's PDF of test/fixtures/rotate.xlsx), a box
 *  3.4px (0.23 em of Calibri 11) taller turned with the text. */
const TURN_EXTRA = 0.23;
/** Where Excel's turned text sits down its cell, in em: at the bottom lower by 0.18 em times the angle's cosine (its
 *  descent below the baseline, turned, is shorter than the line's here) and higher by 0.04 em (2.0px lower at 15
 *  degrees, 1.2 at 45, 0.6 higher at 90, in Calibri 11); at the top 0.14 em lower (1.5 to 2.5px) at any angle. */
const TURN_LOW = 0.18;
const TURN_UP = 0.04;
const TURN_TOP = 0.14;
/** Turned text kept inside its cell: the span turns about its middle, so it is moved by half the difference between
 *  its turned box and its own, towards the cell's inside (up for bottom-aligned text, right for left-aligned). All the
 *  sizes are read before any is set, so the page is laid out once. */
function placeTurned(container) {
  const moves = [...container.querySelectorAll('.oxe-cell.oxe-rot')].map((el) => {
    const span = el.firstElementChild;
    const a = (parseFloat(el.style.getPropertyValue('--oxe-rot')) * Math.PI) / 180;
    const w = span.offsetWidth;
    const h = span.offsetHeight;
    // (Across, Excel keeps text turned short of 90 degrees TURN_EXTRA em further from the cell's sides; turned 90
    // either way, and down the cell, it doesn't.)
    const deg = Math.abs(Math.round((a * 180) / Math.PI));
    const em = parseFloat(getComputedStyle(span).fontSize);
    const tall = h + (deg && deg < 90 ? TURN_EXTRA * em : 0);
    const dw = (w * Math.abs(Math.cos(a)) + tall * Math.abs(Math.sin(a)) - w) / 2;
    const dh = (w * Math.abs(Math.sin(a)) + h * Math.abs(Math.cos(a)) - h) / 2;
    const cl = el.classList;
    // (Down the cell: Excel's turned text sits TURN_LOW em lower at the bottom the flatter it is turned, and TURN_TOP
    // em lower at the top.)
    const dy = cl.contains('oxe-v-top') ? dh + TURN_TOP * em : cl.contains('oxe-v-center') ? 0 : -dh + (TURN_LOW * Math.abs(Math.cos(a)) - TURN_UP) * em;
    return [span, cl.contains('oxe-h-left') ? dw : cl.contains('oxe-h-right') ? -dw : 0, dy];
  });
  for (const [span, dx, dy] of moves) {
    span.style.setProperty('--oxe-rot-dx', `${dx.toFixed(2)}px`);
    span.style.setProperty('--oxe-rot-dy', `${dy.toFixed(2)}px`);
  }
}

/** A printer pixel at 600 dpi, in CSS pixels: the step of the sizes Excel prints shrunk text in. */
const PRINT_DOT = 96 / 600;
/** The room Excel leaves beside shrunk text in print, in CSS pixels. */
const SHRINK_SLACK = 1.7;
/**
 * The font size (CSS px) shrink-to-fit draws `wide` px of text at `px` in, for `room` px. In print, as Excel's PDF of
 * test/fixtures/rotate.xlsx shows: sized for the room less 1.7px, then down to a whole printer pixel at 600 dpi (Excel
 * printed four shrunk texts at 76, 67, 75 and 28 dots for 11pt's 92; a room of 148.7 to 149.7px fits all four, 151 was
 * here before). On screen, to the room itself (not measured).
 */
export function shrunkSize(px, room, wide, print) {
  if (!print) return (px * Math.max(room, 1)) / wide;
  return Math.max(1, Math.floor((px * Math.max(room - SHRINK_SLACK, 1)) / wide / PRINT_DOT)) * PRINT_DOT;
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
    if (el.dataset.oxeShrink) {
      // (The baseline kept where the full-size text's is, as Excel's PDF of rotate.xlsx shows it: the line box here
      // stays the cell font's, so a smaller font's baseline would rise by half its ascent less descent lost.)
      const px = shrunkSize(parseFloat(cs.fontSize), room, wide, !!el.closest('.oxe-page-cells'));
      const m = measure.measureText(text);
      shrinks.push([span, px, ((m.fontBoundingBoxAscent - m.fontBoundingBoxDescent) * (1 - px / parseFloat(cs.fontSize))) / 2]);
    }
    else jobs.push([span, Math.max(1, Math.floor(room / measure.measureText('#').width))]);
  }
  // (Shrink to fit: the font made small enough for the text to fit the cell's width: shrunkSize.)
  for (const [span, px, down] of shrinks) {
    span.style.fontSize = `${px.toFixed(2)}px`;
    if (down > 0) span.style.setProperty('--oxe-shrink-dy', `${down.toFixed(2)}px`);
  }
  for (const [span, n] of jobs) {
    span.classList.remove('oxe-parts', 'oxe-parts-fill');
    span.textContent = '#'.repeat(n);
    span.parentElement.classList.remove('oxe-h-left', 'oxe-h-right');
    span.parentElement.classList.add('oxe-h-center');
  }
}

/** The pixels AutoFit gives column c of the grid's sheet (`st`): its widest text as shown (a General number in full;
 *  wrapped text, the longest line it breaks into at the column's width now, as Excel's AutoFit (excel-sizes2.ps1)),
 *  measured in each cell's font, plus 9 of padding (docs/SIZES.md); merged cells left out, as Excel does; null when
 *  the column is empty. */
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
      const parts = look?.align?.wrap ? wrapLines(text, fitCtx.font, st.cols.size(c) - 4) : String(text).split('\n');
      for (const part of parts) widest = Math.max(widest, fitCtx.measureText(part.replace(/\r/g, '')).width);
    }
    return widest ? Math.ceil(widest) + 9 : null;
}
