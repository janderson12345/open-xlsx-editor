// Pictures, shapes, icons and groups drawn over the cells (ooxml/drawobjects.js read them, drawpaint.js says what each
// paints): each object a box where its anchor puts it on the sheet, holding an SVG of its paths and pictures and its
// text as HTML, so that it follows the rows' and columns' sizes, hidden lines, frozen panes and the zoom as the cells do.
// The same boxes are printed (print.js) and exported (pdf.js reads what each paints from the box). With the charts
// (chartview.js) they are drawn in the drawing's order, later ones over earlier ones.
import { paintObject, isLineObject } from './drawpaint.js';
import { anchorRect } from '../ooxml/drawobjects.js';
import { pathData } from '../ooxml/presets.js';
import { chartBoxes } from './chartview.js';
import { fontStack } from './fonts.js';

const SVG = 'http://www.w3.org/2000/svg';
const svgEl = (name, attrs = {}) => {
  const x = document.createElementNS(SVG, name);
  for (const [k, v] of Object.entries(attrs)) if (v != null) x.setAttribute(k, String(v));
  return x;
};
let clips = 0;
const n2 = (v) => +v.toFixed(3);

/** An object's text box (drawpaint.js's 'text' op) as HTML, at its matrix from the box's top left (x0, y0). */
function textElement(op, x0, y0) {
  const t = document.createElement('div');
  t.className = `oxe-obj-text oxe-obj-text-${op.body.anchor}`;
  const m = op.m;
  t.style.width = `${op.w}px`;
  t.style.height = `${op.h}px`;
  t.style.transform = `matrix(${n2(m[0])}, ${n2(m[1])}, ${n2(m[2])}, ${n2(m[3])}, ${n2(m[4] - x0)}, ${n2(m[5] - y0)})`;
  if (!op.body.wrap) t.style.whiteSpace = 'pre';
  for (const p of op.body.paras) {
    const para = document.createElement('div');
    para.className = 'oxe-obj-para';
    para.style.textAlign = { ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[p.align] || 'left';
    const runs = p.runs.length ? p.runs : [{ text: '', ...p.empty }];
    // (An empty paragraph keeps its line, in its own size.)
    // (Lines the font's own height apart, as Excel sets them at 100% spacing (its PDF of drawings.xlsx: 18px for 11pt
    // Calibri), in the largest run's font, so that the line's strut is that font's.)
    const big = runs.reduce((a, r) => (r.size > a.size ? r : a), runs[0]);
    para.style.fontFamily = fontStack(big.family);
    para.style.fontSize = `${n2(big.size * 4 / 3)}px`;
    para.style.lineHeight = 'normal';
    for (const r of runs) {
      if (r.text === '\n') { para.append(document.createElement('br')); continue; }
      const s = document.createElement('span');
      s.textContent = r.text || (p.runs.length ? '' : '​');
      s.style.fontFamily = fontStack(r.family);
      s.style.fontSize = `${n2(r.size * 4 / 3)}px`;
      if (r.bold) s.style.fontWeight = 'bold';
      if (r.italic) s.style.fontStyle = 'italic';
      const deco = [r.underline && 'underline', r.strike && 'line-through'].filter(Boolean).join(' ');
      if (deco) s.style.textDecoration = deco;
      s.style.color = r.color.css;
      if (r.color.alpha < 1) s.style.opacity = String(r.color.alpha);
      para.append(s);
    }
    t.append(para);
  }
  return t;
}

/**
 * Object `obj` at `rect` (sheet pixels) as a box: `.oxe-obj` at the rect (at least `hit` pixels each way, so that a
 * line can be clicked), its paint (`_oxePaint`, drawpaint.js's list) in an SVG whose user space is the sheet's.
 */
export function objectElement(obj, rect, theme, { hit = 0 } = {}) {
  const ops = paintObject(obj, rect, theme);
  const box = { ...rect };
  if (box.w < hit) { box.x -= (hit - box.w) / 2; box.w = hit; }
  if (box.h < hit) { box.y -= (hit - box.h) / 2; box.h = hit; }
  const div = document.createElement('div');
  div.className = `oxe-obj oxe-obj-${obj.kind}`;
  div.dataset.oxeOrd = String(obj.k);
  div.style.setProperty('--oxe-x', `${box.x}px`);
  div.style.setProperty('--oxe-y', `${box.y}px`);
  div.style.setProperty('--oxe-w', `${box.w}px`);
  div.style.setProperty('--oxe-h', `${box.h}px`);
  const svg = svgEl('svg', { width: n2(box.w), height: n2(box.h), viewBox: `${n2(box.x)} ${n2(box.y)} ${n2(box.w) || 1} ${n2(box.h) || 1}`, overflow: 'visible' });
  for (const op of ops) {
    if (op.t === 'path') {
      const line = op.line;
      svg.append(svgEl('path', {
        d: pathData(op.cmds), fill: op.fill ? op.fill.css : 'none', 'fill-opacity': op.fill && op.fill.alpha < 1 ? op.fill.alpha : null,
        stroke: line ? line.css : null, 'stroke-opacity': line && line.alpha < 1 ? line.alpha : null, 'stroke-width': line ? n2(line.width) : null,
        'stroke-dasharray': line?.dash ? line.dash.map((d) => n2(d * line.width)).join(' ') : null, 'stroke-linecap': line && line.cap !== 'butt' ? line.cap : null, 'stroke-linejoin': line ? 'round' : null,
      }));
    } else if (op.t === 'image') {
      const id = `oxe-clip-${++clips}`;
      const cp = svgEl('clipPath', { id, clipPathUnits: 'userSpaceOnUse' });
      cp.append(svgEl('path', { d: pathData(op.clip) }));
      const g = svgEl('g', { 'clip-path': `url(#${id})` });
      const m = op.m;
      // (An SVG icon draws from its own picture; the PNG beside it is for what can't draw SVG.)
      g.append(svgEl('image', { href: (op.svg || op.src).url, x: 0, y: 0, width: 1, height: 1, preserveAspectRatio: 'none', transform: `matrix(${m.map(n2).join(' ')})` }));
      svg.append(cp, g);
    }
  }
  div.append(svg);
  for (const op of ops) if (op.t === 'text') div.append(textElement(op, box.x, box.y));
  div._oxePaint = ops;
  div._oxeBox = box;
  return div;
}

/** Whether an object at `rect` is drawn: a line may be 0 one way, anything else not (as Excel hides one whose rows or
 *  columns are all hidden). */
const shown = (obj, r) => (isLineObject(obj) ? r.w > 0 || r.h > 0 : r.w >= 1 && r.h >= 1);

/** The object boxes of `st.objects` that reach into window `win` (rows and columns). */
export function objectBoxes(st, win, fresh = false) {
  const out = [];
  st.objectCache ??= new Map();
  for (const obj of st.objects || []) {
    if (obj.hidden || !obj.item) continue;
    const r = anchorRect(obj.anchor, st.rows, st.cols, obj.item);
    if (!shown(obj, r)) continue;
    const x1 = st.cols.start(win.c0);
    const x2 = st.cols.start(win.c1 + 1);
    const y1 = st.rows.start(win.r0);
    const y2 = st.rows.start(win.r1 + 1);
    // (Rotated, an object may reach past its frame: half its diagonal from its middle.)
    const reach = Math.hypot(r.w, r.h) / 2;
    const cx = r.x + r.w / 2;
    const cy = r.y + r.h / 2;
    if (cx + reach < x1 || cx - reach > x2 || cy + reach < y1 || cy - reach > y2) continue;
    const key = `${r.x},${r.y},${r.w},${r.h}`;
    let box = fresh ? null : st.objectCache.get(obj);
    if (!box || box.key !== key) {
      box = { key, el: objectElement(obj, r, st.theme, { hit: 6 }) };
      if (!fresh) st.objectCache.set(obj, box);
    }
    out.push({ ord: obj.k, el: box.el });
  }
  return out;
}

/** Draw the charts and objects that reach into window `win` into `st.chartLayer`, in the drawing's order. */
export function drawDrawings(st, win) {
  const layer = st.chartLayer;
  if (!layer) return;
  const all = [...chartBoxes(st, win), ...objectBoxes(st, win)].sort((a, b) => a.ord - b.ord);
  layer.replaceChildren(...all.map((b) => b.el));
  st.onChartsDrawn?.();
}

/** A frozen pane's charts and objects (gridlayers.js drawPanes): drawn again for its window, so that those in the rows
 *  or columns kept in place stay there. */
export function drawPaneDrawings(sub, layer, win) {
  if (!layer) return;
  if (!win) { layer.replaceChildren(); return; }
  const all = [...chartBoxes(sub, win, true), ...objectBoxes(sub, win, true)].sort((a, b) => a.ord - b.ord);
  layer.replaceChildren(...all.map((b) => b.el));
}

/** What a screen reader is given for an object: a figure with its kind, name, alternative text and its text. */
export function objectFigure(obj) {
  const fig = document.createElement('figure');
  const cap = document.createElement('figcaption');
  const kind = { picture: 'Picture', shape: 'Shape', group: 'Group' }[obj.kind] || 'Object';
  cap.textContent = `${kind}: ${obj.name || kind}${obj.descr ? `. ${obj.descr}` : ''}`;
  fig.append(cap);
  const texts = [];
  const walk = (it) => { if (it.text) texts.push(it.text.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n')); for (const c of it.children || []) walk(c); };
  if (obj.item) walk(obj.item);
  if (texts.length) { const p = document.createElement('p'); p.textContent = texts.join('\n'); fig.append(p); }
  return fig;
}
