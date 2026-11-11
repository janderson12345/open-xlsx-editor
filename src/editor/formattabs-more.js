// The Format Cells box's Border, Fill and Protection tabs (formatbox.js), built like formattabs.js's: on the active
// cell's look, marking in `changed` what was touched, get() giving formatops.js formatCells's change for those only.
import { make, check, select, labelled, colorField, LINES, PATTERNS, SHADINGS } from './formatbox-kit.js';

/** The Border tab's buttons, in Excel's places: the edges and inside lines of the selection, and the two diagonals. */
const SIDES = [['top', 'Top border'], ['horizontal', 'Inside horizontal border'], ['bottom', 'Bottom border'], ['left', 'Left border'],
  ['vertical', 'Inside vertical border'], ['right', 'Right border'], ['diagonalUp', 'Diagonal up border'], ['diagonalDown', 'Diagonal down border']];
const OUTLINE = ['top', 'bottom', 'left', 'right'];
const INSIDE = ['horizontal', 'vertical'];

/** The Border tab: a line style and colour, then the presets (None, Outline, Inside) or one side at a time; a side
 *  pressed with the line it has is taken away. Inside lines only for more than one cell. */
export function borderTab(look, changed) {
  const style = make('select', { className: 'oxe-dialog-list', size: 7 }, ...LINES.map(([v, l]) => make('option', { value: v, textContent: l, selected: v === 'thin' })));
  const color = colorField(null, 'Automatic');
  const b = look.border;
  // (What each side has now: the active cell's lines, when one cell is selected; the inside none.)
  const now = {};
  for (const k of OUTLINE) now[k] = !look.many && b[k]?.style ? { style: b[k].style, color: b[k].color } : null;
  for (const k of INSIDE) now[k] = null;
  const diag = b.diagonal?.style ? { style: b.diagonal.style, color: b.diagonal.color } : null;
  now.diagonalUp = !look.many && b.diagonalUp ? diag : null;
  now.diagonalDown = !look.many && b.diagonalDown ? diag : null;
  const line = () => ({ style: style.value, color: color.value() || { indexed: 64 } });
  const same = (x, y) => !!x && !!y && x.style === y.style && JSON.stringify(x.color || null) === JSON.stringify(y.color || null);
  const preview = make('div', { className: 'oxe-fc-border-preview', role: 'img' }, make('span', { textContent: 'Text' }));
  const buttons = {};
  const paint = () => {
    for (const [k, btn] of Object.entries(buttons)) btn.setAttribute('aria-pressed', String(!!now[k]));
    const css = (x) => (x ? `${x.style === 'double' ? 3 : /medium|thick/.test(x.style) ? 2 : 1}px ${x.style === 'double' ? 'double' : /dash|dot|hair/i.test(x.style) ? 'dashed' : 'solid'} currentColor` : '0');
    for (const k of OUTLINE) preview.style[`border${k[0].toUpperCase()}${k.slice(1)}`] = css(now[k]);
    preview.classList.toggle('oxe-fc-diag-up', !!now.diagonalUp);
    preview.classList.toggle('oxe-fc-diag-down', !!now.diagonalDown);
    preview.setAttribute('aria-label', `Border preview: ${SIDES.filter(([k]) => now[k]).map(([, l]) => l.toLowerCase()).join(', ') || 'no border'}`);
  };
  const set = (keys, value) => { for (const k of keys) { now[k] = value; changed[k] = true; } paint(); };
  for (const [k, label] of SIDES) {
    const inside = INSIDE.includes(k);
    buttons[k] = make('button', { type: 'button', className: `oxe-fc-side oxe-fc-side-${k}`, title: label, disabled: inside && !look.many });
    buttons[k].setAttribute('aria-label', label);
    buttons[k].addEventListener('click', () => set([k], same(now[k], line()) ? null : line()));
  }
  const presets = make('div', { className: 'oxe-fc-presets' },
    make('button', { type: 'button', className: 'oxe-dialog-button', textContent: 'None', onclick: () => set([...OUTLINE, ...INSIDE, 'diagonalUp', 'diagonalDown'], null) }),
    make('button', { type: 'button', className: 'oxe-dialog-button', textContent: 'Outline', onclick: () => set(OUTLINE, line()) }),
    make('button', { type: 'button', className: 'oxe-dialog-button', textContent: 'Inside', disabled: !look.many, onclick: () => set(INSIDE, line()) }));
  paint();
  const get = () => {
    const outline = {};
    const inside = {};
    for (const k of OUTLINE) if (changed[k]) outline[k] = now[k];
    for (const k of INSIDE) if (changed[k]) inside[k] = now[k];
    const borders = { outline, inside };
    if (changed.diagonalUp || changed.diagonalDown) {
      borders.diagonalUp = !!now.diagonalUp;
      borders.diagonalDown = !!now.diagonalDown;
      borders.diagonal = now.diagonalUp || now.diagonalDown || null;
    }
    return { borders };
  };
  const el = make('div', { className: 'oxe-fc-cols' },
    make('fieldset', {}, make('legend', { textContent: 'Line' }), labelled('Style:', style), labelled('Color:', color.el)),
    make('fieldset', {}, make('legend', { textContent: 'Border' }), presets,
      make('div', { className: 'oxe-fc-border-pad' }, ...SIDES.map(([k]) => buttons[k]), preview)));
  return { el, get };
}

/** The Fill tab: the background colour (No Color for none), a pattern's style and colour over it, or Fill Effects'
 *  two-colour gradient. */
export function fillTab(look, changed) {
  const f = look.fill;
  const gradient = f.pattern === 'gradient' ? f.gradient : null;
  const solid = f.pattern === 'solid';
  const bg = colorField(solid ? f.fg : f.pattern !== 'none' && !gradient ? (f.bg?.indexed === 64 ? null : f.bg) : null, 'No Color', () => { changed.bg = true; paint(); });
  const patternColor = colorField(!solid && !gradient && f.pattern !== 'none' ? f.fg : null, 'Automatic', () => { changed.pattern = true; paint(); });
  const pattern = select(PATTERNS, solid || gradient ? 'none' : f.pattern);
  const effects = check('Gradient (Fill Effects)', !!gradient);
  const color1 = colorField(gradient?.stops[0]?.color || { rgb: 'FFFFFFFF' }, 'White', () => { changed.gradient = true; paint(); });
  const color2 = colorField(gradient?.stops.at(-1)?.color || { theme: 4 }, 'White', () => { changed.gradient = true; paint(); });
  const shadeOf = (g) => (!g ? 'horizontal' : g.type === 'path' ? 'center' : (SHADINGS.find(([, , x]) => x.degree === g.degree) || SHADINGS[0])[0]);
  const shading = select(SHADINGS.map(([v, l]) => [v, l]), shadeOf(gradient));
  const sample = make('div', { className: 'oxe-fc-fill-sample', role: 'img', 'aria-label': 'Sample' });
  const paint = () => {
    for (const x of [color1.el, color2.el, shading]) x.closest('.oxe-fc-field')?.toggleAttribute('hidden', !effects.input.checked);
    const css = (c) => (c && look.color ? look.color(c) : null);
    sample.style.background = effects.input.checked ? `linear-gradient(90deg, ${css(color1.value()) || '#ffffff'}, ${css(color2.value()) || '#ffffff'})`
      : pattern.value === 'solid' ? css(patternColor.value()) || '#000000' : css(bg.value()) || 'transparent';
  };
  pattern.addEventListener('change', () => { changed.pattern = true; paint(); });
  effects.input.addEventListener('change', () => { changed.gradient = true; paint(); });
  shading.addEventListener('change', () => { changed.gradient = true; paint(); });
  const noColor = make('button', { type: 'button', className: 'oxe-dialog-button', textContent: 'No Color', onclick: () => { bg.sel.value = 'none'; pattern.value = 'none'; effects.input.checked = false; changed.bg = true; changed.pattern = true; changed.gradient = true; paint(); } });
  const get = () => {
    if (effects.input.checked) {
      const shape = SHADINGS.find(([v]) => v === shading.value)[2];
      return { fill: { gradient: { ...shape, stops: [{ position: 0, color: color1.value() || { theme: 0 } }, { position: 1, color: color2.value() || { theme: 0 } }] } } };
    }
    const back = bg.value();
    const p = pattern.value;
    if (p === 'none') return { fill: back };
    // (A pattern's own colour is its foreground, the cell's colour behind it, as Excel saves them.)
    if (p === 'solid') return { fill: patternColor.value() || back || { indexed: 64 } };
    return { fill: { pattern: p, fg: patternColor.value(), bg: back } };
  };
  const el = make('div', { className: 'oxe-fc-cols' },
    make('fieldset', {}, make('legend', { textContent: 'Background Color' }), labelled('Color:', bg.el), noColor, effects.el,
      labelled('Color 1:', color1.el), labelled('Color 2:', color2.el), labelled('Shading styles:', shading)),
    make('fieldset', {}, make('legend', { textContent: 'Pattern' }), labelled('Pattern Color:', patternColor.el), labelled('Pattern Style:', pattern), sample));
  paint();
  return { el, get };
}

/** The Protection tab: Locked and Hidden, which count once the sheet is protected (Review > Protect Sheet). */
export function protectionTab(look, changed) {
  const p = look.xf.protection || { locked: true, hidden: false };
  const locked = check('Locked', p.locked);
  const hidden = check('Hidden', p.hidden);
  locked.input.addEventListener('change', () => { changed.locked = true; });
  hidden.input.addEventListener('change', () => { changed.hidden = true; });
  const get = () => {
    const protection = {};
    if (changed.locked) protection.locked = locked.input.checked;
    if (changed.hidden) protection.hidden = hidden.input.checked;
    return { protection };
  };
  const el = make('div', {}, locked.el, hidden.el,
    make('p', { className: 'oxe-dialog-text', textContent: 'Locking cells or hiding formulas has no effect until you protect the worksheet (Review tab, Protect Sheet button).' }));
  return { el, get };
}
