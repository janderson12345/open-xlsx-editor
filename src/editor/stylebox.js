// Excel's Style box, for Home > Cell Styles > New Cell Style... and a style's Modify...: the style's name, what it
// includes (Number, Alignment, Font, Border, Fill, Protection, each with a line saying what it is) and Format..., the
// Format Cells box (formatbox.js) for the style's own format. It resolves to { name, includes, change } (change: what
// Format... changed, the Format Cells box's change, merged over the times it was opened), or null when cancelled.
import { modal, button } from './dialogs.js';
import { make, labelled, check } from './formatbox-kit.js';

/** The Style box's parts, in its order, and their labels. */
const PARTS = [['number', 'Number'], ['alignment', 'Alignment'], ['font', 'Font'], ['border', 'Border'], ['fill', 'Fill'], ['protection', 'Protection']];
const H = { general: 'General', left: 'Left', center: 'Center', right: 'Right', fill: 'Fill', justify: 'Justify', centerContinuous: 'Center Across Selection', distributed: 'Distributed' };
const V = { top: 'Top', center: 'Center', bottom: 'Bottom', justify: 'Justify', distributed: 'Distributed' };

/** What each part of a format is, in the Style box's words: `look` as formatbox.js's (xf, font, fontName, fill,
 *  border, code). */
export function describeParts(look) {
  const { xf, font, fill, border } = look;
  const a = xf.align || {};
  const sides = border ? ['left', 'right', 'top', 'bottom'].filter((k) => border[k]?.style) : [];
  return {
    number: look.code || 'General',
    alignment: `${H[a.horizontal] || 'General'}, ${V[a.vertical] || 'Bottom'} Aligned${a.wrap ? ', Wrap Text' : ''}${a.rotation ? `, ${a.rotation === 255 ? 'Vertical' : `${a.rotation > 90 ? 90 - a.rotation : a.rotation} degrees`}` : ''}`,
    font: [look.fontName, font.size, font.b && 'Bold', font.i && 'Italic', font.u && 'Underline', font.strike && 'Strikethrough'].filter(Boolean).join(' '),
    border: sides.length ? `${sides.map((k) => k[0].toUpperCase() + k.slice(1)).join(', ')} Borders` : 'No Borders',
    fill: fill && fill.pattern && fill.pattern !== 'none' ? 'Shaded' : 'No Shading',
    protection: [xf.protection?.locked !== false ? 'Locked' : '', xf.protection?.hidden ? 'Hidden' : ''].filter(Boolean).join(', ') || 'Unlocked',
  };
}

/** Format Cells changes put together: a later one's font, alignment, protection and border sides over the earlier
 *  one's, its fill and number format in place of them. */
export function mergeChanges(a, b) {
  if (!a) return b;
  if (!b) return a;
  const out = { ...a, ...b };
  for (const k of ['font', 'align', 'protection']) if (a[k] && b[k]) out[k] = { ...a[k], ...b[k] };
  if (a.borders && b.borders) out.borders = { ...a.borders, ...b.borders, outline: { ...a.borders.outline, ...b.borders.outline } };
  return out;
}

/**
 * The Style box. `name`: the style's (New: Excel's "Style 1" to start with); `rename`: whether the name can be changed
 * (New, or a style of the workbook's own); `includes`: { number, ... } true or false; `byExample`: New Cell Style's
 * heading "Style Includes (By Example)"; `look(change)`: formatbox.js's look of the style's format with `change` made
 * to it; `format(look)`: the Format Cells box for it (a change or null); `refusal(name)`: why a name can't be taken.
 */
export function styleBox(root, { name, rename = true, includes, byExample = false, look, format, refusal = () => null }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-style-box');
    let change = null;
    const nameInput = make('input', { type: 'text', value: name, className: 'oxe-dialog-field', readOnly: !rename, spellcheck: false });
    nameInput.dataset.first = '';
    const error = make('p', { className: 'oxe-dialog-note', hidden: true, role: 'alert' });
    const checks = PARTS.map(([key, label]) => ({ key, ...check(label, includes[key] !== false), line: make('span', { className: 'oxe-style-part' }) }));
    const describe = () => {
      const words = describeParts(look(change));
      for (const c of checks) {
        c.line.textContent = words[c.key];
        c.input.setAttribute('aria-describedby', (c.line.id ||= `${id}-${c.key}`));
      }
    };
    describe();
    const group = make('fieldset', { className: 'oxe-style-includes' }, make('legend', { textContent: byExample ? 'Style Includes (By Example)' : 'Style Includes' }),
      ...checks.map((c) => make('div', { className: 'oxe-style-row' }, c.el, c.line)));
    const formatBtn = button('Format...', async () => {
      const more = await format(look(change));
      if (more) { change = mergeChanges(change, more); describe(); }
      formatBtn.focus();
    });
    const ok = () => {
      const n = nameInput.value.trim();
      const why = rename && n !== name ? refusal(n) : !n ? refusal(n) : null;
      if (why) { error.textContent = why; error.hidden = false; nameInput.focus(); return; }
      done({ name: n, includes: Object.fromEntries(checks.map((c) => [c.key, c.input.checked])), change });
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT' && e.target.type === 'text') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Style' }), labelled('Style name:', nameInput), error, group,
      make('div', { className: 'oxe-dialog-buttons' }, formatBtn, button('OK', ok), button('Cancel', () => done(null))));
  });
}
