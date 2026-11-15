// A conditional format's own format (New Formatting Rule's Format..., Excel's Format Cells for a rule): the number
// format, bold, italic, underline, strikethrough, the font colour, the fill and a border, as cfedit.js formatDxf writes
// them. Excel's box has the Number, Font, Border and Fill tabs; this one has the same choices on one page.
import { modal, button } from './dialogs.js';
import { make, check, select, labelled } from './formatbox-kit.js';

const argb = (hex) => `FF${hex.replace('#', '').toUpperCase()}`;

/** Resolves to the format ({ numFmt, bold, italic, underline, strike, color, fill, border }) or null. */
export function ruleFormatBox(root) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-cf-format');
    const num = make('input', { type: 'text', className: 'oxe-dialog-field', placeholder: 'General', spellcheck: false });
    const bold = check('Bold', false);
    const italic = check('Italic', false);
    const underline = check('Underline', false);
    const strike = check('Strikethrough', false);
    const colour = (label) => {
      const on = check(label, false);
      const pick = make('input', { type: 'color', value: '#c00000' });
      pick.setAttribute('aria-label', `${label} colour`);
      pick.addEventListener('input', () => { on.input.checked = true; });
      return { on, pick, el: make('div', { className: 'oxe-fc-check' }, on.el, pick) };
    };
    const ink = colour('Font colour');
    const fill = colour('Fill');
    fill.pick.value = '#ffc7ce';
    const border = select([['none', 'None'], ['outline', 'Outline'], ['bottom', 'Bottom'], ['top', 'Top']], 'none');
    const line = select([['thin', 'Thin'], ['dashed', 'Dashed'], ['dotted', 'Dotted'], ['hair', 'Hair']], 'thin');
    const edge = colour('Border colour');
    edge.pick.value = '#000000';
    const ok = () => {
      const f = {};
      if (num.value.trim()) f.numFmt = num.value.trim();
      for (const [k, c] of [['bold', bold], ['italic', italic], ['underline', underline], ['strike', strike]]) if (c.input.checked) f[k] = true;
      if (ink.on.input.checked) f.color = argb(ink.pick.value);
      if (fill.on.input.checked) f.fill = argb(fill.pick.value);
      if (border.value !== 'none') {
        const side = { style: line.value, ...(edge.on.input.checked ? { color: argb(edge.pick.value) } : {}) };
        const names = border.value === 'outline' ? ['left', 'right', 'top', 'bottom'] : [border.value];
        f.border = Object.fromEntries(names.map((k) => [k, side]));
      }
      done(f);
    };
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Format Cells' }),
      labelled('Number format:', num), bold.el, italic.el, underline.el, strike.el, ink.el, fill.el,
      labelled('Border:', border), labelled('Line:', line), edge.el,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
    num.dataset.first = '';
  });
}
