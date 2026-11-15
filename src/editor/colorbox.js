// The colour menus' More Colors... box: Excel's Colors box, its Custom tab (Red, Green and Blue from 0 to 255, and the
// hex value), with a colour field to pick from and the new colour shown beside the current one. Resolves to the colour
// as Excel writes a custom one, 'FFRRGGBB' (tools/excel-compare/excel-drawborders.ps1: Interior.Color, Font.Color and
// a border's Color all saved rgb="FF123456"), or null when cancelled.
import { modal, button } from './dialogs.js';

const two = (n) => n.toString(16).toUpperCase().padStart(2, '0');

export function colorBox(root, { current = null } = {}) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-colorbox');
    const make = (tag, props = {}) => Object.assign(document.createElement(tag), props);
    const start = /^(FF)?[0-9A-F]{6}$/i.test(current || '') ? current.slice(-6).toUpperCase() : '000000';
    const title = make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Colors' });
    const tab = make('p', { className: 'oxe-dialog-text', textContent: 'Custom' });
    const field = (label, key, props) => {
      const input = make('input', { id: `${id}-${key}`, className: 'oxe-dialog-input', ...props });
      const row = make('div', { className: 'oxe-colorbox-row' });
      row.append(make('label', { htmlFor: input.id, textContent: label }), input);
      return { input, row };
    };
    const picker = field('Color:', 'pick', { type: 'color', value: `#${start}` });
    const parts = ['Red', 'Green', 'Blue'].map((name, k) => field(`${name}:`, name.toLowerCase(), { type: 'number', min: 0, max: 255, step: 1, value: String(parseInt(start.slice(k * 2, k * 2 + 2), 16)) }));
    const hex = field('Hex:', 'hex', { type: 'text', value: `#${start}`, maxLength: 7, spellcheck: false });
    const swatch = (label) => {
      const box = make('span', { className: 'oxe-colorbox-swatch' });
      const wrap = make('div', { className: 'oxe-colorbox-sample' });
      wrap.append(box, make('span', { textContent: label }));
      return { box, wrap };
    };
    const fresh = swatch('New');
    const old = swatch('Current');
    old.box.style.setProperty('--oxe-swatch', `#${start}`);
    let value = start;
    // (Each field follows the others: the numbers clamp to 0 to 255, the hex value takes 6 digits with or without #.)
    const set = (rgb, from) => {
      value = rgb;
      if (from !== 'pick') picker.input.value = `#${rgb}`;
      if (from !== 'hex') hex.input.value = `#${rgb}`;
      if (from !== 'parts') parts.forEach((p, k) => { p.input.value = String(parseInt(rgb.slice(k * 2, k * 2 + 2), 16)); });
      fresh.box.style.setProperty('--oxe-swatch', `#${rgb}`);
      fresh.box.setAttribute('aria-label', `New colour: red ${parseInt(rgb.slice(0, 2), 16)}, green ${parseInt(rgb.slice(2, 4), 16)}, blue ${parseInt(rgb.slice(4), 16)}`);
    };
    picker.input.addEventListener('input', () => set(picker.input.value.slice(1).toUpperCase(), 'pick'));
    for (const p of parts) {
      p.input.addEventListener('input', () => {
        const n = parts.map((q) => Math.max(0, Math.min(255, Math.round(Number(q.input.value) || 0))));
        set(n.map(two).join(''), 'parts');
      });
      p.input.addEventListener('change', () => set(value));
    }
    hex.input.addEventListener('input', () => {
      const h = hex.input.value.trim().replace(/^#/, '');
      if (/^[0-9a-f]{6}$/i.test(h)) set(h.toUpperCase(), 'hex');
    });
    hex.input.addEventListener('change', () => set(value));
    set(start);
    const ok = () => done(`FF${value}`);
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); ok(); } });
    parts[0].input.dataset.first = '';
    const samples = make('div', { className: 'oxe-colorbox-samples' });
    samples.append(fresh.wrap, old.wrap);
    const row = make('div', { className: 'oxe-dialog-buttons' });
    row.append(button('OK', ok), button('Cancel', () => done(null)));
    dlg.append(title, tab, picker.row, ...parts.map((p) => p.row), hex.row, samples, row);
  });
}
