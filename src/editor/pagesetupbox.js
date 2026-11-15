// Excel's Page Setup box at its Sheet tab, as Page Layout > Print Titles opens it: Print area, Rows to repeat at top,
// Columns to repeat at left, each typed as Excel takes it (ooxml/pagesetupedit.js readTitles: "1:2", "$1:$2" or a
// range's rows; "A:B" or a range's columns). The messages for what it won't take are the editor's own (Excel's box
// isn't reachable over COM).
import { modal, button } from './dialogs.js';
import { readTitles } from '../ooxml/pagesetupedit.js';
import { parseRange } from '../ooxml/refs.js';

const make = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids); return e; };

/** The print area typed ("$B$2:$D$5,$F$7:$G$8", the sheet's name before each allowed): its ranges, [] for none, or
 *  null when one isn't a range. */
export function readArea(text) {
  const parts = String(text || '').split(',').map((p) => p.trim().replace(/^(?:'(?:[^']|'')+'|[^!]+)!/, '').replace(/\$/g, '').toUpperCase()).filter(Boolean);
  const out = parts.map((p) => (/^[A-Z]{1,3}\d+(:[A-Z]{1,3}\d+)?$/.test(p) ? parseRange(p) : null));
  return out.some((x) => !x) ? null : out.map(({ r0, c0, r1, c1 }) => ({ r0, c0, r1, c1 }));
}

/** The box, its fields started at `now` ({ area, rows, cols }, as text): resolves to { area: ranges, rows, cols:
 *  [first, last] or null }, or null for Cancel and Esc. */
export function pageSetupBox(root, now) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-page-setup');
    const field = (key, label, value) => {
      const input = make('input', { id: `${id}-${key}`, className: 'oxe-dialog-field', value, spellcheck: false });
      return { input, line: make('div', { className: 'oxe-dialog-row' }, make('label', { className: 'oxe-dialog-label', htmlFor: input.id, textContent: label }), input) };
    };
    const area = field('area', 'Print area:', now.area || '');
    const rows = field('rows', 'Rows to repeat at top:', now.rows || '');
    const cols = field('cols', 'Columns to repeat at left:', now.cols || '');
    area.input.dataset.first = '';
    const note = make('p', { className: 'oxe-dialog-note' });
    note.setAttribute('role', 'alert');
    const refuse = (f, text) => { note.textContent = text; f.input.focus(); f.input.select(); };
    const ok = () => {
      const a = readArea(area.input.value);
      if (!a) return refuse(area, 'The print area takes ranges such as $B$2:$D$5, separated by commas.');
      const r = readTitles(rows.input.value, 'rows');
      if (r === undefined) return refuse(rows, 'Rows to repeat take whole rows such as $1:$2.');
      const c = readTitles(cols.input.value, 'cols');
      if (c === undefined) return refuse(cols, 'Columns to repeat take whole columns such as $A:$B.');
      return done({ area: a, rows: r, cols: c });
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Page Setup' }),
      make('p', { className: 'oxe-dialog-text', textContent: 'Sheet' }), area.line,
      make('fieldset', { className: 'oxe-dialog-group' }, make('legend', { textContent: 'Print titles' }), rows.line, cols.line), note,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
  });
}
