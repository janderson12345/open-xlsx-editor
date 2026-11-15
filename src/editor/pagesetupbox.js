// Excel's Page Setup box, its Margins and Sheet tabs: Top, Bottom, Left, Right, Header and Footer margins (inches) and
// Center on page, Horizontally and Vertically; Print area, Rows to repeat at top, Columns to repeat at left, and Page
// order. Opened by Page Layout > Margins > Custom Margins... at Margins, by Print Titles at Sheet. Rows and columns are
// typed as Excel takes them (ooxml/pagesetupedit.js readTitles: "1:2", "$1:$2" or a range's rows; "A:B" or a range's
// columns). The messages for what it won't take are the editor's own (Excel's box isn't reachable over COM).
import { modal, button } from './dialogs.js';
import { readTitles } from '../ooxml/pagesetupedit.js';
import { parseRange } from '../ooxml/refs.js';

const make = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids); return e; };
const SIDES = [['top', 'Top:'], ['bottom', 'Bottom:'], ['left', 'Left:'], ['right', 'Right:'], ['header', 'Header:'], ['footer', 'Footer:']];

/** The print area typed ("$B$2:$D$5,$F$7:$G$8", the sheet's name before each allowed): its ranges, [] for none, or
 *  null when one isn't a range. */
export function readArea(text) {
  const parts = String(text || '').split(',').map((p) => p.trim().replace(/^(?:'(?:[^']|'')+'|[^!]+)!/, '').replace(/\$/g, '').toUpperCase()).filter(Boolean);
  const out = parts.map((p) => (/^[A-Z]{1,3}\d+(:[A-Z]{1,3}\d+)?$/.test(p) ? parseRange(p) : null));
  return out.some((x) => !x) ? null : out.map(({ r0, c0, r1, c1 }) => ({ r0, c0, r1, c1 }));
}

/** The box at `tab` ('margins' or 'sheet'), its fields started at `now` ({ margins, center: { h, v }, area, rows, cols
 *  as text, pageOrder }): resolves to { margins, center, area: ranges, rows, cols: [first, last] or null, pageOrder },
 *  or null for Cancel and Esc. */
export function pageSetupBox(root, now, { tab = 'sheet' } = {}) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-page-setup');
    const field = (key, label, value) => {
      const input = make('input', { id: `${id}-${key}`, className: 'oxe-dialog-field', value: String(value ?? ''), spellcheck: false });
      return { input, line: make('div', { className: 'oxe-dialog-row' }, make('label', { className: 'oxe-dialog-label', htmlFor: input.id, textContent: label }), input) };
    };
    const check = (key, label, on) => {
      const input = make('input', { type: 'checkbox', id: `${id}-${key}`, checked: !!on });
      return { input, line: make('div', { className: 'oxe-dialog-option' }, input, make('label', { htmlFor: input.id, textContent: label })) };
    };
    // Margins tab.
    const sides = Object.fromEntries(SIDES.map(([k, label]) => [k, field(k, label, now.margins?.[k] ?? '')]));
    const centerH = check('ch', 'Horizontally', now.center?.h);
    const centerV = check('cv', 'Vertically', now.center?.v);
    const marginsPanel = make('div', { id: `${id}-p-margins`, role: 'tabpanel' }, ...SIDES.map(([k]) => sides[k].line),
      make('fieldset', { className: 'oxe-dialog-group' }, make('legend', { textContent: 'Center on page' }), centerH.line, centerV.line));
    // Sheet tab.
    const area = field('area', 'Print area:', now.area || '');
    const rows = field('rows', 'Rows to repeat at top:', now.rows || '');
    const cols = field('cols', 'Columns to repeat at left:', now.cols || '');
    const order = [['downThenOver', 'Down, then over'], ['overThenDown', 'Over, then down']].map(([v, label], k) => {
      const input = make('input', { type: 'radio', name: `${id}-order`, id: `${id}-order${k}`, value: v, checked: (now.pageOrder || 'downThenOver') === v });
      return { input, line: make('div', { className: 'oxe-dialog-option' }, input, make('label', { htmlFor: input.id, textContent: label })) };
    });
    const sheetPanel = make('div', { id: `${id}-p-sheet`, role: 'tabpanel' }, area.line,
      make('fieldset', { className: 'oxe-dialog-group' }, make('legend', { textContent: 'Print titles' }), rows.line, cols.line),
      make('fieldset', { className: 'oxe-dialog-group' }, make('legend', { textContent: 'Page order' }), ...order.map((x) => x.line)));
    // (The tabs: Excel's two that are here, arrow keys between them.)
    const panels = { margins: marginsPanel, sheet: sheetPanel };
    const tabs = [['margins', 'Margins'], ['sheet', 'Sheet']].map(([key, label]) => make('button', { type: 'button', role: 'tab', id: `${id}-t-${key}`, textContent: label, className: 'oxe-fc-tab' }));
    const show = (key, focus = false) => {
      tabs.forEach((t, k) => { const on = ['margins', 'sheet'][k] === key; t.setAttribute('aria-selected', String(on)); t.tabIndex = on ? 0 : -1; });
      for (const [k, p] of Object.entries(panels)) { p.hidden = k !== key; p.setAttribute('aria-labelledby', `${id}-t-${k}`); }
      if (focus) tabs[key === 'margins' ? 0 : 1].focus();
    };
    tabs.forEach((t, k) => {
      t.setAttribute('aria-controls', `${id}-p-${['margins', 'sheet'][k]}`);
      t.addEventListener('click', () => show(['margins', 'sheet'][k]));
      t.addEventListener('keydown', (e) => { if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') { e.preventDefault(); show(['margins', 'sheet'][1 - k], true); } });
    });
    show(tab);
    (tab === 'margins' ? sides.top : area).input.dataset.first = '';
    const note = make('p', { className: 'oxe-dialog-note' });
    note.setAttribute('role', 'alert');
    const refuse = (f, key, text) => { show(key); note.textContent = text; f.input.focus(); f.input.select(); };
    const ok = () => {
      const margins = {};
      for (const [k] of SIDES) {
        const v = Number(sides[k].input.value.trim());
        if (sides[k].input.value.trim() === '' || !(v >= 0) || v > 99) return refuse(sides[k], 'margins', 'Margins take a number of inches from 0 up.');
        margins[k] = v;
      }
      const a = readArea(area.input.value);
      if (!a) return refuse(area, 'sheet', 'The print area takes ranges such as $B$2:$D$5, separated by commas.');
      const r = readTitles(rows.input.value, 'rows');
      if (r === undefined) return refuse(rows, 'sheet', 'Rows to repeat take whole rows such as $1:$2.');
      const c = readTitles(cols.input.value, 'cols');
      if (c === undefined) return refuse(cols, 'sheet', 'Columns to repeat take whole columns such as $A:$B.');
      return done({ margins, center: { h: centerH.input.checked, v: centerV.input.checked }, area: a, rows: r, cols: c, pageOrder: order.find((x) => x.input.checked).input.value });
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Page Setup' }),
      make('div', { role: 'tablist', className: 'oxe-fc-tabs', 'aria-label': 'Page Setup' }, ...tabs), marginsPanel, sheetPanel, note,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
  });
}
