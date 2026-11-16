// The Text Import box: Excel's Text Import Wizard's three steps in one box. The file's origin (encoding) and the
// locale its numbers and dates are written in (with the decimal and thousands separators, as the wizard's Advanced box
// sets them); delimited (Tab, Semicolon, Comma, Space, Other, consecutive delimiters as one, the text qualifier) or
// fixed width (column breaks typed, or clicked in the preview, as the wizard's ruler takes them); the row to start
// at; a preview of the first rows with each column's data format (General, Text, Date: MDY, DMY or YMD, Do not
// import) chosen over it. Resolves to the settings (ooxml/textimport.js importCells and splitText take them), or null.
import { modal, button } from './dialogs.js';
import { make, select, check, labelled } from './formatbox-kit.js';
import { splitText, linesOf } from '../ooxml/textimport.js';
import { localeOf } from '../ooxml/locale.js';

const TAB = String.fromCharCode(9);
const ENCODINGS = [['utf-8', 'Unicode (UTF-8)'], ['utf-16le', 'Unicode (UTF-16)'], ['utf-16be', 'Unicode (UTF-16 big endian)'], ['windows-1252', 'Western European (Windows-1252)']];
export const LOCALES = [['en-US', 'English (United States)'], ['en-GB', 'English (United Kingdom)'], ['de-DE', 'German (Germany)'], ['fr-FR', 'French (France)'],
  ['es-ES', 'Spanish (Spain)'], ['it-IT', 'Italian (Italy)'], ['nl-NL', 'Dutch (Netherlands)'], ['pt-BR', 'Portuguese (Brazil)'], ['sv-SE', 'Swedish (Sweden)'],
  ['pl-PL', 'Polish (Poland)'], ['ru-RU', 'Russian (Russia)'], ['ja-JP', 'Japanese (Japan)'], ['zh-CN', 'Chinese (China)']];
const TYPES = [['general', 'General'], ['text', 'Text'], ['MDY', 'Date: MDY'], ['DMY', 'Date: DMY'], ['YMD', 'Date: YMD'], ['skip', 'Do not import column (skip)']];
const SHOWN_ROWS = 12;
const SHOWN_COLS = 30;

/** The box for a file's `bytes` (`fileName` its title), starting from `settings` (textimport.js guessSettings). */
export function textImportBox(root, { bytes, fileName = '', settings }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-textimport-box');
    const s = { ...settings, columns: [...(settings.columns || [])], breaks: [...(settings.breaks || [])] };
    if (!ENCODINGS.some(([v]) => v === s.encoding)) s.encoding = 'utf-8';
    const origin = select(ENCODINGS, s.encoding);
    const tags = LOCALES.some(([v]) => v === s.locale) ? LOCALES : [[s.locale, s.locale], ...LOCALES];
    const locale = select(tags, s.locale);
    const decimal = make('input', { type: 'text', className: 'oxe-fc-num', maxLength: 1, value: s.decimal });
    const group = make('input', { type: 'text', className: 'oxe-fc-num', maxLength: 1, value: s.group });
    const start = make('input', { type: 'number', className: 'oxe-fc-num', min: 1, value: String(s.startRow || 1) });
    const kindName = `${id}-kind`;
    const kinds = Object.fromEntries([['delimited', 'Delimited'], ['fixed', 'Fixed width']].map(([v, l]) => {
      const input = make('input', { type: 'radio', name: kindName, id: `${kindName}-${v}`, value: v, checked: s.kind === v });
      return [v, { input, el: make('div', { className: 'oxe-fc-check' }, input, make('label', { htmlFor: input.id, textContent: l })) }];
    }));
    const seps = { tab: check('Tab', s.seps.includes(TAB)), semicolon: check('Semicolon', s.seps.includes(';')), comma: check('Comma', s.seps.includes(',')), space: check('Space', s.seps.includes(' ')) };
    const otherChar = [...s.seps].find((ch) => ![TAB, ';', ',', ' '].includes(ch)) || '';
    const other = check('Other:', !!otherChar);
    const otherText = make('input', { type: 'text', className: 'oxe-fc-num', maxLength: 1, value: otherChar });
    otherText.setAttribute('aria-label', 'Other delimiter');
    const consecutive = check('Treat consecutive delimiters as one', s.consecutive);
    const quote = select([['"', '"'], ["'", "'"], ['', '{none}']], s.quote);
    const breaks = make('input', { type: 'text', className: 'oxe-dialog-field', spellcheck: false, placeholder: '10, 20', value: s.breaks.join(', ') });
    const raw = make('pre', { className: 'oxe-textimport-raw', tabIndex: 0, title: 'Click to add or remove a column break' });
    raw.setAttribute('aria-label', 'File preview with column breaks');
    const table = make('table', { className: 'oxe-textimport-preview' });
    table.setAttribute('aria-label', 'Data preview');
    const why = make('p', { className: 'oxe-dialog-text', role: 'alert' });

    const delimitedEl = make('fieldset', { className: 'oxe-textimport-delims' }, make('legend', { textContent: 'Delimiters' }), seps.tab.el, seps.semicolon.el, seps.comma.el, seps.space.el,
      make('div', { className: 'oxe-textimport-other' }, other.el, otherText), consecutive.el, labelled('Text qualifier:', quote));
    const fixedEl = make('fieldset', { className: 'oxe-textimport-fixed' }, make('legend', { textContent: 'Column breaks' }),
      labelled('Breaks (characters from the left):', breaks), raw);

    const read = () => {
      s.encoding = origin.value;
      s.kind = kinds.fixed.input.checked ? 'fixed' : 'delimited';
      s.seps = [seps.tab.input.checked ? TAB : '', seps.semicolon.input.checked ? ';' : '', seps.comma.input.checked ? ',' : '', seps.space.input.checked ? ' ' : '',
        other.input.checked ? otherText.value : ''].join('');
      s.consecutive = consecutive.input.checked;
      s.quote = quote.value;
      s.breaks = [...new Set(breaks.value.split(/[\s,;]+/).filter(Boolean).map(Number).filter((n) => Number.isInteger(n) && n > 0))].sort((a, b) => a - b);
      s.startRow = Math.max(1, Math.floor(Number(start.value) || 1));
      s.locale = locale.value;
      s.decimal = decimal.value || '.';
      s.group = group.value || ',';
    };
    const typeList = (k) => {
      const list = select(TYPES, s.columns[k] || 'general');
      list.setAttribute('aria-label', `Column ${k + 1} data format`);
      list.addEventListener('change', () => { s.columns[k] = list.value; draw(); });
      return list;
    };
    const draw = () => {
      read();
      delimitedEl.hidden = s.kind === 'fixed';
      fixedEl.hidden = s.kind !== 'fixed';
      const { text, rows } = splitText(bytes, s);
      if (s.kind === 'fixed') {
        // (The file's first lines with a bar where each column starts.)
        const at = [...s.breaks].reverse();
        raw.textContent = linesOf(text).slice(s.startRow - 1, s.startRow - 1 + SHOWN_ROWS).map((t) => at.reduce((line, n) => `${line.padEnd(n).slice(0, n)}|${line.slice(n)}`, t)).join('\n');
      }
      const shown = rows.slice(s.startRow - 1, s.startRow - 1 + SHOWN_ROWS);
      const width = Math.min(SHOWN_COLS, Math.max(1, ...shown.map((r) => r.length)));
      const head = make('tr', {}, ...Array.from({ length: width }, (_, k) => make('th', { scope: 'col' }, typeList(k))));
      const body = shown.map((r) => make('tr', {}, ...Array.from({ length: width }, (_, k) => make('td', {
        className: s.columns[k] === 'skip' ? 'oxe-textimport-skip' : '', textContent: r[k] ?? '',
      }))));
      table.replaceChildren(make('thead', {}, head), make('tbody', {}, ...body));
    };
    // (A click in the fixed-width preview adds a break where it is, or takes away the one there, as the wizard's ruler.)
    raw.addEventListener('click', (e) => {
      const probe = make('span', { textContent: '0' });
      raw.append(probe);
      const ch = probe.getBoundingClientRect().width || 8;
      probe.remove();
      const n = Math.round((e.clientX - raw.getBoundingClientRect().left - raw.clientLeft + raw.scrollLeft - parseFloat(getComputedStyle(raw).paddingLeft || '0')) / ch);
      if (n <= 0) return;
      // (The bars drawn: each one before moves the text after it one place right. A click on a bar takes it away.)
      const bars = s.breaks.map((b, k) => b + k);
      const hit = bars.findIndex((p) => p === n || p === n - 1);
      const next = hit >= 0 ? s.breaks.filter((_, k) => k !== hit) : [...s.breaks, n - bars.filter((p) => p < n).length];
      breaks.value = next.sort((a, b) => a - b).join(', ');
      draw();
    });
    // (Another locale gives its separators, as the wizard's Advanced box would be set for it.)
    locale.addEventListener('change', () => { const l = localeOf(locale.value); decimal.value = l.decimal; group.value = l.group; draw(); });
    for (const x of [decimal, group, start, otherText, breaks]) x.addEventListener('input', draw);
    for (const x of [origin, kinds.delimited.input, kinds.fixed.input, seps.tab.input, seps.semicolon.input, seps.comma.input, seps.space.input, other.input, consecutive.input, quote]) x.addEventListener('change', draw);
    const ok = () => {
      read();
      if (s.kind === 'fixed' && !s.breaks.length && breaks.value.trim()) { why.textContent = 'Type where the columns break, in characters from the left, like 10, 20.'; return; }
      done({ ...s, columns: [...s.columns] });
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT' && e.target.type !== 'checkbox' && e.target.type !== 'radio') { e.preventDefault(); ok(); } });
    const top = make('div', { className: 'oxe-textimport-top' }, labelled('File origin:', origin), labelled('Locale:', locale), labelled('Decimal separator:', decimal),
      labelled('Thousands separator:', group), labelled('Start import at row:', start));
    const kindEl = make('fieldset', { className: 'oxe-fc-radios' }, make('legend', { textContent: 'Original data type' }), kinds.delimited.el, kinds.fixed.el);
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: fileName ? `Text Import - ${fileName.replace(/^.*[\\/]/, '')}` : 'Text Import' }),
      top, kindEl, delimitedEl, fixedEl, make('div', { className: 'oxe-textimport-scroll' }, table), why,
      make('div', { className: 'oxe-dialog-buttons' }, button('Import', ok), button('Cancel', () => done(null))));
    origin.dataset.first = '';
    draw();
  });
}
