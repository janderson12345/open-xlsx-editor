// Insert > PivotTable's boxes: Excel 365's "PivotTable from table or range" (the source, and New Worksheet or a cell of
// one there is), then the fields, as Excel's field list places them: each field to Rows, Columns or Values, a value
// summarized by Sum, Count, Average, Max or Min (Sum for a field of numbers, Count otherwise, as Excel starts them).
// Rows and Columns take their fields in the list's order; so do Values.
import { modal, button, messageBox } from './dialogs.js';
import { PIVOT_FUNCTIONS, pivotSourceFields, layoutRefusal } from '../ooxml/pivotnew.js';
import { PIVOT_MESSAGES, pivotAt } from '../ooxml/pivot.js';
import { pivotLayout } from '../ooxml/pivotfields.js';
import { goToTarget } from '../ooxml/goto.js';
import { readTables } from '../ooxml/tables.js';
import { currentRegion } from '../ooxml/special.js';
import { sheetPrefix } from '../calc/tokens.js';
import { colName } from '../ooxml/refs.js';

/** What the Table/Range field starts with, as Excel fills it: the table the active cell is in, by name, or the
 *  current region round it (with its sheet, absolute), or nothing on an empty cell alone. */
export function pivotSourceGuess(wb, si, r, c) {
  const t = readTables(wb).find((x) => x.si === si && r >= x.r0 && r <= x.r1 && c >= x.c0 && c <= x.c1);
  if (t) return t.name;
  const a = currentRegion(wb.sheet(si), r, c);
  if (a.r0 === a.r1 && a.c0 === a.c1 && wb.sheet(si).cells.get(r)?.get(c)?.v == null) return '';
  return `${sheetPrefix(wb.sheets[si].name)}$${colName(a.c0)}$${a.r0 + 1}:$${colName(a.c1)}$${a.r1 + 1}`;
}

/** Insert > PivotTable, from the active cell (r, c) of sheet si: the two boxes, then `run('newPivot', ...)`
 *  (commands-pivot.js), Excel's question asked when it answers { ask }. */
export async function insertPivot(root, wb, si, { r, c }, run) {
  const answer = await pivotSourceBox(root, { range: pivotSourceGuess(wb, si, r, c) });
  if (!answer) return false;
  // (A table's name is the table with its headings, without its totals row; anything else is read as the Name Box
  // reads it.)
  const table = readTables(wb).find((x) => x.name.toLowerCase() === answer.range.toLowerCase());
  const to = table ? null : answer.range ? goToTarget(wb, si, answer.range) : { error: true };
  const src = table ? { si: table.si, r0: table.r0, c0: table.c0, r1: table.r1 - table.totals, c1: table.c1 }
    : to.error ? null : { si: to.si, r0: to.sel.r0, c0: to.sel.c0, r1: to.sel.r1, c1: to.sel.c1 };
  const fields = src ? pivotSourceFields(wb, src) : { error: PIVOT_MESSAGES.source };
  if (fields.error) { await messageBox(root, fields.error); return false; }
  let at = null;
  if (answer.where === 'existing') {
    const loc = answer.location ? goToTarget(wb, si, answer.location) : { error: true };
    if (loc.error) { await messageBox(root, PIVOT_MESSAGES.source); return false; }
    at = { si: loc.si, r: loc.sel.r0, c: loc.sel.c0 };
  }
  const sheet = wb.sheet(src.si);
  const numeric = fields.names.map((_, k) => {
    let any = false;
    for (let rr = src.r0 + 1; rr <= src.r1; rr++) {
      const cell = sheet.cells.get(rr)?.get(src.c0 + k);
      if (cell?.v == null || cell.v === '') continue;
      if (cell.t !== 'n') return false;
      any = true;
    }
    return any;
  });
  const layout = await pivotFieldsBox(root, fields.names, numeric, layoutRefusal);
  if (!layout) return false;
  const done = run('newPivot', src, layout, at);
  if (!done?.ask) return done;
  if (await messageBox(root, done.ask, ['OK', 'Cancel']) !== 'OK') return false;
  return run('newPivot', src, layout, at, { replace: true });
}

const make = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids); return e; };

/** The first box: resolves to { range, where: 'new' | 'existing', location }, or null for Cancel and Esc. */
export function pivotSourceBox(root, { range = '' } = {}) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-pivot-box');
    const field = (key, label, value) => {
      const input = make('input', { id: `${id}-${key}`, className: 'oxe-dialog-field', type: 'text', spellcheck: false, value });
      return [input, make('label', { className: 'oxe-dialog-text oxe-dialog-label', htmlFor: input.id, textContent: label })];
    };
    const [rangeInput, rangeLabel] = field('range', 'Table/Range:', range);
    rangeInput.dataset.first = '';
    const radio = (value, label, checked) => {
      const input = make('input', { type: 'radio', name: `${id}-where`, id: `${id}-${value}`, value, checked });
      return [input, make('div', { className: 'oxe-dialog-option' }, input, make('label', { htmlFor: input.id, textContent: label }))];
    };
    const [newInput, newLine] = radio('new', 'New Worksheet', true);
    const [, oldLine] = radio('existing', 'Existing Worksheet', false);
    const [locInput, locLabel] = field('loc', 'Location:', '');
    const sync = () => { locInput.disabled = newInput.checked; };
    const where = make('fieldset', { className: 'oxe-dialog-group' }, make('legend', { textContent: 'Choose where you want the PivotTable to be placed' }), newLine, oldLine);
    where.style.setProperty('--oxe-rows', '2');
    where.addEventListener('change', sync);
    sync();
    const ok = () => done({ range: rangeInput.value.trim(), where: newInput.checked ? 'new' : 'existing', location: locInput.value.trim() });
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); ok(); } });
    dlg.append(
      make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'PivotTable from table or range' }),
      make('p', { className: 'oxe-dialog-text', textContent: 'Select a table or range' }), rangeLabel, rangeInput, where, locLabel, locInput,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))),
    );
  });
}

/** The fields box for fields `names` (`numeric[i]`: field i holds only numbers): resolves to the layout ({ rows, cols,
 *  values: [{ field, fn }] }), or null. `check(layout)` says why a layout can't be made (shown in the box), or null.
 *  `start`: a pivot's layout now (Show Field List), its fields kept in their order, one newly put in an area last. */
export function pivotFieldsBox(root, names, numeric, check, start = null) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-pivot-fields');
    const select = (label, options, value) => {
      const s = make('select', { className: 'oxe-dialog-select' });
      s.setAttribute('aria-label', label);
      s.append(...options.map(([v, text]) => make('option', { value: v, textContent: text, selected: v === value })));
      return s;
    };
    const rows = names.map((name, i) => {
      const value = start?.values.find((v) => v.field === i);
      const now = start?.rows.includes(i) ? 'rows' : start?.cols.includes(i) ? 'cols' : value ? 'values' : '';
      const area = select(`${name}: area`, [['', '(not used)'], ['rows', 'Rows'], ['cols', 'Columns'], ['values', 'Values']], now);
      const fn = select(`${name}: summarize by`, Object.entries(PIVOT_FUNCTIONS), value ? value.fn : numeric[i] ? 'sum' : 'count');
      const sync = () => { fn.disabled = area.value !== 'values'; };
      area.addEventListener('change', sync);
      sync();
      return { area, fn, row: make('tr', {}, make('th', { scope: 'row', textContent: name }), make('td', {}, area), make('td', {}, fn)) };
    });
    if (rows[0]) rows[0].area.dataset.first = '';
    const table = make('table', { className: 'oxe-pivot-table' },
      make('thead', {}, make('tr', {}, make('th', { scope: 'col', textContent: 'Field' }), make('th', { scope: 'col', textContent: 'Area' }), make('th', { scope: 'col', textContent: 'Summarize by' }))),
      make('tbody', {}, ...rows.map((x) => x.row)));
    const note = make('p', { className: 'oxe-dialog-text oxe-pivot-why', id: `${id}-why` });
    note.setAttribute('role', 'alert');
    /** The fields in an area: those there already in their order, then the ones newly put there. */
    const ordered = (where, was = []) => {
      const now = rows.flatMap((x, i) => (x.area.value === where ? [i] : []));
      return [...was.filter((i) => now.includes(i)), ...now.filter((i) => !was.includes(i))];
    };
    const layout = () => ({
      rows: ordered('rows', start?.rows),
      cols: ordered('cols', start?.cols),
      values: ordered('values', start?.values.map((v) => v.field)).map((i) => ({ field: i, fn: rows[i].fn.value })),
    });
    const ok = () => {
      const l = layout();
      const why = check(l);
      if (why) { note.textContent = why; return; }
      done(l);
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON' && e.target.tagName !== 'SELECT') { e.preventDefault(); ok(); } });
    dlg.append(
      make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'PivotTable Fields' }),
      make('p', { className: 'oxe-dialog-text', textContent: 'Choose fields to add to report:' }), table, note,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))),
    );
  });
}

/** Show Field List, on a PivotTable's cell (r, c) of sheet si: the fields box with its layout now, then
 *  `run('movePivot', layout)` (commands-pivot.js). */
export async function showFieldList(root, wb, si, { r, c }, run) {
  const p = pivotAt(wb, si, r, c);
  const now = p && pivotLayout(wb, p);
  if (!now) { await messageBox(root, p ? "This PivotTable's values can't be changed here yet" : 'The active cell is not in a PivotTable'); return false; }
  const layout = await pivotFieldsBox(root, now.names, now.numeric, layoutRefusal, now);
  return layout ? run('movePivot', layout) : false;
}
