// Function Arguments: the box Excel opens after a function is picked in Insert Function, from the Function Library's
// lists, or with Ctrl+A after a function's name in a cell being edited. One labelled box per argument (the required
// ones' names in bold, a repeated one adding a box once the last is filled: calc/fnargs.js), each with its value worked
// out beside it at the cell; the function's description and the description of the argument with focus
// (calc/describe.js); "Formula result =" for the whole call; OK and Cancel. Resolves to the call written (NAME(...),
// `refused` when Excel would refuse it) or null.
//
// As in Excel the box doesn't hold the page: the sheet can be clicked, dragged over or pointed at with the arrow keys
// while a box has focus, its reference going in (editor/pointing.js), on another sheet too (its tab clicked). Each box
// has Excel's collapse button, which shrinks the box to that argument alone so the sheet shows, and back. An empty box
// shows the kind of value it takes (= number, = text, = logical, = reference, = array, = any: calc/argkinds.js); text
// typed in a text or any box is put in quotes when focus leaves it, if it isn't a formula (calc/fnargs.js needsQuotes).
import { modal, button } from './dialogs.js';
import { functionInfo } from '../calc/catalog.js';
import { argBoxes, buildCall, needsQuotes, OWN_NAMES } from '../calc/fnargs.js';
import { argKind } from '../calc/argkinds.js';
import { describeFunction, describeArg } from '../calc/describe.js';
import { normalizeFormula } from '../calc/fileform.js';
import { valueText } from '../calc/stepper.js';
import { XlError, numberText } from '../calc/values.js';

const make = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids); return e; };
/** "number1" as Excel labels it: "Number1". */
const label = (name) => name.charAt(0).toUpperCase() + name.slice(1);

/** The value an argument typed (`text`) has at cell (r, c) of sheet si, as the box shows it beside the argument:
 *  6, "abc", TRUE, {1;2;3}; '' when it is empty or can't be worked out. */
export function argumentValue(calc, si, r, c, text) {
  if (!String(text).trim()) return '';
  let rows;
  try { rows = calc.formulaAt(si, r, c, normalizeFormula(text), true); } catch { return ''; }
  if (!rows || !rows.length) return '';
  if (rows.length === 1 && rows[0].length === 1) return valueText(rows[0][0]);
  if (rows.length * rows[0].length > 100) return '{...}';
  return `{${rows.map((row) => row.map((x) => valueText(x)).join(',')).join(';')}}`;
}

/** The whole call's value, as "Formula result =" shows it: text as it is, a number as Excel writes it. */
export function resultText(calc, si, r, c, call) {
  let v;
  try { v = calc.formulaAt(si, r, c, normalizeFormula(call)); } catch { return ''; }
  if (v == null) return '';
  if (v instanceof XlError) return v.code;
  if (typeof v === 'number') return numberText(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

/** The rows of boxes for `vals` (changed in place as they are typed in), in `box` (the box's parts and callbacks). */
function drawRows(box) {
  const { info, vals, rows } = box;
  // (Only the rows not there yet: the one typed in, or pointed into, stays the same element.)
  const boxes = argBoxes(info, vals);
  rows.append(...boxes.slice(rows.children.length).map((arg, j) => argRow(box, arg, rows.children.length + j)));
}

/** One argument's row: its name, its box, Collapse Dialog and its value. */
function argRow(box, arg, k) {
  const { info, vals, id, calc, at, help } = box;
  const kind = argKind(info.name, arg.name);
  const input = make('input', { id: `${id}-arg${k}`, className: 'oxe-dialog-field oxe-fnargs-field', type: 'text', spellcheck: false, value: vals[k] || '' });
  const value = make('output', { id: `${id}-val${k}`, className: 'oxe-fnargs-value', htmlFor: input.id });
  // (Its value; empty, the kind of value it takes, as Excel's box shows it.)
  const show = () => {
    const v = argumentValue(calc, at.si, at.r, at.c, input.value);
    value.textContent = `= ${v || (input.value.trim() ? '' : kind)}`;
    value.classList.toggle('oxe-fnargs-hint', !input.value.trim());
  };
  show();
  input.setAttribute('aria-describedby', `${value.id} ${help.id}`);
  if (k === 0) input.dataset.first = '';
  input.addEventListener('focus', () => { box.focused = k; box.current(input); });
  input.addEventListener('keydown', (e) => { if (box.pointing?.key(e)) e.preventDefault(); });
  input.addEventListener('input', () => {
    vals[k] = input.value;
    show();
    box.changed();
    drawRows(box);
  });
  // (Text that isn't a formula, in a text or any box: quoted as focus leaves it.)
  input.addEventListener('change', () => {
    if ((kind !== 'text' && kind !== 'any') || box.ownNames || !needsQuotes(input.value, box.isName)) return;
    input.value = `"${input.value.trim().replace(/"/g, '""')}"`;
    input.dispatchEvent(new Event('input'));
  });
  const name = make('label', { className: `oxe-dialog-text oxe-fnargs-name${arg.optional ? '' : ' oxe-fnargs-required'}`, htmlFor: input.id, textContent: label(arg.name) });
  // (Out of the Tab order: Tab goes from argument to argument, as in Excel's box.)
  const fold = make('button', { type: 'button', className: 'oxe-fnargs-fold', title: 'Collapse Dialog', textContent: '⇱', tabIndex: -1 });
  fold.setAttribute('aria-label', 'Collapse Dialog');
  // (The button doesn't take focus: the argument keeps it, to be pointed at.)
  fold.addEventListener('mousedown', (e) => e.preventDefault());
  fold.addEventListener('click', () => { input.focus(); box.fold(); });
  return make('div', { className: 'oxe-fnargs-row' }, name, input, fold, value);
}

/** The box for function `name` at cell `at` ({ si, r, c }) of `calc`, its boxes started with `values`. `pointer(input)`
 *  points from that argument (returns the pointing, for its keys); `isName(name)`: a name defined in the workbook. */
export function functionArgumentsBox(root, { name, values = [], calc, at, pointer = () => null, isName = () => false }) {
  const info = functionInfo(name);
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-fnargs-box');
    const rows = make('div', { className: 'oxe-fnargs-rows' });
    const help = make('p', { id: `${id}-help`, className: 'oxe-dialog-text oxe-fnargs-help' });
    help.setAttribute('aria-live', 'polite');
    const result = make('p', { id: `${id}-result`, className: 'oxe-dialog-text oxe-fnargs-result' });
    const box = { info, vals: values.map((v) => String(v).trim()), rows, id, calc, at, help, focused: 0, pointing: null, isName };
    box.ownNames = OWN_NAMES.has(info.name);
    const call = () => buildCall(info.name, box.vals, { isName });
    const showHelp = () => {
      const arg = argBoxes(info, box.vals)[box.focused];
      help.replaceChildren(...(arg ? [make('b', { textContent: `${label(arg.name)}: ` }), describeArg(info.name, arg.name) || ''] : []));
    };
    box.changed = () => { result.textContent = `Formula result = ${resultText(calc, at.si, at.r, at.c, call().text)}`; };
    box.current = (input) => {
      for (const row of rows.querySelectorAll('.oxe-fnargs-current')) row.classList.remove('oxe-fnargs-current');
      input.closest('.oxe-fnargs-row').classList.add('oxe-fnargs-current');
      box.pointing = pointer(input);
      showHelp();
    };
    // (Collapse Dialog: the box shrunk to the argument with focus, its button now Expand Dialog; and back.)
    box.fold = () => {
      const folded = dlg.classList.toggle('oxe-fnargs-collapsed');
      for (const b of rows.querySelectorAll('.oxe-fnargs-fold')) {
        const t = folded ? 'Expand Dialog' : 'Collapse Dialog';
        b.title = t;
        b.setAttribute('aria-label', t);
      }
    };
    drawRows(box);
    showHelp();
    box.changed();
    const ok = () => {
      // (Quotes added to the box with focus first, as leaving it would.)
      document.activeElement?.dispatchEvent?.(new Event('change'));
      done(call());
    };
    dlg.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter' || e.target.tagName === 'BUTTON') return;
      e.preventDefault();
      ok();
    });
    const group = make('fieldset', { className: 'oxe-fnargs-group' }, make('legend', { className: 'oxe-dialog-text', textContent: info.name }), rows);
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Function Arguments' }), group,
      make('p', { id: `${id}-about`, className: 'oxe-dialog-text oxe-fnargs-about', textContent: describeFunction(info.name) }), help, result,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
    dlg.setAttribute('aria-describedby', `${id}-about`);
  }, { modeless: true });
}
