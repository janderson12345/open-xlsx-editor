// Function Arguments: the box Excel opens after a function is picked in Insert Function, from the Function Library's
// lists, or with Ctrl+A after a function's name in a cell being edited. One labelled box per argument (the required
// ones' names in bold, a repeated one adding a box once the last is filled: calc/fnargs.js), each with its value worked
// out beside it at the cell; the function's description and the description of the argument with focus
// (calc/describe.js); "Formula result =" for the whole call; OK and Cancel. Resolves to the call written (NAME(...),
// `refused` when Excel would refuse it) or null.
import { modal, button } from './dialogs.js';
import { functionInfo } from '../calc/catalog.js';
import { argBoxes, buildCall } from '../calc/fnargs.js';
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

/** The box for function `name` at cell `at` ({ si, r, c }) of `calc`, its boxes started with `values`. */
export function functionArgumentsBox(root, { name, values = [], calc, at }) {
  const info = functionInfo(name);
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-fnargs-box');
    const vals = values.map((v) => String(v).trim());
    const rows = make('div', { className: 'oxe-fnargs-rows' });
    const help = make('p', { id: `${id}-help`, className: 'oxe-dialog-text oxe-fnargs-help' });
    help.setAttribute('aria-live', 'polite');
    const result = make('p', { id: `${id}-result`, className: 'oxe-dialog-text oxe-fnargs-result' });
    let focused = 0;
    const call = () => buildCall(info.name, vals);
    const showHelp = () => {
      const box = argBoxes(info, vals)[focused];
      help.replaceChildren(...(box ? [make('b', { textContent: `${label(box.name)}: ` }), describeArg(info.name, box.name) || ''] : []));
    };
    const showResult = () => {
      const shown = resultText(calc, at.si, at.r, at.c, call().text);
      result.textContent = `Formula result = ${shown}`;
    };
    /** The rows drawn again when the number of boxes changes (a repeated argument's next box). */
    const draw = () => {
      const boxes = argBoxes(info, vals);
      rows.replaceChildren(...boxes.map((box, k) => {
        const input = make('input', { id: `${id}-arg${k}`, className: 'oxe-dialog-field oxe-fnargs-field', type: 'text', spellcheck: false, value: vals[k] || '' });
        const value = make('output', { id: `${id}-val${k}`, className: 'oxe-fnargs-value', htmlFor: input.id });
        const show = () => { const v = argumentValue(calc, at.si, at.r, at.c, input.value); value.textContent = v ? `= ${v}` : ''; };
        show();
        input.setAttribute('aria-describedby', `${value.id} ${help.id}`);
        if (k === 0) input.dataset.first = '';
        input.addEventListener('focus', () => { focused = k; showHelp(); });
        input.addEventListener('input', () => {
          vals[k] = input.value;
          show();
          showResult();
          if (argBoxes(info, vals).length !== boxes.length) { const caret = input.selectionStart; draw(); const again = rows.querySelector(`#${id}-arg${k}`); again.focus(); again.setSelectionRange(caret, caret); }
        });
        const name = make('label', { className: `oxe-dialog-text oxe-fnargs-name${box.optional ? '' : ' oxe-fnargs-required'}`, htmlFor: input.id, textContent: label(box.name) });
        return make('div', { className: 'oxe-fnargs-row' }, name, input, value);
      }));
    };
    draw();
    showHelp();
    showResult();
    const ok = () => done(call());
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
  });
}
