// Formulas > Error Checking: Excel's box going through the sheet's cells with an error indicator (ooxml/errorcheck.js)
// a row at a time from the active cell, each selected in turn: "Error in cell G3", what it holds, the error's name and
// what it means, its fix (errormenu.js errorChoices), Ignore Error, Edit in Formula Bar, Previous and Next. Past the
// last it starts again at the top; back where it started, "The error check is complete for the entire sheet."
import { modal, button, messageBox } from './dialogs.js';
import { errorFlags, ERROR_OPTIONS } from '../ooxml/errorcheck.js';
import { refName } from '../ooxml/refs.js';
import { editText } from '../ooxml/input.js';
import { errorChoices } from './errormenu.js';

const DONE = 'The error check is complete for the entire sheet.';
const make = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids); return e; };

/** The cells with an indicator in reading order (rows, then columns), each [r, c, rule]. */
function inOrder(env) {
  const flags = errorFlags(env.state.workbook, env.state.sheetIndex);
  return [...flags].map(([k, rules]) => [...k.split(',').map(Number), rules[0]]).sort((a, b) => a[0] - b[0] || a[1] - b[1]);
}
/** The first in `list` at or after (dir 1) or before (dir -1) cell (r, c), going round the sheet's end. */
function nextFrom(list, r, c, dir, inclusive = false) {
  if (!list.length) return null;
  const cmp = (e) => e[0] - r || e[1] - c;
  if (dir > 0) return list.find((e) => cmp(e) > 0 || (inclusive && cmp(e) === 0)) || list[0];
  return [...list].reverse().find((e) => cmp(e) < 0) || list[list.length - 1];
}

/** The Error Checking box over the sheet shown. Resolves when it is closed. */
export async function errorCheckingBox(env) {
  const { root, state, editing } = env;
  const s = editing().selection;
  const start = nextFrom(inOrder(env), s.ar, s.ac, 1, true);
  if (!start) return messageBox(root, DONE);
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-errcheck-box');
    const title = make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Error Checking' });
    const where = make('p', { className: 'oxe-dialog-text oxe-errcheck-where' });
    const what = make('p', { className: 'oxe-dialog-text oxe-errcheck-formula' });
    const name = make('p', { className: 'oxe-dialog-text oxe-errcheck-name' });
    const text = make('p', { className: 'oxe-dialog-text' });
    text.setAttribute('aria-live', 'polite');
    const side = make('div', { className: 'oxe-errcheck-actions' });
    const nav = make('div', { className: 'oxe-dialog-buttons' });
    let at = start;
    // (Whether the check has gone round past the sheet's end: back at or past where it started, it is done.)
    let wrapped = false;
    const before = (a, b) => a[0] - b[0] || a[1] - b[1];
    const finish = () => { done(null); return messageBox(root, DONE); };
    const show = () => {
      const [r, c, rule] = at;
      editing().run('select', r, c);
      const cell = state.workbook.sheet(state.sheetIndex).cells.get(r)?.get(c);
      const ch = errorChoices(env, rule, [{ r, c }], { r, c });
      where.textContent = `Error in cell ${refName(r, c)}`;
      what.textContent = editText(cell);
      name.textContent = ch.heading;
      text.textContent = ch.text;
      // (A fix or Ignore Error, then on to the next, as Excel goes on; Edit in Formula Bar closes the box.)
      const act = (choice, close = false) => button(choice.label, () => { choice.run(); if (close) done(null); else go(1); });
      side.replaceChildren(...ch.fixes.map((f) => act(f)), act(ch.ignore), act(ch.edit, true),
        // (Options...: Excel's Error Checking Options; the box closes, as Excel's waits on its options.)
        button('Options...', () => { done(null); env.act?.('errorOptions'); }));
      for (const b of side.children) b.disabled = state.readOnly;
      side.firstChild?.setAttribute('data-first', '');
    };
    /** On to the next (dir 1) or previous (-1) cell with an indicator. */
    const go = (dir) => {
      const next = nextFrom(inOrder(env), at[0], at[1], dir);
      if (dir > 0 && next && before(next, at) <= 0) wrapped = true;
      if (!next || (dir > 0 && wrapped && before(next, start) >= 0)) { finish(); return; }
      at = next;
      show();
    };
    nav.append(button('Previous', () => go(-1)), button('Next', () => go(1)), button('Close', () => done(null)));
    dlg.append(title, where, what, name, text, side, nav);
    show();
  });
}

/**
 * Excel's Error Checking Options (File > Options > Formulas): Enable background error checking, and each rule's tick
 * box. `now`: { background, rules: [on] }. Resolves to the same shape, or null. Excel keeps them as its own settings,
 * not in the workbook; so does the editor (index.js, in the browser).
 */
export function errorOptionsBox(root, now) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-erropts-box');
    const tick = (label, on) => {
      const input = make('input', { type: 'checkbox', checked: on, id: `${id}-${label.replace(/\W/g, '')}` });
      return { input, el: make('div', { className: 'oxe-dialog-option' }, input, make('label', { htmlFor: input.id, textContent: label })) };
    };
    const background = tick('Enable background error checking', now.background);
    const rules = ERROR_OPTIONS.map(([key, label]) => ({ key, ...tick(label, now.rules.includes(key)) }));
    const ok = () => done({ background: background.input.checked, rules: rules.filter((x) => x.input.checked).map((x) => x.key) });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Error Checking Options' }), background.el,
      make('fieldset', {}, make('legend', { textContent: 'Error checking rules' }), ...rules.map((x) => x.el)),
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
    background.input.dataset.first = '';
  });
}

