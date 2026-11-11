// Excel's Go To Special box (Home > Find & Select > Go To Special..., or Special... in the Go To box): its choices in
// two columns, as Excel lays them out, each with the options it takes under it. The types under Formulas are those of
// Constants too, Direct only and All levels under Dependents those of Precedents too, and All and Same under Data
// validation those of Conditional formats too; each is turned off unless its choice is. The last choice is kept for
// the next time, as Excel keeps it.
import { modal, button } from './dialogs.js';
import { TYPES } from '../ooxml/special.js';

const LEFT = [['notes', 'Notes'], ['constants', 'Constants'], ['formulas', 'Formulas'], ['blanks', 'Blanks'], ['region', 'Current region'], ['array', 'Current array']];
const RIGHT = [['rowDiff', 'Row differences'], ['colDiff', 'Column differences'], ['precedents', 'Precedents'], ['dependents', 'Dependents'], ['lastCell', 'Last cell'],
  ['visible', 'Visible cells only'], ['cf', 'Conditional formats'], ['dv', 'Data validation']];
const TYPE_LABELS = [['numbers', 'Numbers'], ['text', 'Text'], ['logicals', 'Logicals'], ['errors', 'Errors']];

let last = { kind: 'notes', types: 1 | 2 | 4 | 16, all: false, same: false };

/** Resolves to the choice ({ kind, types, all } as special.js's specialCells takes it), or null for Cancel and Esc. */
export function specialBox(root) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-special-box');
    const make = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids); return e; };
    const h = make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Go To Special' });
    const set = make('fieldset', { className: 'oxe-special-choices' }, make('legend', { textContent: 'Select' }));
    const radios = new Map();
    const option = (input, label) => make('div', { className: 'oxe-dialog-option' }, input, make('label', { htmlFor: input.id, textContent: label }));
    const choice = ([key, label]) => {
      const input = make('input', { type: 'radio', name: `${id}-kind`, id: `${id}-${key}`, value: key, checked: kindKey(last.kind) === key });
      if (input.checked) input.dataset.first = '';
      radios.set(key, input);
      return option(input, label);
    };
    const sub = (...lines) => make('div', { className: 'oxe-special-sub' }, ...lines);
    const types = TYPE_LABELS.map(([key, label]) => make('input', { type: 'checkbox', id: `${id}-t-${key}`, checked: !!(last.types & TYPES[key]), title: label }));
    const levels = ['Direct only', 'All levels'].map((label, i) => make('input', { type: 'radio', name: `${id}-levels`, id: `${id}-l${i}`, checked: last.all === (i === 1) }));
    const sames = ['All', 'Same'].map((label, i) => make('input', { type: 'radio', name: `${id}-same`, id: `${id}-s${i}`, checked: last.same === (i === 1) }));
    const left = make('div', { className: 'oxe-special-column' });
    for (const c of LEFT) {
      left.append(choice(c));
      if (c[0] === 'formulas') left.append(sub(...types.map((input, i) => option(input, TYPE_LABELS[i][1]))));
    }
    const right = make('div', { className: 'oxe-special-column' });
    for (const c of RIGHT) {
      right.append(choice(c));
      if (c[0] === 'dependents') right.append(sub(...levels.map((input, i) => option(input, ['Direct only', 'All levels'][i]))));
      if (c[0] === 'dv') right.append(sub(...sames.map((input, i) => option(input, ['All', 'Same'][i]))));
    }
    set.append(make('div', { className: 'oxe-special-columns' }, left, right));
    const picked = () => [...radios].find(([, input]) => input.checked)?.[0] || 'notes';
    const enable = () => {
      const k = picked();
      for (const t of types) t.disabled = k !== 'constants' && k !== 'formulas';
      for (const l of levels) l.disabled = k !== 'precedents' && k !== 'dependents';
      for (const s of sames) s.disabled = k !== 'cf' && k !== 'dv';
    };
    set.addEventListener('change', enable);
    enable();
    const ok = () => {
      const k = picked();
      const sum = types.reduce((n, t, i) => n + (t.checked ? TYPES[TYPE_LABELS[i][0]] : 0), 0);
      last = { kind: k, types: sum, all: levels[1].checked, same: sames[1].checked };
      // (Constants or Formulas with no type ticked finds nothing.)
      const kind = k === 'cf' || k === 'dv' ? `${k}${last.same ? 'Same' : 'All'}` : k;
      done({ kind, types: sum, all: last.all });
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); ok(); } });
    const row = make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null)));
    dlg.append(h, set, row);
  });
}

/** The box's radio for a kind ('cfAll' and 'cfSame' are Conditional formats). */
const kindKey = (kind) => kind.replace(/(All|Same)$/, '');
