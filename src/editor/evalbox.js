// Formulas > Evaluate Formula: Excel's box over the active cell's formula (calc/stepper.js). Reference names the cell;
// Evaluation shows the formula with the next part underlined and the last value worked out in italics; Evaluate works
// out the underlined part, and at the end becomes Restart; Step In, when the part underlined is a cell holding a
// formula, shows that formula beneath to work out, and Step Out goes back up with its value; Close.
import { modal, button } from './dialogs.js';
import { Stepper } from '../calc/stepper.js';
import { colName } from '../ooxml/refs.js';
import { editText } from '../ooxml/input.js';

const make = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids); return e; };
const quoted = (name) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`);

/** The Evaluate Formula box for the active cell. Resolves when it is closed. */
export function evaluateFormulaBox(env) {
  const { root, state, editing } = env;
  const wb = state.workbook;
  const si = state.sheetIndex;
  const s = editing().selection;
  const cell = wb.sheet(si).cells.get(s.ar)?.get(s.ac);
  wb.calc.build();
  const entry = wb.calc.entries.get(`${si}:${s.ar}:${s.ac}`);
  const stepper = cell?.f != null && entry?.tree ? new Stepper(cell.f, wb.calc.context(entry)) : null;
  // (Step In's levels: the cell's formula, then each formula stepped into, the last being worked out.)
  const levels = stepper ? [{ stepper, ref: `${quoted(wb.sheets[si].name)}!$${colName(s.ac)}$${s.ar + 1}` }] : [];
  const top = () => levels[levels.length - 1];
  /** The formula of the cell the part underlined refers to, as a level to step into, or null. */
  const inner = () => {
    const at = top()?.stepper.stepInTarget();
    if (!at) return null;
    const one = wb.sheet(at.si).cells.get(at.r)?.get(at.c);
    const e = wb.calc.entries.get(`${at.si}:${at.r}:${at.c}`);
    return one?.f != null && e?.tree ? { stepper: new Stepper(one.f, wb.calc.context(e)), ref: `${quoted(wb.sheets[at.si].name)}!$${colName(at.c)}$${at.r + 1}` } : null;
  };
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-evaluate-box');
    const title = make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Evaluate Formula' });
    const ref = make('p', { className: 'oxe-dialog-text', textContent: `Reference: ${quoted(wb.sheets[si].name)}!$${colName(s.ac)}$${s.ar + 1}` });
    const shown = make('div', { id: `${id}-evaluation`, className: 'oxe-evaluate-text', tabIndex: 0 });
    shown.setAttribute('role', 'textbox');
    shown.setAttribute('aria-readonly', 'true');
    shown.setAttribute('aria-label', 'Evaluation');
    shown.setAttribute('aria-live', 'polite');
    const note = make('p', { className: 'oxe-dialog-text oxe-evaluate-note', textContent: 'To show the result of the underlined expression, click Evaluate. The most recent result appears italicized.' });
    const evaluate = button('Evaluate', () => {
      const s0 = top().stepper;
      if (s0.done && levels.length === 1) s0.restart(); else if (!s0.done) s0.step();
      draw();
    });
    const stepIn = button('Step In', () => { const next = inner(); if (next) { levels.push(next); draw(); } });
    // (Step Out: the level stepped into done with, its cell's value put in place of the reference above.)
    const stepOut = button('Step Out', () => { if (levels.length > 1) { levels.pop(); top().stepper.step(); draw(); } });
    const line = (st) => st.pieces().map((p) => {
      let el = document.createTextNode(p.text);
      if (p.last) el = make('em', {}, el);
      if (p.under) el = make('u', {}, el);
      return el;
    });
    const draw = () => {
      if (!stepper) { shown.textContent = editText(cell); evaluate.disabled = true; stepIn.disabled = true; stepOut.disabled = true; return; }
      // (Underlined: the part Evaluate works out next; italic: the value it just gave, as Excel shows them. Each level
      // stepped into on a line of its own, its cell named.)
      shown.replaceChildren(...levels.flatMap((lv, i) => [make('div', { className: 'oxe-evaluate-level' }, ...(i ? [make('span', { className: 'oxe-evaluate-ref', textContent: `${lv.ref} = ` })] : []), ...line(lv.stepper))]));
      ref.textContent = `Reference: ${top().ref}`;
      evaluate.textContent = top().stepper.done && levels.length === 1 ? 'Restart' : 'Evaluate';
      evaluate.disabled = top().stepper.done && levels.length > 1;
      stepIn.disabled = !inner();
      stepOut.disabled = levels.length < 2;
    };
    evaluate.dataset.first = '';
    dlg.append(title, ref, make('p', { className: 'oxe-dialog-text', textContent: 'Evaluation:' }), shown, note,
      make('div', { className: 'oxe-dialog-buttons' }, evaluate, stepIn, stepOut, button('Close', () => done(null))));
    draw();
  });
}
