// Formulas > Evaluate Formula: Excel's box over the active cell's formula (calc/stepper.js). Reference names the cell;
// Evaluation shows the formula with the next part underlined and the last value worked out in italics; Evaluate works
// out the underlined part, and at the end becomes Restart; Close. Step In and Step Out are there, not yet working.
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
      if (stepper.done) stepper.restart(); else stepper.step();
      draw();
    });
    const stepIn = button('Step In', () => {});
    const stepOut = button('Step Out', () => {});
    stepIn.disabled = true;
    stepOut.disabled = true;
    const draw = () => {
      if (!stepper) { shown.textContent = editText(cell); evaluate.disabled = true; return; }
      // (Underlined: the part Evaluate works out next; italic: the value it just gave, as Excel shows them.)
      shown.replaceChildren(...stepper.pieces().map((p) => {
        let el = document.createTextNode(p.text);
        if (p.last) el = make('em', {}, el);
        if (p.under) el = make('u', {}, el);
        return el;
      }));
      evaluate.textContent = stepper.done ? 'Restart' : 'Evaluate';
    };
    evaluate.dataset.first = '';
    dlg.append(title, ref, make('p', { className: 'oxe-dialog-text', textContent: 'Evaluation:' }), shown, note,
      make('div', { className: 'oxe-dialog-buttons' }, evaluate, stepIn, stepOut, button('Close', () => done(null))));
    draw();
  });
}
