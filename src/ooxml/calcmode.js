// Formulas > Calculation Options, as Excel has them and saves them in workbook.xml's calcPr (measured: Excel saves
// Manual as calcMode="manual", Automatic Except for Data Tables as calcMode="autoNoTable", Automatic as no calcMode;
// "Recalculate workbook before saving" off as calcOnSave="0"). In Manual a change works out only a formula just
// entered; what reads the change waits for Calculate Now (F9) or Calculate Sheet (Shift+F9), as in Excel, and the
// status bar says Calculate meanwhile. Going back to Automatic works the waiting formulas out, and saving does too
// unless calcOnSave is off. There are no data tables here, so autoNoTable is Automatic.
import { kid, attr } from './xml.js';

export const CALC_MODES = ['auto', 'autoNoTable', 'manual'];

/** workbook.xml's calcPr, made where there is none (after definedNames, before oleSize and the rest). */
function calcPr(wb, make) {
  const root = wb.doc.documentElement;
  let pr = kid(root, 'calcPr', wb.ns);
  if (!pr && make) {
    pr = wb.doc.createElementNS(wb.ns, root.prefix ? `${root.prefix}:calcPr` : 'calcPr');
    const AFTER = ['oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];
    const next = [...root.childNodes].find((n) => n.nodeType === 1 && AFTER.includes(n.localName));
    root.insertBefore(pr, next || null);
    pr.setAttribute('calcId', '191029');
  }
  return pr;
}

/** The workbook's calculation mode: 'auto', 'autoNoTable' or 'manual'. */
export function calcMode(wb) {
  const m = attr(calcPr(wb, false), 'calcMode');
  return CALC_MODES.includes(m) ? m : 'auto';
}
/** Whether a manual workbook is worked out before it is saved (Excel's default, yes). */
export function calcOnSave(wb) { return attr(calcPr(wb, false), 'calcOnSave') !== '0'; }

/** Change the calculation mode; to Automatic, the formulas waiting are worked out. Returns the engine's answer
 *  ({ circular }) when something was worked out, else null. */
export function setCalcMode(wb, mode) {
  if (!CALC_MODES.includes(mode)) throw new Error(`No calculation mode ${mode}`);
  if (mode === calcMode(wb)) return null;
  const pr = calcPr(wb, mode !== 'auto');
  if (mode === 'auto') pr?.removeAttribute('calcMode'); else pr.setAttribute('calcMode', mode);
  wb.bookDirty = true;
  wb.edited = true;
  return mode !== 'manual' ? calculateNow(wb) : null;
}
/** Recalculate workbook before saving, on or off (calcOnSave="0" when off). */
export function setCalcOnSave(wb, on) {
  const pr = calcPr(wb, !on);
  if (on) pr?.removeAttribute('calcOnSave'); else pr.setAttribute('calcOnSave', '0');
  wb.bookDirty = true;
  wb.edited = true;
}

/** Iterative calculation (Excel's File > Options > Formulas): { on, count, delta }, from calcPr's iterate,
 *  iterateCount (100 when left out) and iterateDelta (0.001). */
export function iteration(wb) {
  const pr = calcPr(wb, false);
  const count = Number(attr(pr, 'iterateCount') ?? 100);
  const delta = Number(attr(pr, 'iterateDelta') ?? 0.001);
  return { on: ['1', 'true'].includes(attr(pr, 'iterate')), count: count > 0 ? count : 100, delta: delta >= 0 ? delta : 0.001 };
}
/** Iterative calculation on or off, with its Maximum Iterations and Maximum Change, written as Excel writes them
 *  (measured: iterate="1", the other two only when not 100 and 0.001); the circular formulas worked out again. */
export function setIteration(wb, { on, count = iteration(wb).count, delta = iteration(wb).delta }) {
  if (!Number.isInteger(count) || count < 1 || count > 32767 || !(delta >= 0)) throw new Error('No such iteration settings');
  const pr = calcPr(wb, true);
  const put = (name, value, dflt) => (value === dflt ? pr.removeAttribute(name) : pr.setAttribute(name, String(value)));
  put('iterate', on ? 1 : 0, 0);
  put('iterateCount', count, 100);
  put('iterateDelta', delta, 0.001);
  wb.bookDirty = true;
  wb.edited = true;
  return calculateNow(wb);
}

/** After cells changed (`rects`) and formulas were entered (`also`): everything they reach worked out, or in Manual
 *  the entered formulas alone, the rest waiting. */
export function recalcAfter(wb, rects, also) {
  if (calcMode(wb) !== 'manual') return wb.calc.recalc(rects, also);
  (wb._waiting ||= []).push(...rects);
  return wb.calc.recalc([], also, { only: true });
}
/** Whether formulas are waiting to be worked out (the status bar's Calculate). */
export const needsCalculation = (wb) => !!wb._waiting?.length;

/** Calculate Now (F9): what the changes since reach, and the volatile formulas, worked out. */
export function calculateNow(wb) {
  const rects = wb._waiting || [];
  wb._waiting = [];
  return wb.calc.recalc(rects, []);
}
/** Calculate Sheet (Shift+F9): the same, on sheet `si` alone; the other sheets' formulas go on waiting. */
export function calculateSheet(wb, si) {
  return wb.calc.recalc(wb._waiting || [], [], { sheet: si });
}
/** Every formula of the workbook (Ctrl+Alt+F9), nothing left waiting. */
export function calculateFull(wb) {
  wb._waiting = [];
  return wb.calc.calculateAll();
}
/** Before saving: a manual workbook worked out, unless "Recalculate workbook before saving" is off. True when it was. */
export function calculateBeforeSave(wb) {
  if (calcMode(wb) !== 'manual' || !calcOnSave(wb) || !needsCalculation(wb)) return false;
  calculateNow(wb);
  return true;
}
