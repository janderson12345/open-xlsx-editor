// Excel's error button on a cell with an error indicator (ooxml/errorcheck.js): its menu, headed by the error's name,
// with the fixes Excel offers for it (Convert to Number, Copy Formula from Above or Left, Lock Cell), Ignore Error and
// Edit in Formula Bar. Each applies to the selected cells with the same error, as one step of the history.
import { errorFlags } from '../ooxml/errorcheck.js';
import { ignoreErrors, toNumberChanges, copyFormulaChange, inconsistentFrom } from '../ooxml/errorfix.js';
import { areasOf } from './selection.js';

/** The headings Excel's menu has, by rule (and for an error result, by its error). */
const HEADINGS = {
  numberStoredAsText: 'Number Stored as Text',
  twoDigitTextYear: 'Text Date with 2-Digit Year',
  formula: 'Inconsistent Formula',
  formulaRange: 'Formula Omits Adjacent Cells',
  unlockedFormula: 'Unlocked Cell Containing Formula',
};
const ERROR_NAMES = { '#DIV/0!': 'Divide by Zero Error', '#N/A': 'Value Not Available Error', '#VALUE!': 'Error in Value', '#REF!': 'Invalid Cell Reference Error',
  '#NAME?': 'Invalid Name Error', '#NUM!': 'Invalid Number Error', '#NULL!': 'Null Error' };

/** The active cell's error menu at `point` ({ x, y }); false when it has no error indicator. */
export function openErrorMenu(env, point) {
  const { state, editing, menu } = env;
  const wb = state.workbook;
  const si = state.sheetIndex;
  const s = editing().selection;
  const flags = errorFlags(wb, si);
  const rule = flags.get(`${s.ar},${s.ac}`)?.[0];
  if (!rule) return false;
  const ro = state.readOnly;
  // (The selected cells with the same error; the active cell alone when the selection is large.)
  const cells = [];
  for (const g of areasOf(s)) {
    if ((g.r1 - g.r0 + 1) * (g.c1 - g.c0 + 1) > 100000) continue;
    for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) if (flags.get(`${r},${c}`)?.includes(rule)) cells.push({ r, c });
  }
  if (!cells.length) cells.push({ r: s.ar, c: s.ac });
  const cell = wb.sheet(si).cells.get(s.ar)?.get(s.ac);
  const heading = rule === 'evalError' ? ERROR_NAMES[cell?.v] || 'Error' : HEADINGS[rule];
  const contents = (list) => () => editing().run('applyChanges', [{ si, contents: list, styles: [] }]);
  const fixes = [];
  if (rule === 'numberStoredAsText') fixes.push({ label: 'Convert to Number', run: contents(toNumberChanges(wb, si, cells)), disabled: ro });
  if (rule === 'formula') {
    const way = inconsistentFrom(wb, si, s.ar, s.ac);
    if (way) {
      const changes = cells.map(({ r, c }) => copyFormulaChange(wb, si, r, c, way.dr, way.dc)).filter(Boolean);
      fixes.push({ label: way.dr ? 'Copy Formula from Above' : 'Copy Formula from Left', run: contents(changes), disabled: ro });
    }
  }
  // (Lock Cell: the selection locked, as Format Cells > Protection does.)
  if (rule === 'unlockedFormula') fixes.push({ label: 'Lock Cell', run: () => editing().run('format', { protection: { locked: true } }), disabled: ro });
  // (Focus on the grid first: the menu gives it back there when it closes.)
  state.grid.el.focus({ preventScroll: true });
  menu.open([
    { label: heading, heading: true },
    ...fixes,
    { label: 'Ignore Error', run: () => editing().run('sheetChange', (book, i) => ignoreErrors(book, i, cells, rule)), disabled: ro },
    { label: 'Edit in Formula Bar', run: () => env.root.querySelector('.oxe-formula')?.focus(), disabled: ro },
  ], point, { label: `${heading}: error options` });
  return true;
}
