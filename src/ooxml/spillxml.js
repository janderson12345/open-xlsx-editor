// Dynamic array formulas in the sheet's XML, as Excel 365 writes them (measured: tools/excel-compare/excel-dynamic.ps1):
// the formula's cell `<c cm="1"><f t="array" ref="D1:D4">…</f>` (cm pointing at metadata.xml's dynamic array entry,
// dynamicmeta.js), the cells it spills into holding values only, and a formula whose result isn't an array a plain
// `<f>`. #SPILL! and #CALC! are rich error values Excel keeps in other parts; written here as #VALUE!, with the workbook
// marked for Excel to work its formulas out again on opening, which shows them as Excel's own.
import { kid } from './xml.js';
import { rangeName } from './refs.js';
import { setCellXml, cellEl } from './sheetedit.js';
import { dynamicCm } from './dynamicmeta.js';

/** The formula in (r, c) of sheet `i` a dynamic array formula spilling into `rect` (its own cell for one value, or a
 *  blocked spill), or a plain formula (rect null). */
export function setDynamicXml(wb, i, r, c, rect) {
  const part = wb._part(i);
  const cell = part.model.cells.get(r)?.get(c);
  const el = cell ? cellEl(part.doc, r, c) : null;
  const f = el && kid(el, 'f', el.namespaceURI);
  if (!cell || !f) return;
  const ref = rect ? rangeName(rect) : null;
  if (rect) {
    if (f.getAttribute('t') === 'array' && f.getAttribute('ref') === ref && el.getAttribute('cm')) return;
    f.setAttribute('t', 'array');
    f.setAttribute('ref', ref);
    el.setAttribute('cm', String(dynamicCm(wb)));
  } else {
    if (!el.getAttribute('cm') && f.getAttribute('t') !== 'array') return;
    f.removeAttribute('t');
    f.removeAttribute('ref');
    el.removeAttribute('cm');
  }
  for (const a of ['aca', 'ca']) f.removeAttribute(a);
  cell.spill = rect;
  wb.dirty.add(wb.sheets[i].path);
}

/** A cell a dynamic formula no longer spills into: its value gone (its format kept). */
export function clearSpilledXml(wb, i, r, c) {
  const part = wb._part(i);
  const cell = part.model.cells.get(r)?.get(c);
  if (!cell || cell.f != null) return;
  setCellXml(part.doc, r, c, null);
  cell.t = 'n';
  cell.v = null;
  wb.dirty.add(wb.sheets[i].path);
}

/** Excel's dynamic array errors, which the file holds as rich values: written as #VALUE! (Excel works them out again). */
export const RICH_ERRORS = new Set(['#SPILL!', '#CALC!']);
