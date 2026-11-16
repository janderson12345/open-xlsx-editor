// A table's own filter (its part's autoFilter, docs/TABLES.md): set from its header's buttons as the sheet's AutoFilter
// is (filterops.js: the same filters, the same rows hidden), written into the table part, where Excel keeps it.
import { serializeXml } from './xml.js';
import { setAutoFilterXml, hiddenRows } from './filterops.js';
import { setRows } from './sizes.js';

/** Table `t`'s filter set to `filter` ({ r0, c0, r1, c1, columns: Map }): its part's autoFilter written (the hidden
 *  buttons kept) and the rows it leaves out hidden, those it lets through shown. */
export function applyTableFilter(wb, t, filter) {
  const doc = wb.pkg.xml(t.path);
  // (A column whose button is hidden keeps its filterColumn: filterops.js keeps an "other" one as it is.)
  const columns = new Map(filter.columns);
  for (const k of t.filter?.hidden || []) if (!columns.has(t.c0 + k)) columns.set(t.c0 + k, { other: true });
  setAutoFilterXml(doc, { ...filter, columns }, wb, t.si, wb.today);
  wb.pkg.setText(t.path, serializeXml(doc));
  wb.dirty.add(t.path);
  const part = wb._part(t.si);
  const hide = hiddenRows(wb, t.si, filter, wb.today);
  for (let r = filter.r0 + 1; r <= filter.r1;) {
    const on = hide.has(r);
    let end = r;
    while (end + 1 <= filter.r1 && hide.has(end + 1) === on) end++;
    setRows(part.doc, r, end, { hidden: on });
    r = end + 1;
  }
  wb._reread(t.si, part);
  wb.resetCalc();
}
