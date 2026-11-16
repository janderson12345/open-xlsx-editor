// The Table Design tab, Excel's contextual tab, shown while the active cell is in a table: Properties (Table Name,
// Resize Table), Tools (Convert to Range), Table Style Options (Header Row, Total Row, Banded Rows, First Column, Last
// Column, Banded Columns, Filter Button) and Table Styles (the gallery, tablegallery.js). Its actions are
// acts-tables.js's.
import { el } from './ribbonkit.js';
import { tableGallery } from './tablegallery.js';

/** The options' boxes: [key, label]. */
const OPTIONS = [['headers', 'Header Row'], ['totals', 'Total Row'], ['rows', 'Banded Rows'], ['first', 'First Column'], ['last', 'Last Column'], ['cols', 'Banded Columns'], ['filterButton', 'Filter Button']];

/** The tab. `act(name, ...args)` runs a command; `workbook()` the workbook (for the gallery's styles and theme). */
export function tableTab(kit, { act, workbook }) {
  const { btn, toggle, grp, col, menu } = kit;
  let t = null;
  const name = el('input', { type: 'text', class: 'oxe-table-name', 'aria-label': 'Table Name', spellcheck: false,
    // (Enter, or leaving the box, commits it: the input's change event, once; Esc puts the name back.)
    onkeydown: (e) => { if (e.key === 'Escape') e.target.value = t?.name || ''; },
    onchange: (e) => act('tableRename', e.target.value.trim()) });
  const on = (key) => (key === 'headers' ? !!t?.headers : key === 'totals' ? !!t?.totals : key === 'filterButton' ? !!t?.filter && !t.filter.hidden.size : !!t?.style[key]);
  const boxes = Object.fromEntries(OPTIONS.map(([key, label]) => [key, toggle('☐', label, () => act('tableOption', key, !on(key)), { label })]));
  const options = () => ({ headers: !!t?.headers, totals: !!t?.totals, ...(t?.style || {}) });
  const gallery = menu('▦', 'Table Styles', () => [el('div', { class: 'oxe-tstyle-menu' }, ...tableGallery(kit, { wb: workbook(), options: options(), onPick: (n) => act('tableStyle', n), withNone: true }))],
    { label: 'Quick Styles', big: true });
  const groups = [
    grp('Properties', col(el('label', { class: 'oxe-cap' }, 'Table Name:'), name, btn('⤡', 'Resize Table', () => act('tableResize'), { label: 'Resize Table' }))),
    grp('Tools', btn('⇲', 'Convert to Range', () => act('tableConvert'), { label: 'Convert to Range', big: true })),
    grp('Table Style Options', col(boxes.headers, boxes.totals, boxes.rows), col(boxes.first, boxes.last, boxes.cols), col(boxes.filterButton)),
    grp('Table Styles', gallery),
  ];
  return {
    groups,
    /** Shown only while the active cell is in a table. */
    contextual: () => !!act('tableAt'),
    update() {
      t = act('tableAt') || null;
      if (document.activeElement !== name) name.value = t?.name || '';
      for (const [key, b] of Object.entries(boxes)) {
        b.setOn(on(key));
        b.querySelector('.oxe-ico').textContent = on(key) ? '☑' : '☐';
        // (The filter button belongs to the header row: Excel greys it without one.)
        b.disabled = key === 'filterButton' && !t?.headers;
      }
    },
  };
}
