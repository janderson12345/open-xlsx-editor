// The filter list's Text Filters, Number Filters, Date Filters and Filter by Color, as one box: Excel's Custom
// AutoFilter (two conditions, And or Or: equals, does not equal, begins with, ends with, contains, does not contain,
// greater and less than; * and ? wildcards), and beside it the choices that take no value: Top 10, Above and Below
// Average for numbers, today, this week, last month... for dates, and the column's fill and font colours. Resolves to
// a column filter (filtermore.js's) or null.
import { modal, button } from './dialogs.js';
import { make, select, labelled } from './formatbox-kit.js';
import { fromCriteria } from '../ooxml/filtermore.js';
import { parseInput } from '../ooxml/input.js';

const OPS = [['', '(none)'], ['=', 'equals'], ['<>', 'does not equal'], ['>', 'is greater than'], ['>=', 'is greater than or equal to'], ['<', 'is less than'],
  ['<=', 'is less than or equal to'], ['begins', 'begins with'], ['notBegins', 'does not begin with'], ['ends', 'ends with'], ['notEnds', 'does not end with'],
  ['contains', 'contains'], ['notContains', 'does not contain']];
const DATE_CHOICES = [['tomorrow', 'Tomorrow'], ['today', 'Today'], ['yesterday', 'Yesterday'], ['nextWeek', 'Next Week'], ['thisWeek', 'This Week'], ['lastWeek', 'Last Week'],
  ['nextMonth', 'Next Month'], ['thisMonth', 'This Month'], ['lastMonth', 'Last Month'], ['nextQuarter', 'Next Quarter'], ['thisQuarter', 'This Quarter'], ['lastQuarter', 'Last Quarter'],
  ['nextYear', 'Next Year'], ['thisYear', 'This Year'], ['lastYear', 'Last Year'], ['yearToDate', 'Year to Date']];

/** One condition as Excel's criteria: the operator and the value typed (a date read as its serial, for a date column). */
function criteria(op, typed, dates) {
  let v = String(typed ?? '');
  if (dates && v.trim()) { const p = parseInput(v.trim()); if (typeof p.value === 'number') v = String(p.value); }
  switch (op) {
    case 'begins': return `=${v}*`;
    case 'notBegins': return `<>${v}*`;
    case 'ends': return `=*${v}`;
    case 'notEnds': return `<>*${v}`;
    case 'contains': return `=*${v}*`;
    case 'notContains': return `<>*${v}*`;
    default: return `${op}${v}`;
  }
}

/** The box for column `title`: `kind` 'text', 'number' or 'date'; `colors` { fill: ["#RRGGBB"], font: [...] }. */
export function customFilterBox(root, { title, kind, colors = { fill: [], font: [] } }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-custom-filter');
    const quick = [['custom', 'Custom (the conditions below)']];
    if (kind === 'number') quick.push(['top', 'Top 10...'], ['aboveAverage', 'Above Average'], ['belowAverage', 'Below Average']);
    if (kind === 'date') quick.push(...DATE_CHOICES);
    for (const c of colors.fill) quick.push([`fill:${c}`, `Cell color ${c}`]);
    for (const c of colors.font) quick.push([`font:${c}`, `Font color ${c}`]);
    const pick = select(quick, 'custom');
    pick.dataset.first = '';
    const op1 = select(OPS.slice(1), kind === 'text' ? '=' : '>');
    const v1 = make('input', { type: 'text', className: 'oxe-dialog-field', spellcheck: false });
    const join = select([['and', 'And'], ['or', 'Or']], 'and');
    const op2 = select(OPS, '');
    const v2 = make('input', { type: 'text', className: 'oxe-dialog-field', spellcheck: false });
    v1.setAttribute('aria-label', 'First value');
    v2.setAttribute('aria-label', 'Second value');
    const topWhich = select([['top', 'Top'], ['bottom', 'Bottom']], 'top');
    const topN = make('input', { type: 'number', min: 1, max: 500, value: '10', className: 'oxe-fc-num' });
    topN.setAttribute('aria-label', 'How many');
    const topUnit = select([['items', 'Items'], ['percent', 'Percent']], 'items');
    const conds = make('fieldset', {}, make('legend', { textContent: 'Show rows where:' }), labelled(title, op1), v1, labelled('', join), op2, v2);
    const top = make('fieldset', {}, make('legend', { textContent: 'Top 10 AutoFilter' }), topWhich, topN, topUnit);
    const sync = () => { conds.hidden = pick.value !== 'custom'; top.hidden = pick.value !== 'top'; };
    pick.addEventListener('change', sync);
    sync();
    const why = make('p', { className: 'oxe-dialog-text', role: 'alert' });
    const ok = () => {
      const p = pick.value;
      if (p === 'top') return done({ top10: { top: topWhich.value === 'top', percent: topUnit.value === 'percent', val: Math.max(1, Math.min(500, Math.round(Number(topN.value) || 10))) }, other: false });
      if (p === 'aboveAverage' || p === 'belowAverage' || DATE_CHOICES.some(([k]) => k === p)) return done({ dynamic: { type: p }, other: false });
      if (p.startsWith('fill:') || p.startsWith('font:')) return done({ color: { rgb: `FF${p.slice(6).toUpperCase()}`, font: p.startsWith('font:') }, other: false });
      if (!v1.value.trim() && op1.value !== '=' && op1.value !== '<>') { why.textContent = 'Type a value to compare with.'; return; }
      const dates = kind === 'date';
      return done(op2.value ? fromCriteria(criteria(op1.value, v1.value, dates), join.value, criteria(op2.value, v2.value, dates)) : fromCriteria(criteria(op1.value, v1.value, dates)));
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: `Custom AutoFilter: ${title}` }), labelled('Filter:', pick), conds, top, why,
      make('p', { className: 'oxe-dialog-text', textContent: 'Use ? for any one character, * for any run of characters.' }),
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
  });
}
