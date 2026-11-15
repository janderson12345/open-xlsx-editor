// AutoFilter's other filters, as Excel does them (measured over COM: tools/excel-compare/excel-filter-more.ps1,
// test/filter-more-excel.json): Text and Number Filters (one or two conditions, And or Or, * and ? wildcards), Top 10
// (items or percent, top or bottom), Above and Below Average, the Date Filters (today, this week, last month, this
// year... from today's date), and Filter by Color (a fill or a font colour). A column's filter is read from its
// filterColumn, tested row by row, and written back as Excel writes it.
import { kids, kid, attr } from './xml.js';
import { resolveColor } from './theme.js';

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const OPS = ['equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual'];
/** The date filters and the [from, to) span of serial days each covers, from today's serial `t` (weeks from Sunday,
 *  as Excel's: excel-filter-more.ps1 saw this week as 46334 to 46341 for Wednesday 46337). */
const day = (serial) => { const d = new Date(Date.UTC(1899, 11, 30) + serial * 86400000); return { y: d.getUTCFullYear(), m: d.getUTCMonth(), w: d.getUTCDay() }; };
const serialOf = (y, m, dd = 1) => Math.round((Date.UTC(y, m, dd) - Date.UTC(1899, 11, 30)) / 86400000);
/** Excel 365's namespace for a filterColumn's date groups in its extLst, and the ext's uri. */
export const XLRD2 = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2';
export const DATE_EXT = '{1AD28BCE-077C-4C59-8B6E-1921CE8616D4}';
/** Whether a date's parts ({ y, m, d }) fall in one of `dates` ([{ year, month?, day? }]). */
export const inDateGroups = (p, dates) => dates.some((g) => g.year === p.y && (g.month == null || g.month === p.m) && (g.day == null || g.day === p.d));
/** A filterColumn's date groups (and values) as Excel 365 writes them: in its extLst, xlrd2's filters. */
export function dateFiltersXml(col) {
  const level = (g) => (g.day != null ? 'day' : g.month != null ? 'month' : 'year');
  const items = col.dates.map((g) => `<xlrd2:dateGroupItem year="${g.year}"${g.month != null ? ` month="${g.month}"` : ''}${g.day != null ? ` day="${g.day}"` : ''} dateTimeGrouping="${level(g)}"/>`).join('');
  const values = (col.values || []).map((v) => `<xlrd2:filter val="${esc(v)}"/>`).join('');
  return `<extLst><ext uri="${DATE_EXT}" xmlns:xlrd2="${XLRD2}"><xlrd2:filterColumn><xlrd2:filters${col.blank ? ' blank="1"' : ''}>${values}${items}</xlrd2:filters></xlrd2:filterColumn></ext></extLst>`;
}

export const DATE_SPANS = {
  today: (t) => [t, t + 1],
  yesterday: (t) => [t - 1, t],
  tomorrow: (t) => [t + 1, t + 2],
  thisWeek: (t) => { const s = t - day(t).w; return [s, s + 7]; },
  lastWeek: (t) => { const s = t - day(t).w - 7; return [s, s + 7]; },
  nextWeek: (t) => { const s = t - day(t).w + 7; return [s, s + 7]; },
  thisMonth: (t) => { const { y, m } = day(t); return [serialOf(y, m), serialOf(y, m + 1)]; },
  lastMonth: (t) => { const { y, m } = day(t); return [serialOf(y, m - 1), serialOf(y, m)]; },
  nextMonth: (t) => { const { y, m } = day(t); return [serialOf(y, m + 1), serialOf(y, m + 2)]; },
  thisQuarter: (t) => { const { y, m } = day(t); const q = m - (m % 3); return [serialOf(y, q), serialOf(y, q + 3)]; },
  lastQuarter: (t) => { const { y, m } = day(t); const q = m - (m % 3); return [serialOf(y, q - 3), serialOf(y, q)]; },
  nextQuarter: (t) => { const { y, m } = day(t); const q = m - (m % 3); return [serialOf(y, q + 3), serialOf(y, q + 6)]; },
  thisYear: (t) => { const { y } = day(t); return [serialOf(y, 0), serialOf(y + 1, 0)]; },
  lastYear: (t) => { const { y } = day(t); return [serialOf(y - 1, 0), serialOf(y, 0)]; },
  nextYear: (t) => { const { y } = day(t); return [serialOf(y + 1, 0), serialOf(y + 2, 0)]; },
  yearToDate: (t) => { const { y } = day(t); return [serialOf(y, 0), t + 1]; },
};
/** Today's serial date, local. */
export const todaySerial = () => { const d = new Date(); return serialOf(d.getFullYear(), d.getMonth(), d.getDate()); };

/** A filterColumn as the model has it: { values, blank } for a list of values, or { custom: { and, items: [{ op,
 *  val }] } }, { top10: { top, percent, val } }, { dynamic: { type } }, { color: { dxfId, font } }; `other` for a kind
 *  not read here (icons). */
export function readFilterColumn(fc, ns) {
  // (Dates by year, month or day: dateGroupItems, in filters or, as Excel 365 writes them, in the filterColumn's extLst.)
  const rich = [...fc.getElementsByTagNameNS(XLRD2, 'filters')][0];
  const filters = rich || kid(fc, 'filters', ns);
  const fns = rich ? XLRD2 : ns;
  if (filters && [...filters.childNodes].every((n) => n.nodeType !== 1 || n.localName === 'filter' || n.localName === 'dateGroupItem')) {
    const dates = kids(filters, 'dateGroupItem', fns).map((d) => ({ year: Number(attr(d, 'year')), ...(attr(d, 'month') ? { month: Number(attr(d, 'month')) } : {}), ...(attr(d, 'day') ? { day: Number(attr(d, 'day')) } : {}) }));
    return { values: kids(filters, 'filter', fns).map((f) => attr(f, 'val') ?? ''), blank: ['1', 'true'].includes(attr(filters, 'blank')), ...(dates.length ? { dates } : {}), other: false };
  }
  const custom = kid(fc, 'customFilters', ns);
  if (custom) return { custom: { and: ['1', 'true'].includes(attr(custom, 'and')), items: kids(custom, 'customFilter', ns).map((f) => ({ op: attr(f, 'operator') || 'equal', val: attr(f, 'val') ?? '' })) }, other: false };
  const top = kid(fc, 'top10', ns);
  if (top) return { top10: { top: attr(top, 'top') !== '0', percent: attr(top, 'percent') === '1', val: Number(attr(top, 'val') || 10) }, other: false };
  const dyn = kid(fc, 'dynamicFilter', ns);
  if (dyn && (DATE_SPANS[attr(dyn, 'type')] || /^(above|below)Average$/.test(attr(dyn, 'type')))) return { dynamic: { type: attr(dyn, 'type') }, other: false };
  const color = kid(fc, 'colorFilter', ns);
  if (color) return { color: { dxfId: Number(attr(color, 'dxfId') || 0), font: attr(color, 'cellColor') === '0' }, other: false };
  return { values: [], blank: false, other: true };
}

/** Excel's Custom AutoFilter criteria as a column's filter, as COM's Range.AutoFilter takes them: "=x" (a value list
 *  when it has no wildcard), "<>x", ">x", ">=x", "<x", "<=x"; `join` 'and' or 'or' a second one. */
export function fromCriteria(c1, join = null, c2 = null) {
  const item = (c) => {
    const m = /^(<>|>=|<=|=|>|<)?(.*)$/s.exec(String(c));
    return { op: { '<>': 'notEqual', '>=': 'greaterThanOrEqual', '<=': 'lessThanOrEqual', '>': 'greaterThan', '<': 'lessThan' }[m[1]] || 'equal', val: m[2] };
  };
  const a = item(c1);
  if (!c2 && a.op === 'equal' && !/[*?]/.test(a.val)) return { values: [a.val], blank: a.val === '', other: false };
  return { custom: { and: join === 'and', items: c2 ? [a, item(c2)] : [a] }, other: false };
}

/** The filterColumn's content Excel writes for a column (its colour filter's dxf already made: `dxfId`). */
export function filterColumnXml(col, { values = [], average = null, today = todaySerial(), dxfId = 0 } = {}) {
  if (col.custom) {
    const items = col.custom.items.map((x) => `<customFilter${x.op === 'equal' ? '' : ` operator="${x.op}"`} val="${esc(x.val)}"/>`).join('');
    return `<customFilters${col.custom.and && col.custom.items.length > 1 ? ' and="1"' : ''}>${items}</customFilters>`;
  }
  if (col.top10) {
    const { top, percent, val } = col.top10;
    return `<top10${top ? '' : ' top="0"'}${percent ? ' percent="1"' : ''} val="${val}" filterVal="${topThreshold(values, col.top10)}"/>`;
  }
  if (col.dynamic) {
    const { type } = col.dynamic;
    if (DATE_SPANS[type]) { const [a, b] = DATE_SPANS[type](today); return `<dynamicFilter type="${type}" val="${a}" maxVal="${b}"/>`; }
    return `<dynamicFilter type="${type}" val="${average}"/>`;
  }
  if (col.color) return `<colorFilter dxfId="${dxfId}"${col.color.font ? ' cellColor="0"' : ''}/>`;
  return null;
}

/** A Top 10 filter's threshold among the column's numbers: the last number kept (items: the nth; percent: n% of the
 *  count, at least one: Excel's 50% of 7 keeps 3). */
function topThreshold(values, t) {
  const nums = values.filter((v) => typeof v === 'number').sort((a, b) => (t.top ? b - a : a - b));
  if (!nums.length) return 0;
  const n = t.percent ? Math.max(1, Math.floor((nums.length * t.val) / 100)) : Math.max(1, t.val);
  return nums[Math.min(n, nums.length) - 1];
}
/** A wildcard pattern (* any run, ? any one, ~ escaping) as a RegExp, case aside. */
const pattern = (v) => new RegExp(`^${String(v).replace(/~([*?~])|([*?])|([.+^${}()|[\]\\])/g, (m, lit, wild, re) => (lit ? `\\${lit}` : wild === '*' ? '[\\s\\S]*' : wild ? '[\\s\\S]' : `\\${re}`))}$`, 'i');

/**
 * A test of row r for column c's filter `col` (not a value list): `value(r)` the cell's value, `text(r)` its shown
 * text, `rows` the filter's rows (for Top 10 and the average), `cellLook(r)` its fill and font colours
 * ({ fill, font } as "#RRGGBB"), `dxfColor(id)` a colour filter's colour, `today` the serial date.
 */
export function columnTest(col, { value, text, rows, cellLook, dxfColor, today = todaySerial() }) {
  if (col.custom) {
    const one = ({ op, val }) => (r) => {
      const v = value(r);
      const n = Number(val);
      const numeric = val.trim() !== '' && Number.isFinite(n);
      if (op === 'equal' || op === 'notEqual') {
        const hit = numeric && typeof v === 'number' ? v === n : pattern(val).test(text(r));
        return op === 'equal' ? hit : !hit;
      }
      // (Comparisons: numbers with numbers, text with text by its letters, case aside; a blank never.)
      if (v == null || v === '') return false;
      let d;
      if (numeric) { if (typeof v !== 'number') return false; d = v - n; } else { if (typeof v === 'number') return false; d = text(r).toLowerCase().localeCompare(val.toLowerCase()); }
      return op === 'greaterThan' ? d > 0 : op === 'greaterThanOrEqual' ? d >= 0 : op === 'lessThan' ? d < 0 : d <= 0;
    };
    const tests = col.custom.items.map(one);
    return (r) => (col.custom.and ? tests.every((t) => t(r)) : tests.some((t) => t(r)));
  }
  if (col.top10) {
    const at = topThreshold(rows.map(value), col.top10);
    return (r) => { const v = value(r); return typeof v === 'number' && (col.top10.top ? v >= at : v <= at); };
  }
  if (col.dynamic) {
    const { type } = col.dynamic;
    if (DATE_SPANS[type]) { const [a, b] = DATE_SPANS[type](today); return (r) => { const v = value(r); return typeof v === 'number' && Math.floor(v) >= a && Math.floor(v) < b; }; }
    const nums = rows.map(value).filter((v) => typeof v === 'number');
    const avg = nums.reduce((s, v) => s + v, 0) / (nums.length || 1);
    return (r) => { const v = value(r); return typeof v === 'number' && (type === 'aboveAverage' ? v > avg : v < avg); };
  }
  if (col.color) {
    const want = (dxfColor(col.color.dxfId) || '').toUpperCase();
    return (r) => ((col.color.font ? cellLook(r).font : cellLook(r).fill) || '').toUpperCase() === want;
  }
  return () => true;
}

/** The colour of a colour filter's dxf ("#RRGGBB"): its fill's foreground, as Excel writes both kinds. */
export function dxfFillColor(wb, id) {
  const sroot = wb.stylesDoc?.documentElement;
  const dxf = sroot && kids(kid(sroot, 'dxfs', sroot.namespaceURI), 'dxf', sroot.namespaceURI)[id];
  const fg = dxf && dxf.getElementsByTagNameNS(sroot.namespaceURI, 'fgColor')[0];
  if (!fg) return null;
  const c = { rgb: attr(fg, 'rgb'), theme: attr(fg, 'theme') != null ? Number(attr(fg, 'theme')) : null, indexed: attr(fg, 'indexed') != null ? Number(attr(fg, 'indexed')) : null, tint: Number(attr(fg, 'tint') || 0) };
  return resolveColor(c, wb.theme, wb.styles.palette);
}

/** The dxf Excel writes for a colour filter (both kinds: the colour as a solid fill's foreground, white behind). */
export const colorDxfXml = (rgb) => `<dxf><fill><patternFill patternType="solid"><fgColor rgb="${rgb}"/><bgColor rgb="FFFFFFFF"/></patternFill></fill></dxf>`;
export { OPS as FILTER_OPS };
