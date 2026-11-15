// The ribbon's sparkline actions (acts.js): Insert > Sparklines and the Sparkline tab, each through the editor's
// sheetChange command as one step of the history, the sheet's XML changed by ooxml/sparkedit.js as Excel changes it.
import { rangesBox, hiddenEmptyBox } from './sparkboxes.js';
import { numberBox } from './dialogs.js';
import { areasOf } from './selection.js';
import { parseRange, rangeName } from '../ooxml/refs.js';
import { parseSheetRange } from '../ooxml/sparklines.js';
import { sparkPieces, addSparklines, setGroupOptions, setGroupColor, clearSparklines, groupSparklines, ungroupSparklines, setGroupData, setSparklineData, SparklineError, applyStyle } from '../ooxml/sparkedit.js';
import { sheetPrefix } from '../calc/tokens.js';

const TYPES = { line: 'Line', column: 'Column', stacked: 'Win/Loss' };
/** A range as Excel's boxes show the location: $O$2:$O$4. */
const absolute = (a) => rangeName(a).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');

export function createSparklineActs(env) {
  const { root, state, editing } = env;
  const wb = () => state.workbook;
  const si = () => state.sheetIndex;
  const sheetName = () => wb().sheets[si()].name;
  const active = () => { const s = editing().selection; return [s.ar, s.ac]; };
  /** The group drawn in the active cell (sparklines.js readGroup), or null. */
  const groupHere = () => {
    const [r, c] = active();
    return wb().sheet(si()).sparklines?.find((g) => g.sparklines.some((s) => s.r === r && s.c === c)) || null;
  };
  /** A typed reference ({ sheet, range }): Sheet1!$B$2:$M$2, 'My data'!B2:M2, or B2:M2 on the sheet shown. */
  const refOf = (text) => {
    const t = String(text || '').trim().replace(/^=/, '');
    if (!t) return null;
    if (t.includes('!')) { const p = parseSheetRange(t); return p && wb().sheetIndexOf(p.sheet) >= 0 ? p : null; }
    const range = parseRange(t.replace(/\$/g, ''));
    return range && { sheet: sheetName(), range };
  };
  /** The location typed: a range of the sheet shown. */
  const locationOf = (text) => { const p = refOf(text); return p && p.sheet === sheetName() ? p.range : null; };
  /** A refusal for a data and location answer, or null (sparkedit.js's rules, Excel's). */
  const refusal = (data, loc) => {
    if (!data) return 'Data Source reference is not valid.';
    if (!loc) return 'Location reference is not valid.';
    try { sparkPieces(loc, data); return null; } catch (e) { if (e instanceof SparklineError) return e.message; throw e; }
  };
  /** A change to the sheet's sparklines, one step of the history; false when `fn` changed nothing. */
  const change = (fn) => editing().run('sheetChange', (book, k) => book.changeSparklines(k, fn) !== false, 'always') && (state.grid.el.focus({ preventScroll: true }), true);
  const home = (x) => { state.grid.el.focus({ preventScroll: true }); return x; };

  const acts = {
    /** The group at the active cell, for the Sparkline tab. */
    sparklineAt: () => groupHere(),
    /** Insert > Sparklines (`type`: line, column, stacked): Create Sparklines, the selection as its data, or (empty
     *  cells selected) as its location. */
    sparklineInsert(type) {
      const s = editing().selection;
      const sheet = wb().sheet(si());
      let filled = false;
      for (let r = s.r0; r <= Math.min(s.r1, s.r0 + 500) && !filled; r++) for (let c = s.c0; c <= Math.min(s.c1, s.c0 + 500); c++) if (sheet.cells.get(r)?.get(c)?.v != null) { filled = true; break; }
      const sel = { r0: s.r0, c0: s.c0, r1: s.r1, c1: s.c1 };
      return rangesBox(root, {
        title: 'Create Sparklines', text: `Choose the data that you want (${TYPES[type]})`,
        fields: [{ key: 'data', label: 'Data Range:', value: filled ? rangeName(sel) : '' }, { key: 'location', label: 'Location Range:', value: filled ? '' : absolute(sel) }],
        check: (a) => refusal(refOf(a.data), locationOf(a.location)),
      }).then((a) => (a ? change((doc) => addSparklines(doc, type, locationOf(a.location), refOf(a.data))) : home(false)));
    },
    /** Edit Data > Edit Group Location & Data...: the group's data and location as one range each. */
    sparklineGroupData() {
      const g = groupHere();
      if (!g) return false;
      const box = (list) => ({ r0: Math.min(...list.map((a) => a.r0)), c0: Math.min(...list.map((a) => a.c0)), r1: Math.max(...list.map((a) => a.r1)), c1: Math.max(...list.map((a) => a.c1)) });
      const sources = g.sparklines.map((s) => parseSheetRange(s.source)).filter(Boolean);
      const data = sources.length ? `${sources[0].sheet === sheetName() ? '' : sheetPrefix(sources[0].sheet)}${rangeName(box(sources.map((p) => p.range)))}` : '';
      const loc = absolute(box(g.sparklines.map((s) => ({ r0: s.r, c0: s.c, r1: s.r, c1: s.c }))));
      const [r, c] = active();
      return rangesBox(root, {
        title: 'Edit Sparklines', text: 'Choose the data that you want',
        fields: [{ key: 'data', label: 'Data Range:', value: data }, { key: 'location', label: 'Location Range:', value: loc }],
        check: (a) => refusal(refOf(a.data), locationOf(a.location)),
      }).then((a) => (a ? change((doc) => setGroupData(doc, r, c, locationOf(a.location), refOf(a.data))) : home(false)));
    },
    /** Edit Data > Edit Single Sparkline's Data...: the active cell's sparkline alone. */
    sparklineSingleData() {
      const g = groupHere();
      if (!g) return false;
      const [r, c] = active();
      const here = g.sparklines.find((s) => s.r === r && s.c === c);
      const p = parseSheetRange(here.source);
      return rangesBox(root, {
        title: 'Edit Sparkline Data', fields: [{ key: 'data', label: `Select a new source data range for the sparkline (${absolute({ r0: r, c0: c, r1: r, c1: c })}):`, value: p ? `${p.sheet === sheetName() ? '' : sheetPrefix(p.sheet)}${rangeName(p.range)}` : here.source }],
        check: (a) => refusal(refOf(a.data), { r0: r, c0: c, r1: r, c1: c }),
      }).then((a) => (a ? change((doc) => setSparklineData(doc, r, c, refOf(a.data))) : home(false)));
    },
    /** Edit Data > Hidden & Empty Cells...: Hidden and Empty Cell Settings. */
    sparklineEmpty() {
      const g = groupHere();
      if (!g) return false;
      const [r, c] = active();
      return hiddenEmptyBox(root, { empty: g.empty, hidden: g.hidden, line: g.type === 'line' })
        .then((a) => (a ? change((doc) => setGroupOptions(doc, r, c, a)) : home(false)));
    },
    /** Axis > Date Axis Type...: the dates' range, as many cells as each sparkline's data. */
    sparklineDateAxis() {
      const g = groupHere();
      if (!g) return false;
      const [r, c] = active();
      const p = g.dates ? parseSheetRange(g.dates) : null;
      const size = (a) => (a.r1 - a.r0 + 1) * (a.c1 - a.c0 + 1);
      const want = parseSheetRange(g.sparklines[0].source);
      return rangesBox(root, {
        title: 'Sparkline Date Range', fields: [{ key: 'dates', label: 'Select the range that contains the date values for the sparkline data:', value: p ? rangeName(p.range) : '' }],
        check: (a) => { const d = refOf(a.dates); return !d ? 'Date range reference is not valid.' : want && size(d.range) !== size(want.range) ? 'The date range must contain as many cells as the sparkline data.' : null; },
      }).then((a) => { if (!a) return home(false); const d = refOf(a.dates); return change((doc) => setGroupOptions(doc, r, c, { dates: sheetPrefix(d.sheet) + rangeName(d.range) })); });
    },
    /** Axis > Custom Value... for the vertical axis's minimum or maximum (`which`: 'min' or 'max'). */
    sparklineAxisValue(which) {
      const g = groupHere();
      if (!g) return false;
      const [r, c] = active();
      return numberBox(root, { title: 'Sparkline Vertical Axis Settings', label: `Enter the ${which === 'min' ? 'minimum' : 'maximum'} value for the vertical axis:`, value: String(g[which] ?? 0), min: -1e307, max: 1e307, refusal: 'Enter a number.' })
        .then((n) => (n == null ? home(false) : change((doc) => setGroupOptions(doc, r, c, { [`${which}Type`]: 'custom', [which]: n }))));
    },
    /** The tab's settings (sparkedit.js setGroupOptions's names) for the active cell's group. */
    sparklineSet(changes) { const [r, c] = active(); return change((doc) => setGroupOptions(doc, r, c, changes)); },
    /** Sparkline Color and Marker Color's colours (`name`: series, negative, markers, high, low, first, last). */
    sparklineColor(name, color) { const [r, c] = active(); return color ? change((doc) => setGroupColor(doc, r, c, name, color)) : false; },
    /** A style of the gallery (0 to 35) for the selected sparklines' groups. */
    sparklineStyle(k) { const [r, c] = active(); return change((doc) => applyStyle(doc, areasOf(editing().selection), r, c, k)); },
    /** Group: the selected sparklines into the active cell's group. */
    sparklineGroup() { const [r, c] = active(); return change((doc) => groupSparklines(doc, areasOf(editing().selection), r, c)); },
    /** Ungroup: each selected sparkline a group of its own. */
    sparklineUngroup() { return change((doc) => ungroupSparklines(doc, areasOf(editing().selection))); },
    /** Clear Selected Sparklines, or (`whole`) Clear Selected Sparkline Groups. */
    sparklineClear(whole = false) { return change((doc) => clearSparklines(doc, areasOf(editing().selection), whole) > 0); },
  };
  return (name, args) => (acts[name] ? acts[name](...args) : undefined);
}
