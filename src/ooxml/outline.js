// Outlines (Data > Group and Ungroup, Show Detail, Hide Detail, Clear Outline, the Settings, the outline's level and
// +/- buttons), as Excel writes them (tools/excel-compare/excel-outline.ps1, 34 cases, docs/OUTLINE.md). A row or
// column carries its outlineLevel (1 to 7); a group of level L is a run of lines at level L or more, its summary line
// the one after it (before it with outlinePr summaryBelow="0" or summaryRight="0"). Collapsing hides the group's lines
// (hidden="1", a column also width 0 unless it has a width of its own) and marks the summary line collapsed="1", which
// Excel writes even on a row with nothing else. sheetFormatPr's outlineLevelRow and outlineLevelCol hold the deepest
// level. Excel refuses an eighth level and Ungroup of lines that aren't grouped.
import { filledRows } from './sheetcells.js';
import { kids, kid, attr } from './xml.js';
import { setColumns, setSheetFormat, pxWidth } from './sizes.js';
import { pinRefs, forgetRows } from './sheetedit.js';
import { defaultColPx } from './units.js';
import { MAX_COLS } from './refs.js';
import { setShown } from './view.js';
import { styleSummaryLines, summariesOf, applyOutlineStyles } from './outlinestyles.js';

/** Excel's deepest level: Group refuses an eighth (measured: "Unable to get the Group property"). */
export const MAX_LEVEL = 7;
/** An element's outline level as Excel allows it, 0 to 7, whatever the file says. */
const levelOf = (x) => Math.max(0, Math.min(MAX_LEVEL, Math.trunc(Number(attr(x, 'outlineLevel'))) || 0));
const on = (v) => v === '1' || v === 'true';
const make = (doc, ns, name) => doc.createElementNS(ns, doc.documentElement.prefix ? `${doc.documentElement.prefix}:${name}` : name);

/** The sheet's outline settings: summary rows below their detail and columns right of theirs (outlinePr), and whether
 *  its symbols are shown (sheetView showOutlineSymbols). */
export function readOutlinePr(root, ns, view) {
  const pr = kid(kid(root, 'sheetPr', ns), 'outlinePr', ns);
  const no = (name) => !!pr && (attr(pr, name) === '0' || attr(pr, name) === 'false');
  return { below: !no('summaryBelow'), right: !no('summaryRight'), styles: !!pr && (attr(pr, 'applyStyles') === '1' || attr(pr, 'applyStyles') === 'true'), symbols: !view || !(attr(view, 'showOutlineSymbols') === '0' || attr(view, 'showOutlineSymbols') === 'false') };
}

/** The groups of an outline, from `levels` ([[index, level]] of the lines with a level, any order): for each level L
 *  every run of lines at L or more, { level, start, end, summary } (summary the line after the run, or before it when
 *  `after` is false; null off the sheet's start). */
export function outlineGroups(levels, after = true) {
  const list = levels.filter(([, l]) => l > 0).sort((a, b) => a[0] - b[0]);
  const out = [];
  for (let L = 1; L <= MAX_LEVEL; L++) {
    let run = null;
    for (const [i, l] of list) {
      if (l < L) continue;
      if (run && i === run.end + 1) { run.end = i; continue; }
      if (run) out.push(run);
      run = { level: L, start: i, end: i };
    }
    if (run) out.push(run);
  }
  for (const g of out) g.summary = after ? g.end + 1 : g.start > 0 ? g.start - 1 : null;
  return out;
}

const ROW_ORDER = ['r', 'spans', 's', 'customFormat', 'ht', 'hidden', 'customHeight', 'outlineLevel', 'collapsed', 'thickTop', 'thickBot', 'ph'];
const COL_ORDER = ['min', 'max', 'width', 'style', 'hidden', 'bestFit', 'customWidth', 'phonetic', 'outlineLevel', 'collapsed'];
/** An element's attributes put in its type's order (CT_Row, CT_Col), others after them as they came. */
function reorder(el, order) {
  const all = [...el.attributes].map((a) => [a.name, a.value]);
  const rank = (k) => { const i = order.indexOf(k); return i < 0 ? order.length : i; };
  for (const [k] of all) el.removeAttribute(k);
  for (const [k, v] of all.map((x, n) => [...x, n]).sort((a, b) => rank(a[0]) - rank(b[0]) || a[2] - b[2])) el.setAttribute(k, v);
}

/** One axis of a sheet's XML as lines: get(i) { level, collapsed, hidden }, set(i, patch), levels() [[i, level]], and
 *  finish() (attributes ordered, empty rows dropped, like columns joined, sheetFormatPr's deepest level). */
function axisLines(doc, axis, format) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const touched = new Set();
  if (axis === 'row') {
    const sheetData = kid(root, 'sheetData', ns);
    pinRefs(sheetData, ns);
    forgetRows(doc);
    const byRow = new Map(kids(sheetData, 'row', ns).map((el) => [Number(attr(el, 'r')) - 1, el]));
    const el = (r, create) => {
      let x = byRow.get(r);
      if (!x && create) {
        x = make(doc, ns, 'row');
        x.setAttribute('r', String(r + 1));
        sheetData.insertBefore(x, kids(sheetData, 'row', ns).find((y) => Number(attr(y, 'r')) - 1 > r) || null);
        byRow.set(r, x);
      }
      return x;
    };
    return {
      get: (r) => { const x = byRow.get(r); return { level: levelOf(x), collapsed: on(attr(x, 'collapsed')), hidden: on(attr(x, 'hidden')) }; },
      set(r, patch) {
        const x = el(r, (patch.level > 0) || patch.collapsed || patch.hidden);
        if (!x) return;
        touched.add(x);
        if ('level' in patch) { if (patch.level > 0) x.setAttribute('outlineLevel', String(patch.level)); else x.removeAttribute('outlineLevel'); }
        if ('collapsed' in patch) { if (patch.collapsed) x.setAttribute('collapsed', '1'); else x.removeAttribute('collapsed'); }
        if ('hidden' in patch) { if (patch.hidden) x.setAttribute('hidden', '1'); else x.removeAttribute('hidden'); }
      },
      levels: () => [...byRow].map(([r, x]) => [r, Number(attr(x, 'outlineLevel') || 0)]).filter(([, l]) => l > 0),
      finish() {
        for (const x of touched) {
          reorder(x, ROW_ORDER);
          // (A row left with nothing of its own goes, as Excel leaves none.)
          if (!x.firstChild && [...x.attributes].every((a) => a.name === 'r' || a.name === 'spans')) { x.parentNode.removeChild(x); byRow.delete(Number(attr(x, 'r')) - 1); }
        }
        setSheetFormat(doc, { outlineLevelRow: this.levels().reduce((m, [, l]) => Math.max(m, l), 0) || null });
      },
    };
  }
  const colsEl = () => kid(root, 'cols', ns);
  const at = (c) => kids(colsEl(), 'col', ns).find((x) => c >= Number(attr(x, 'min')) - 1 && c <= Number(attr(x, 'max')) - 1);
  const defaultWidth = String(pxWidth(defaultColPx(format, format.mdw ?? 7), format.mdw ?? 7));
  return {
    get: (c) => { const x = at(c); return { level: levelOf(x), collapsed: on(attr(x, 'collapsed')), hidden: on(attr(x, 'hidden')) }; },
    set(c, patch) {
      if (!at(c) && !((patch.level > 0) || patch.collapsed || patch.hidden)) return;
      // (The column given an element of its own, split from its neighbours'; joined again in finish.)
      setColumns(doc, c, c, {}, format);
      const x = at(c);
      touched.add(c);
      if ('level' in patch) { if (patch.level > 0) x.setAttribute('outlineLevel', String(patch.level)); else x.removeAttribute('outlineLevel'); }
      if ('collapsed' in patch) { if (patch.collapsed) x.setAttribute('collapsed', '1'); else x.removeAttribute('collapsed'); }
      if (patch.hidden === true && !on(attr(x, 'hidden'))) {
        // (Hidden by the outline, a column without a width of its own is written width 0; one with keeps it.)
        if (!on(attr(x, 'customWidth'))) x.setAttribute('width', '0');
        x.setAttribute('hidden', '1');
        x.setAttribute('customWidth', '1');
      } else if (patch.hidden === false && on(attr(x, 'hidden'))) {
        x.removeAttribute('hidden');
        if (Number(attr(x, 'width') || 0) === 0) x.setAttribute('width', defaultWidth);
      }
    },
    levels() {
      const out = [];
      for (const x of kids(colsEl(), 'col', ns)) {
        const l = levelOf(x);
        // (Within the sheet's columns, whatever the file's min and max say.)
        if (l > 0) for (let c = Math.max(0, Number(attr(x, 'min')) - 1); c <= Math.min(MAX_COLS - 1, Number(attr(x, 'max')) - 1); c++) out.push([c, l]);
      }
      return out;
    },
    finish() {
      const parent = colsEl();
      if (parent) {
        const els = kids(parent, 'col', ns);
        for (const x of els) reorder(x, COL_ORDER);
        // (A column left with only the default width goes, as Excel leaves none: a summary column no longer collapsed.)
        const bare = (x) => [...x.attributes].every((p) => ['min', 'max', 'width'].includes(p.name)) && attr(x, 'width') === defaultWidth;
        for (const c of touched) { const x = at(c); if (x && bare(x)) parent.removeChild(x); }
        els.splice(0, els.length, ...kids(parent, 'col', ns));
        // (Neighbouring columns alike written as one element, as Excel writes them.)
        const same = (a, b) => Number(attr(a, 'max')) + 1 === Number(attr(b, 'min')) && [...a.attributes].length === [...b.attributes].length && [...a.attributes].every((p) => p.name === 'min' || p.name === 'max' || attr(b, p.name) === p.value);
        for (let k = els.length - 1; k > 0; k--) if (same(els[k - 1], els[k])) { els[k - 1].setAttribute('max', attr(els[k], 'max')); parent.removeChild(els[k]); }
        // (cols must hold a col: an empty one goes.)
        if (!els.length) root.removeChild(parent);
      }
      setSheetFormat(doc, { outlineLevelCol: this.levels().reduce((m, [, l]) => Math.max(m, l), 0) || null });
    },
  };
}

/** `edit(lines, pr)` on axis 'row' or 'col' of sheet si, then the sheet read again; what edit returns (false: nothing
 *  changed, the sheet left as it was). */
function outlineEdit(wb, si, axis, edit) {
  const part = wb._part(si);
  const lines = axisLines(part.doc, axis, part.model.format);
  const pr = readOutlinePr(part.doc.documentElement, part.doc.documentElement.namespaceURI);
  const done = edit(lines, pr);
  if (done === false) return false;
  lines.finish();
  wb._reread(si, part);
  wb.resetCalc?.();
  return true;
}
const after = (axis, pr) => (axis === 'row' ? pr.below : pr.right);

/** Group lines a to b one level deeper; false when one is at the deepest level already (Excel refuses). */
export function groupLines(wb, si, axis, a, b) {
  const done = outlineEdit(wb, si, axis, (lines) => {
    for (let i = a; i <= b; i++) if (lines.get(i).level >= MAX_LEVEL) return false;
    for (let i = a; i <= b; i++) lines.set(i, { level: lines.get(i).level + 1 });
    return true;
  });
  // (Automatic styles on: the groups' summary lines styled, "group with automatic styles".)
  if (done && wb.sheet(si).outline?.styles) styleSummaryLines(wb, si, axis, summariesOf(wb.sheet(si), axis, a, b));
  return done;
}

/** Ungroup lines a to b one level; false when none is grouped (Excel refuses). Hidden lines and the summary's
 *  collapsed stay, as Excel leaves them ("ungroup a collapsed group"). */
export function ungroupLines(wb, si, axis, a, b) {
  return outlineEdit(wb, si, axis, (lines) => {
    let any = false;
    for (let i = a; i <= b; i++) { const l = lines.get(i).level; if (l > 0) { lines.set(i, { level: l - 1 }); any = true; } }
    return any;
  });
}

/** The level button n clicked (Outline.ShowLevels): lines at level n or deeper hidden, the rest of the outline shown;
 *  a group of level n or deeper collapsed where its summary line stays in view. */
export function showLevels(wb, si, axis, n) {
  return outlineEdit(wb, si, axis, (lines, pr) => {
    const levels = lines.levels();
    if (!levels.length) return false;
    for (const [i, l] of levels) lines.set(i, { hidden: l >= n });
    for (const g of outlineGroups(levels, after(axis, pr))) {
      if (g.summary == null) continue;
      const shown = lines.get(g.summary).level < n;
      lines.set(g.summary, { collapsed: g.level >= n && shown });
    }
    return true;
  });
}

/** The group that line i's Show Detail or Hide Detail acts on: the one it sums up (the outermost, when several), else
 *  the innermost it is part of. */
function groupAt(groups, i) {
  const summed = groups.filter((g) => g.summary === i).sort((p, q) => p.level - q.level);
  if (summed.length) return summed[0];
  return groups.filter((g) => i >= g.start && i <= g.end).sort((p, q) => q.level - p.level)[0] || null;
}

/** Hide Detail (the − button, Range.ShowDetail = False) at line i: its group's lines hidden, its summary collapsed;
 *  false when i is in no group. */
export function hideDetail(wb, si, axis, i) {
  return outlineEdit(wb, si, axis, (lines, pr) => {
    const g = groupAt(outlineGroups(lines.levels(), after(axis, pr)), i);
    if (!g) return false;
    for (let k = g.start; k <= g.end; k++) lines.set(k, { hidden: true });
    if (g.summary != null) lines.set(g.summary, { collapsed: true });
    return true;
  });
}

/** Show Detail (the + button, Range.ShowDetail = True) at summary line i of a collapsed group: its lines shown but for
 *  those of groups inside it that stay collapsed ("hide inner then outer then show outer"); false otherwise (a hidden
 *  line's Show Detail changes nothing, as in Excel). */
export function showDetail(wb, si, axis, i) {
  return outlineEdit(wb, si, axis, (lines, pr) => {
    const groups = outlineGroups(lines.levels(), after(axis, pr));
    const g = groups.filter((x) => x.summary === i).sort((p, q) => p.level - q.level)
      .find((x) => lines.get(i).collapsed || lines.get(x.start).hidden);
    if (!g || lines.get(i).hidden) return false;
    const inner = groups.filter((x) => x.level > g.level && x.start >= g.start && x.end <= g.end && x.summary != null && lines.get(x.summary).collapsed);
    for (let k = g.start; k <= g.end; k++) lines.set(k, { hidden: inner.some((x) => k >= x.start && k <= x.end) });
    lines.set(i, { collapsed: false });
    return true;
  });
}

/** Clear Outline over lines a to b: their levels gone, those hidden by the outline shown, and the summaries of the groups
 *  cleared no longer collapsed ("clear outline": Excel's grouped columns shown again at the default width). */
export function clearOutline(wb, si, axis, a = 0, b = Infinity) {
  return outlineEdit(wb, si, axis, (lines, pr) => {
    const levels = lines.levels().filter(([i]) => i >= a && i <= b);
    if (!levels.length) return false;
    const groups = outlineGroups(lines.levels(), after(axis, pr)).filter((g) => g.end >= a && g.start <= b);
    for (const [i] of levels) lines.set(i, { level: 0, hidden: false, collapsed: false });
    for (const g of groups) if (g.summary != null) lines.set(g.summary, { collapsed: false });
    return true;
  });
}

/** The Settings: summary rows below their detail (`below`), columns right of theirs (`right`) and Automatic styles
 *  (`styles`), written in sheetPr's outlinePr as Excel writes it: summaryBelow="0", summaryRight="0", applyStyles="1",
 *  nothing for the defaults (setting them back takes outlinePr away, and sheetPr when it is left empty); `symbols`
 *  for Ctrl+8. */
export function setSummary(wb, si, { below, right, symbols, styles }) {
  const part = wb._part(si);
  const doc = part.doc;
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  let sheetPr = kid(root, 'sheetPr', ns);
  if (!sheetPr) { sheetPr = make(doc, ns, 'sheetPr'); root.insertBefore(sheetPr, [...root.childNodes].find((n) => n.nodeType === 1) || null); }
  let pr = kid(sheetPr, 'outlinePr', ns);
  if (!pr) { pr = make(doc, ns, 'outlinePr'); sheetPr.insertBefore(pr, [...sheetPr.childNodes].find((n) => n.nodeType === 1 && n.localName !== 'tabColor') || null); }
  for (const [name, v] of [['summaryBelow', below], ['summaryRight', right], ['showOutlineSymbols', symbols]]) {
    if (v == null) continue;
    if (v) pr.removeAttribute(name); else pr.setAttribute(name, '0');
  }
  // (Automatic styles: applyStyles="1", nothing when off, as Excel writes Outline.AutomaticStyles.)
  if (styles != null) { if (styles) pr.setAttribute('applyStyles', '1'); else pr.removeAttribute('applyStyles'); }
  reorder(pr, ['applyStyles', 'summaryBelow', 'summaryRight', 'showOutlineSymbols']);
  if (!pr.attributes.length) sheetPr.removeChild(pr);
  if (!sheetPr.attributes.length && ![...sheetPr.childNodes].some((n) => n.nodeType === 1)) root.removeChild(sheetPr);
  wb._reread(si, part);
  return true;
}

/** Ctrl+8 (the window's DisplayOutline): the outline symbols shown or hidden, written as Excel writes it, sheetView's
 *  showOutlineSymbols="0" and sheetPr's outlinePr showOutlineSymbols="0" when hidden (even on a sheet without an
 *  outline), neither when shown ("outline symbols off",
 *  "outline symbols off then on"). A view setting kept in the file, as gridlines are: not a step of the history. */
export function setOutlineSymbols(wb, si, on) {
  const part = wb._part(si);
  setShown(part.doc, 'showOutlineSymbols', on);
  return setSummary(wb, si, { symbols: on });
}

/** Columns inserted right of a hidden or collapsed column (rowcol.js moveCols): a copy of its element for columns at to
 *  at+n-1, shown (the default width when it had none) and not collapsed. */
export function insertedLines(el, at, n, defaultWidth) {
  const x = el.cloneNode(true);
  x.setAttribute('min', String(at + 1));
  x.setAttribute('max', String(at + n));
  x.removeAttribute('collapsed');
  if (on(attr(x, 'hidden'))) { x.removeAttribute('hidden'); if (Number(attr(x, 'width') || 0) === 0) x.setAttribute('width', String(defaultWidth)); }
  return x;
}

/** Lines at to at+n-1 just inserted (rowcol.js changeLines): the groups they are part of no longer collapsed, as Excel
 *  leaves them ("insert row in a collapsed group": the new row shown, row 6 without collapsed). */
export function openJoined(doc, axis, at, n, format) {
  const lines = axisLines(doc, axis, format);
  if (!lines.levels().length) return;
  const pr = readOutlinePr(doc.documentElement, doc.documentElement.namespaceURI);
  for (const g of outlineGroups(lines.levels(), after(axis, pr))) {
    if (g.end >= at && g.start <= at + n - 1 && g.summary != null && lines.get(g.summary).collapsed) lines.set(g.summary, { collapsed: false });
  }
  lines.finish();
}

/** Lines' levels set: `levels` [[i, level]] (Subtotal's outline, subtotal.js); the current ones as [[i, level]] with
 *  none given. */
export function lineLevels(wb, si, axis, levels = null) {
  if (!levels) { const part = wb._part(si); return axisLines(part.doc, axis, part.model.format).levels(); }
  return outlineEdit(wb, si, axis, (lines) => { for (const [i, level] of levels) lines.set(i, { level: Math.min(MAX_LEVEL, level) }); });
}

/** Auto Outline (Range.AutoOutline, "auto outline"): the sheet's outline made again from its formulas. A formula over a
 *  range of its own column ending on the row before it (starting on the row after, summaries above) groups those rows;
 *  one over its own row's cells left of it (right of it), those columns; a line in several such ranges goes deeper.
 *  False when no formula sums up lines, as Excel then has no outline to make. */
export function autoOutline(wb, si) {
  const sheet = wb.sheet(si);
  const pr = sheet.outline || { below: true, right: true };
  const spans = { row: [], col: [] };
  const col = (s) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  for (const [r, line] of filledRows(sheet.cells)) {
    for (const [c, cell] of line) {
      if (!cell.f) continue;
      const refs = [...cell.f.toUpperCase().matchAll(/\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)/g)].map((m) => ({ c0: col(m[1]), r0: Number(m[2]) - 1, c1: col(m[3]), r1: Number(m[4]) - 1 }));
      for (const x of refs) {
        if (x.c0 === c && x.c1 === c && (pr.below ? x.r1 === r - 1 : x.r0 === r + 1) && x.r1 > x.r0 - 1) spans.row.push([x.r0, x.r1]);
        if (x.r0 === r && x.r1 === r && (pr.right ? x.c1 === c - 1 : x.c0 === c + 1)) spans.col.push([x.c0, x.c1]);
      }
    }
  }
  if (!spans.row.length && !spans.col.length) return false;
  for (const axis of ['row', 'col']) {
    clearOutline(wb, si, axis);
    // (Each span counted once: several formulas over the same lines make one group.)
    const unique = [...new Set(spans[axis].map((s) => s.join(':')))].map((s) => s.split(':').map(Number));
    const depth = new Map();
    for (const [a, b] of unique) for (let i = a; i <= b; i++) depth.set(i, (depth.get(i) || 0) + 1);
    if (depth.size) lineLevels(wb, si, axis, [...depth]);
  }
  // (Automatic styles on, as the Settings box's Create: the used range's summary cells styled.)
  if (wb.sheet(si).outline?.styles) applyOutlineStyles(wb, si, { r0: 0, c0: 0, r1: wb.sheet(si).maxRow, c1: wb.sheet(si).maxCol });
  return true;
}

/** Sheet si's outline as the grid draws it: for rows and columns, their groups and the deepest level. */
export function outlineOf(sheet) {
  const rows = [...sheet.rows].filter(([, x]) => x.level > 0).map(([r, x]) => [r, x.level]);
  const cols = sheet.cols.flatMap((x) => (x.level > 0 ? Array.from({ length: x.max - x.min + 1 }, (_, k) => [x.min + k, x.level]) : []));
  const pr = sheet.outline || { below: true, right: true, symbols: true };
  const level = (list) => list.reduce((m, [, l]) => Math.max(m, l), 0);
  return { ...pr, rows: { level: level(rows), groups: outlineGroups(rows, pr.below) }, cols: { level: level(cols), groups: outlineGroups(cols, pr.right) } };
}
