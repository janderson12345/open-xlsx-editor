// Data validation rules and sparklines when rows, columns or cells are inserted or deleted, as Excel moves them
// (measured over COM: tools/excel-compare/excel-refmoves.ps1 on refmove-cases.mjs, test/refmoves-excel.json):
// - a rule's ranges move as ranges do (rowcol.js moveRanges, and here the Excel 2010 form's xm:sqref, which Excel
//   writes for a rule reading another sheet); its formulas stay relative to its first cell: when that cell is deleted
//   they read as they did from the first cell left (Excel kept D10>C10 on D10:D15 with rows 10:12 deleted). Cells
//   shifted through part of a rule split it: the part that moves is a rule of its own when its formulas have
//   relative references (D12's D12>C12 became C12's C12>#REF!), else it stays in the rule. Lines and cells inserted
//   take the rules of the line or cell before them;
// - a sparkline's cell (xm:sqref) moves with it, and its data (xm:f) as a reference does (sheetops.js mapFormulas
//   moves every xm:f); one is deleted with its cell or when all of its data is deleted; lines and cells inserted
//   take the sparklines of the line or cell before them (the row above, the column or cell to the left), their data
//   moved as a copy moves it (O2's B2:M2 gives the new O3 B3:M3), in the same group, after their source.
import { shiftFormula } from '../calc/tokens.js';
import { parseRange, rangeName, MAX_ROWS, MAX_COLS } from './refs.js';

const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XM = 'http://schemas.microsoft.com/office/excel/2006/main';
const ALL = (el, ns, name) => [...el.getElementsByTagNameNS(ns, name)];
const child = (el, ns, name) => [...el.childNodes].find((n) => n.nodeType === 1 && n.namespaceURI === ns && n.localName === name) || null;
const remove = (el) => el.parentNode && el.parentNode.removeChild(el);

/** The data validation rules of a sheet's XML, each with what holds its ranges and formulas: the main rules and the
 *  Excel 2010 ones in its extLst. */
function rulesOf(doc) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const main = ALL(root, ns, 'dataValidation').map((el) => ({
    el,
    get: () => el.getAttribute('sqref') || '',
    set: (text) => el.setAttribute('sqref', text),
    formulas: () => ['formula1', 'formula2'].map((n) => child(el, ns, n)).filter(Boolean),
  }));
  const ext = ALL(root, X14, 'dataValidation').map((el) => ({
    el,
    get: () => child(el, XM, 'sqref')?.textContent || '',
    set: (text) => { const s = child(el, XM, 'sqref'); if (s) s.textContent = text; },
    formulas: () => ['formula1', 'formula2'].map((n) => child(el, X14, n)).filter(Boolean).map((f) => child(f, XM, 'f')).filter(Boolean),
  }));
  return [...main, ...ext];
}

/** A range cut by a band of cells shifting ([b0, b1]: columns for axis 'row', rows for 'col'), from line `at` on:
 *  the pieces outside the band's lines and the one inside, or null when the shift doesn't cut it. */
function cut(range, axis, at, band) {
  const [p0, p1] = axis === 'row' ? [range.c0, range.c1] : [range.r0, range.r1];
  const reach = axis === 'row' ? range.r1 : range.c1;
  if (reach < at || p1 < band[0] || p0 > band[1] || (p0 >= band[0] && p1 <= band[1])) return null;
  const piece = (q0, q1) => (axis === 'row' ? { ...range, c0: q0, c1: q1 } : { ...range, r0: q0, r1: q1 });
  const out = [];
  if (p0 < band[0]) out.push(piece(p0, band[0] - 1));
  if (p1 > band[1]) out.push(piece(band[1] + 1, p1));
  return { outside: out, inside: piece(Math.max(p0, band[0]), Math.min(p1, band[1])) };
}

/** The first cell of a range (its top-left, or the first after the lines deleted) that a change keeps, or null. */
function firstKept(range, axis, at, n, move) {
  const tries = [{ r: range.r0, c: range.c0 }];
  if (n < 0) tries.push(axis === 'row' ? { r: at - n, c: range.c0 } : { r: range.r0, c: at - n });
  return tries.find((p) => p.r <= range.r1 && p.c <= range.c1 && move({ r0: p.r, c0: p.c, r1: p.r, c1: p.c })) || null;
}

/**
 * Before a change's formulas move (sheetops.js mapFormulas): each data validation rule of the sheet's XML split where
 * a shift of cells cuts it, and its formulas made to read from the cell that will be its first after the change.
 * `move(range)`: the range after the change, or null when it is deleted (rowcol.js shiftRange).
 */
export function prepareRules(doc, axis, at, n, band, move) {
  /** A rule's formulas, now relative to `from`, made to read from `to`. */
  const rebase = (rule, from, to) => {
    if (to.r === from.r && to.c === from.c) return;
    for (const f of rule.formulas()) f.textContent = shiftFormula(f.textContent, to.r - from.r, to.c - from.c);
  };
  /** Its formulas from the cell that will be its first: its own first cell, or the first one left after a delete. */
  const fromKept = (rule, ranges) => {
    const kept = ranges.map((r) => firstKept(r, axis, at, n, move)).find(Boolean);
    if (kept) rebase(rule, { r: ranges[0].r0, c: ranges[0].c0 }, kept);
  };
  for (const rule of rulesOf(doc)) {
    const ranges = rule.get().split(/\s+/).map(parseRange).filter(Boolean);
    if (!ranges.length) continue;
    const anchor = { r: ranges[0].r0, c: ranges[0].c0 };
    const relative = rule.formulas().some((f) => shiftFormula(f.textContent, 1, 1) !== f.textContent);
    const kept = []; // (the ranges, or the pieces of them, that stay in the rule)
    const moving = []; // (pieces a shift of cells moves, cut from them)
    let cutAway = false;
    for (const r of ranges) {
      const pieces = band && cut(r, axis, at, band);
      if (!pieces) { kept.push(r); continue; }
      const moves = !!move(pieces.inside);
      // (A piece the shift deletes leaves the rule whole when its formulas are absolute, as a reference is left; with
      // relative ones it goes: C12:E12 deleted shifting left kept C10:C15's whole number rule and E10:E15's list
      // whole, and left D10:D15's custom rule D10:D11 D13:D15.)
      if (!moves && !relative) { kept.push(r); continue; }
      kept.push(...pieces.outside);
      if (moves) moving.push(pieces.inside); else cutAway = true;
    }
    let own = kept;
    if (moving.length && kept.length && relative) {
      // (A rule of its own for what moves, its formulas as they read at its first cell, after the others; without
      // the original's revision id, which is the original's.)
      const copy = rule.el.cloneNode(true);
      copy.removeAttributeNS('http://schemas.microsoft.com/office/spreadsheetml/2014/revision', 'uid');
      rule.el.parentNode.appendChild(copy);
      const piece = rulesOf(doc).find((x) => x.el === copy);
      piece.set(moving.map(rangeName).join(' '));
      rebase(piece, anchor, { r: moving[0].r0, c: moving[0].c0 });
      fromKept(piece, moving);
    } else own = [...kept, ...moving];
    if (moving.length || cutAway) {
      rule.set(own.map(rangeName).join(' '));
      if (own.length) rebase(rule, anchor, { r: own[0].r0, c: own[0].c0 });
      const wrap = rule.el.parentNode;
      if (wrap.getAttribute('count') != null) wrap.setAttribute('count', String([...wrap.childNodes].filter((x) => x.nodeType === 1 && x.localName === 'dataValidation').length));
    }
    if (own.length) fromKept(rule, own);
  }
}

/**
 * After the change: the lines or cells inserted (n > 0 at `at`) taken into the rules of the line or cell before them,
 * as Excel gives them that line's validation with its format: a range ending there grows over them (B10:B15 is
 * B10:C15 after a column is inserted at C), the part of one in the band's cells gets them as a range of its own (D10:D15
 * gets E12:E13 when E12:E13 is inserted shifting right). Both the main rules and the Excel 2010 ones.
 */
export function extendRules(doc, axis, at, n, band) {
  if (n <= 0 || at <= 0) return;
  for (const rule of rulesOf(doc)) {
    const ranges = rule.get().split(/\s+/).map(parseRange).filter(Boolean);
    const added = [];
    let changed = false;
    for (const r of ranges) {
      const [l1, p0, p1] = axis === 'row' ? [r.r1, r.c0, r.c1] : [r.c1, r.r0, r.r1];
      if (l1 !== at - 1) continue;
      const q0 = band ? Math.max(p0, band[0]) : p0;
      const q1 = band ? Math.min(p1, band[1]) : p1;
      if (q0 > q1) continue;
      changed = true;
      const end = Math.min(at + n - 1, (axis === 'row' ? MAX_ROWS : MAX_COLS) - 1);
      if (q0 === p0 && q1 === p1) { if (axis === 'row') r.r1 = end; else r.c1 = end; continue; }
      added.push(axis === 'row' ? { r0: at, r1: end, c0: q0, c1: q1 } : { c0: at, c1: end, r0: q0, r1: q1 });
    }
    if (changed) rule.set([...ranges, ...added].map(rangeName).join(' '));
  }
}

/** An extLst's ext left empty by what was taken out of it goes too. */
function dropEmpty(el, name) {
  if ([...el.childNodes].some((x) => x.nodeType === 1)) {
    if (el.getAttribute('count') != null) el.setAttribute('count', String([...el.childNodes].filter((x) => x.nodeType === 1 && x.localName === name).length));
    return;
  }
  const ext = el.parentNode;
  remove(el);
  if (ext && ext.localName === 'ext' && ![...ext.childNodes].some((x) => x.nodeType === 1)) {
    const list = ext.parentNode;
    remove(ext);
    if (list && list.localName === 'extLst' && ![...list.childNodes].some((x) => x.nodeType === 1)) remove(list);
  }
}

/** The sheet's sparklines: [{ el, f, sqref }] in their groups' order. */
const sparklinesOf = (doc) => ALL(doc.documentElement, X14, 'sparkline').map((el) => ({ el, f: child(el, XM, 'f'), sqref: child(el, XM, 'sqref') }));

/**
 * After the change: the ranges in the sheet's extLst moved (`shift(sqref)` gives a list of ranges after it, '' when
 * none is left): the Excel 2010 data validation rules' and the sparklines' cells; the sparklines of the line or cell
 * before lines or cells inserted copied into them.
 */
export function moveExtRanges(doc, axis, at, n, band, shift) {
  for (const el of ALL(doc.documentElement, X14, 'dataValidation')) {
    const s = child(el, XM, 'sqref');
    const next = s && shift(s.textContent);
    if (!s) continue;
    if (next) s.textContent = next;
    else { const wrap = el.parentNode; remove(el); dropEmpty(wrap, 'dataValidation'); }
  }
  for (const sp of sparklinesOf(doc)) {
    const next = sp.sqref && shift(sp.sqref.textContent);
    if (next) sp.sqref.textContent = next; else dropSparkline(sp.el);
  }
  if (n > 0 && at > 0) copySparklines(doc, axis, at, n, band);
}

/** A sparkline taken out, its group when it was the last, and the groups' ext when that was the last group. */
function dropSparkline(el) {
  const list = el.parentNode;
  remove(el);
  if (list && ![...list.childNodes].some((x) => x.nodeType === 1)) {
    const group = list.parentNode;
    const groups = group?.parentNode;
    remove(group);
    if (groups) dropEmpty(groups, 'sparklineGroup');
  }
}

/** The sparklines in the line before an insert (in the band's cells, for a shift) copied into the new lines. */
function copySparklines(doc, axis, at, n, band) {
  for (const sp of sparklinesOf(doc)) {
    const cell = sp.sqref && parseRange(sp.sqref.textContent);
    if (!cell || !sp.f) continue;
    const [line, across] = axis === 'row' ? [cell.r0, cell.c0] : [cell.c0, cell.r0];
    if (line !== at - 1 || (band && (across < band[0] || across > band[1]))) continue;
    let after = sp.el;
    for (let k = 1; k <= n; k++) {
      const [dr, dc] = axis === 'row' ? [k, 0] : [0, k];
      const copy = sp.el.cloneNode(true);
      child(copy, XM, 'f').textContent = shiftFormula(sp.f.textContent, dr, dc);
      child(copy, XM, 'sqref').textContent = rangeName({ r0: cell.r0 + dr, c0: cell.c0 + dc, r1: cell.r0 + dr, c1: cell.c0 + dc });
      after.parentNode.insertBefore(copy, after.nextSibling);
      after = copy;
    }
  }
}

/** Sparklines whose data was all deleted (their xm:f now #REF!) taken out, as Excel deletes them. True when any was. */
export function dropDeadSparklines(doc) {
  let any = false;
  for (const sp of sparklinesOf(doc)) if (sp.f && sp.f.textContent.includes('#REF!')) { dropSparkline(sp.el); any = true; }
  return any;
}
