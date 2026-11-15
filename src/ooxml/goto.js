// What the Name Box goes to when a reference or a name is typed into it, as Excel's Name Box takes them (Excel's
// documented behaviour: COM doesn't reach the Name Box): a cell or a range ("B5", "b5:d10", "$B$5"), whole columns or
// rows ("B:D", "5:5"), any of these on another sheet ("Sheet2!A1", "'My data'!B:B"), or a defined name (the sheet's
// own first, then the workbook's) that stands for a range or for ranges of one sheet, or a table: its name for its data
// rows, or a table reference ("Fruit[Qty]", "Fruit[#All]"). And the other way round, the name a selection has, which
// the Name Box shows instead of the cell (selectionName, measured over COM).
import { parseRange, MAX_ROWS, MAX_COLS, colIndex } from './refs.js';
import { readTables, tableArea } from './tables.js';
import { sheetPrefix } from '../calc/tokens.js';

export const GOTO_MESSAGES = {
  invalid: 'You must enter a valid reference you want to go to, or type a valid name for the selection.',
  defining: "Naming the selection from the Name Box isn't available yet",
};

/** "B:D" or "5:7" (with or without $) as a range of whole columns or rows; null otherwise. */
function wholeLines(text) {
  let m = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/.exec(text);
  if (m) {
    const [a, b] = [colIndex(m[1]), colIndex(m[2])];
    if (a < MAX_COLS && b < MAX_COLS) return { r0: 0, r1: MAX_ROWS - 1, c0: Math.min(a, b), c1: Math.max(a, b) };
  }
  m = /^\$?(\d{1,7}):\$?(\d{1,7})$/.exec(text);
  if (m) {
    const [a, b] = [Number(m[1]) - 1, Number(m[2]) - 1];
    if (a >= 0 && b >= 0 && a < MAX_ROWS && b < MAX_ROWS) return { r0: Math.min(a, b), r1: Math.max(a, b), c0: 0, c1: MAX_COLS - 1 };
  }
  return null;
}

/** A name as Excel allows one: a letter, _ or \ first, then letters, digits, _, . and \ (not a cell reference). */
const NAME = /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/;

/**
 * Where `text` typed into the Name Box goes from sheet si: { si, sel: { r0, c0, r1, c1, ar, ac } } (the active cell
 * the range's first), or { error } with Excel's words, or { error, name: true } for a name that isn't there yet (Excel
 * would name the selection with it).
 */
export function goToTarget(wb, si, text, depth = 0) {
  const t = String(text ?? '').trim();
  const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(t);
  let sheet = si;
  let ref = t;
  if (m) {
    sheet = wb.sheetIndexOf((m[1] ?? m[2]).replace(/''/g, "'"));
    // (A hidden sheet can't be gone to, as Excel can't select on one.)
    if (sheet < 0 || !wb.canShow(sheet)) return { error: GOTO_MESSAGES.invalid };
    ref = m[3];
  }
  // (A table, by its name (its data rows) or a table reference.)
  const tm = !m && /^([A-Za-z_\\][A-Za-z0-9_.\\]*)(?:\[(.*)\])?$/.exec(t);
  if (tm) {
    const tables = readTables(wb);
    if (tables.some((x) => x.name.toLowerCase() === tm[1].toLowerCase())) {
      const a = tableArea(tables, tm[1], tm[2] ?? '#Data', { si, r: -1, c: -1 });
      if (a.error) return { error: GOTO_MESSAGES.invalid };
      if (!wb.canShow(a.si)) return { error: GOTO_MESSAGES.invalid };
      return { si: a.si, sel: { r0: a.r0, c0: a.c0, r1: a.r1, c1: a.c1, ar: a.r0, ac: a.c0 } };
    }
  }
  // (parseRange reads a first and a second corner only: a third part is no reference here.)
  const area = ref.split(':').length <= 2 && (parseRange(ref) || wholeLines(ref));
  if (area) return { si: sheet, sel: { ...area, ar: area.r0, ac: area.c0 } };
  if (!m && depth === 0 && NAME.test(t)) {
    const named = wb.definedName(t, si);
    if (!named) return { error: GOTO_MESSAGES.defining, name: true };
    // (A name standing for one range, its text a reference like Sheet1!$B$2:$B$9, or for ranges of one sheet.)
    return nameTarget(wb, si, named.text);
  }
  return { error: GOTO_MESSAGES.invalid };
}

/** What Excel's Go To box lists: the defined names that stand for ranges (not its own hidden ones, nor those of
 *  another sheet) and the tables, in order. */
export function goToNames(wb, si) {
  const names = wb.definedNames.filter((d) => !d.name.startsWith('_xlnm.') && (d.sheet == null || d.sheet === si) && !nameTarget(wb, si, d.text).error).map((d) => d.name);
  return [...new Set([...names, ...readTables(wb).map((t) => t.name)])].sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' }));
}

/** A name's text split into its references at the commas between them (not those inside a quoted sheet name). */
function nameParts(text) {
  const parts = [];
  let from = 0;
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    if (text[i] === "'") quoted = !quoted;
    else if (text[i] === ',' && !quoted) { parts.push(text.slice(from, i)); from = i + 1; }
  }
  return [...parts, text.slice(from)];
}

/** Where a name goes: its one range, or its ranges of one sheet, as Application.Goto went (measured over COM, from
 *  Sheet2, test/namebox-excel.json): =Sheet1!$C$3:$D$4,Sheet1!$A$1 selected C3:D4 and A1 in that order, A1 (the last
 *  range's first cell) active; =Sheet1!$B$2:$C$4 selected it, B2 active. */
function nameTarget(wb, si, text) {
  const gone = nameParts(text).map((part) => goToTarget(wb, si, part, 1));
  if (gone.some((t) => t.error || t.si !== gone[0].si)) return { error: GOTO_MESSAGES.invalid };
  if (gone.length === 1) return gone[0];
  const areas = gone.map(({ sel: { r0, c0, r1, c1 } }) => ({ r0, c0, r1, c1 }));
  const last = areas[areas.length - 1];
  return { si: gone[0].si, sel: { ...last, ar: last.r0, ac: last.c0, more: areas.slice(0, -1) } };
}

// (A reference all of whose ends are absolute, after a sheet's name: Sheet1!$B$2:$C$4, Sheet1!$B:$B, Sheet1!$3:$5.)
const ABSOLUTE = /^(?:'((?:[^']|'')+)'|([^'!]+))!(\$[A-Za-z]{1,3}\$\d{1,7}(?::\$[A-Za-z]{1,3}\$\d{1,7})?|\$[A-Za-z]{1,3}:\$[A-Za-z]{1,3}|\$\d{1,7}:\$\d{1,7})$/;

/** A name's ranges, when its text is absolute references only: [{ si, r0, c0, r1, c1 }], or null. */
function absoluteAreas(wb, text) {
  const out = [];
  for (const part of nameParts(text.trim())) {
    const m = ABSOLUTE.exec(part.trim());
    if (!m) return null;
    const area = parseRange(m[3].replace(/\$/g, '')) || wholeLines(m[3]);
    if (!area) return null;
    out.push({ si: wb.sheetIndexOf((m[1] ?? m[2]).replace(/''/g, "'")), ...area });
  }
  return out;
}

const nameKey = (d) => d.name.replace(/^_xlnm\./i, '').toLowerCase();

/**
 * The name the selection's ranges (`areas`, in Excel's Areas order) on sheet si have, as Excel reports it for the
 * selection (Selection.Name over COM, 35 cases in test/namebox-excel.json; COM can't read the Name Box itself), or
 * null. A name whose ranges are exactly these, in this order (B2:C3 or B2:C5 under a name of B2:C4 have none; A1 then
 * C3:D4 has the name of =Sheet1!$A$1,Sheet1!$C$3:$D$4, C3:D4 then A1 hasn't), all ends absolute (a relative =Sheet1!B2
 * and an =OFFSET(...) gave none). Hidden names, print areas and other sheets' own names count. Of more than one, the
 * first by name, capitals aside (Alpha2 before beta; Aw before a sheet's Zown; Ahid, hidden, before Seen); the same
 * name, the sheet's own (Sheet1's before Sheet2's) before the workbook's. A table isn't a name: its rows gave none.
 * The answer: the name, a print area's as Print_Area, another sheet's own after its sheet ("Sheet2!Theirs", as COM
 * gives it; COM gives the shown sheet's own as "Sheet1!Own" too, shown here without it, as the Name Box has no room for
 * the sheet it is on: not measured).
 */
export function selectionName(wb, si, areas) {
  const same = (d) => {
    const got = absoluteAreas(wb, d.text);
    return got && got.length === areas.length && got.every((a, k) => a.si === si && a.r0 === areas[k].r0 && a.c0 === areas[k].c0 &&
      a.r1 === areas[k].r1 && a.c1 === areas[k].c1);
  };
  const scope = (d) => (d.sheet == null ? Infinity : d.sheet);
  const found = wb.definedNames.filter(same).sort((a, b) => (nameKey(a) < nameKey(b) ? -1 : nameKey(a) > nameKey(b) ? 1 : scope(a) - scope(b)))[0];
  if (!found) return null;
  const name = found.name.replace(/^_xlnm\./i, '');
  return found.sheet == null || found.sheet === si ? name : `${sheetPrefix(wb.sheets[found.sheet].name)}${name}`;
}
