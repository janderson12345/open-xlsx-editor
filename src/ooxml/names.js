// Defined names made, changed and deleted, as Excel's Name Manager and Define Name do (Formulas > Defined Names), and
// as the Name Box names a selection. Which names Excel takes, and how it saves them, were measured
// (tools/excel-compare/excel-names.ps1, test/names-excel.json): a letter, _ or \ first, then letters, digits, _ . ? and \ (\ alone, but not \ and more);
// not a cell reference (A1 to XFD1048576) nor anything R1C1 reads (R, C, R2, C3, RC1, R1C1), nor TRUE or FALSE; 255
// characters at most; one name per scope (the workbook, or a sheet). In the file they are sorted, capitals aside, a
// built-in name (_xlnm.Print_Area) by its own name, and a sheet's name before the workbook's of the same name.
import { kids, kid } from './xml.js';
import { colIndex, MAX_ROWS, MAX_COLS } from './refs.js';
import { withEl } from './sheetops.js';
import { tokenize, moveRefs } from '../calc/tokens.js';
import { normalizeFormula } from '../calc/fileform.js';
import { parseFormula } from '../calc/parse.js';
import { bookIndexes, bookNames } from './extlinks.js';

/** Excel's words (its dialogs', which COM doesn't reach: not measured). */
export const NAME_MESSAGES = {
  syntax: "The syntax of this name isn't correct.\n\nVerify that the name:\n- Starts with a letter or underscore (_)\n- Doesn't include a space or character that isn't allowed\n- Doesn't conflict with an existing name in the workbook",
  exists: 'The name entered already exists. Enter a unique name.',
  formula: "There's a problem with this formula.",
};

/** Whether Excel takes `name` for a defined name (the measured rules above). */
export function validName(name) {
  const n = String(name ?? '');
  if (!n || n.length > 255) return false;
  if (n[0] === '\\') return n.length === 1;
  if (!/^[A-Za-z_\u00C0-\uFFFF][A-Za-z0-9_.?\\\u00C0-\uFFFF]*$/.test(n)) return false;
  if (/^(TRUE|FALSE)$/i.test(n)) return false;
  // (Excel's own names, which it keeps for itself: Print_Area measured, Print_Titles its pair.)
  if (/^(Print_Area|Print_Titles)$/i.test(n)) return false;
  // (A cell's address on the sheet, A1 to XFD1048576.)
  const a1 = /^([A-Za-z]{1,3})(\d+)$/.exec(n);
  if (a1 && colIndex(a1[1]) < MAX_COLS && Number(a1[2]) >= 1 && Number(a1[2]) <= MAX_ROWS) return false;
  // (R1C1's: R, C, R2, C3, RC, RC1, R1C, R1C1.)
  if (/^(R\d*|C\d*|R\d*C\d*)$/i.test(n)) return false;
  return true;
}

/** The order Excel saves names in: by name, capitals aside, a built-in one by its own name; a sheet's first. */
const sortKey = (d) => d.name.replace(/^_xlnm\./i, '').toLowerCase();
const before = (a, b) => (sortKey(a) < sortKey(b) ? -1 : sortKey(a) > sortKey(b) ? 1 : (a.sheet == null ? 1 : 0) - (b.sheet == null ? 1 : 0));

/** The workbook's <definedNames>, made in its place if it has none. */
function container(wb) {
  const root = wb.doc.documentElement;
  let names = kid(root, 'definedNames', wb.ns);
  if (!names) {
    names = wb.doc.createElementNS(wb.ns, 'definedNames');
    const AFTER = ['calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];
    root.insertBefore(names, [...root.childNodes].find((n) => n.nodeType === 1 && AFTER.includes(n.localName)) || null);
  }
  return names;
}

/** What a name refers to, typed with or without "=": its file form (another workbook's [Book.xlsx]Sheet1!$A$1 as its
 *  link's [1]Sheet1!$A$1, the link made if there is none: extlinks.js), or throws Excel's words. */
function refersText(wb, refersTo) {
  const typed = String(refersTo ?? '').trim().replace(/^=/, '');
  if (!typed) throw new Error(NAME_MESSAGES.formula);
  try {
    const text = normalizeFormula(typed);
    // (Several ranges apart by commas, =Sheet1!$A$2:$B$3,Sheet1!$D$6:$E$7: a union, written as Excel writes it.)
    try { parseFormula(text); } catch { parseFormula(`(${text})`); }
    return bookIndexes(wb, text);
  } catch {
    throw new Error(NAME_MESSAGES.formula);
  }
}

/** The name `name` of scope `sheet` (an index, or null for the workbook), or null. */
export function findName(wb, name, sheet = null) {
  const n = String(name).toLowerCase();
  return wb.definedNames.find((d) => d.name.toLowerCase() === n && (d.sheet ?? null) === (sheet ?? null)) || null;
}

/** A new name: { name, refersTo (typed, "=Sheet1!$A$1"), sheet (null: the workbook), comment, hidden }. Throws Excel's
 *  words for a name it wouldn't take, or one there is in that scope. */
export function defineName(wb, { name, refersTo, sheet = null, comment = '', hidden = false }) {
  if (!validName(name)) throw new Error(NAME_MESSAGES.syntax);
  if (findName(wb, name, sheet)) throw new Error(NAME_MESSAGES.exists);
  const text = refersText(wb, refersTo);
  const el = wb.doc.createElementNS(wb.ns, 'definedName');
  el.setAttribute('name', name);
  if (comment) el.setAttribute('comment', comment);
  if (sheet != null) el.setAttribute('localSheetId', String(sheet));
  if (hidden) el.setAttribute('hidden', '1');
  el.textContent = text;
  const d = withEl({ name, sheet, text }, el);
  // (In Excel's order: before the first name that sorts after it.)
  const names = container(wb);
  const next = wb.definedNames.filter((x) => x.el.parentNode === names).sort(before).find((x) => before(d, x) < 0);
  names.insertBefore(el, next ? next.el : null);
  wb.definedNames.push(d);
  wb.bookDirty = true;
  namesChanged(wb, [name]);
  return d;
}

/** A name changed: { name, refersTo, comment } (its scope stays, as Excel's Edit Name keeps it). A new name is checked
 *  as a new one, and formulas that used the old one use the new. */
export function editName(wb, { name, sheet = null }, change) {
  const d = findName(wb, name, sheet);
  if (!d) throw new Error(`no name ${name}`);
  const next = change.name ?? d.name;
  if (next.toLowerCase() !== d.name.toLowerCase()) {
    if (!validName(next)) throw new Error(NAME_MESSAGES.syntax);
    if (findName(wb, next, sheet)) throw new Error(NAME_MESSAGES.exists);
  }
  const text = change.refersTo != null ? refersText(wb, change.refersTo) : d.text;
  if (change.comment != null) { if (change.comment) d.el.setAttribute('comment', change.comment); else d.el.removeAttribute('comment'); }
  d.text = text;
  d.el.textContent = text;
  if (next !== d.name) {
    const old = d.name;
    d.name = next;
    d.el.setAttribute('name', next);
    // (Kept in Excel's order.)
    const names = d.el.parentNode;
    names.removeChild(d.el);
    const after = wb.definedNames.filter((x) => x !== d && x.el.parentNode === names).sort(before).find((x) => before(d, x) < 0);
    names.insertBefore(d.el, after ? after.el : null);
    renameInFormulas(wb, old, next, sheet);
  }
  wb.bookDirty = true;
  namesChanged(wb, [d.name]);
  return d;
}

/** A name deleted: formulas using it show #NAME?, as in Excel. */
export function deleteName(wb, { name, sheet = null }) {
  const d = findName(wb, name, sheet);
  if (!d) throw new Error(`no name ${name}`);
  const names = d.el.parentNode;
  names.removeChild(d.el);
  if (!kids(names, 'definedName', wb.ns).length) names.parentNode.removeChild(names);
  wb.definedNames.splice(wb.definedNames.indexOf(d), 1);
  wb.bookDirty = true;
  namesChanged(wb, [d.name]);
}

/** Formulas that use `from` (of scope `sheet`: on that sheet only) using `to`, as Excel's rename does. */
function renameInFormulas(wb, from, to, sheet) {
  const lower = from.toLowerCase();
  wb.sheets.forEach((info, si) => {
    if (info.kind !== 'worksheet' || !info.path || (sheet != null && si !== sheet)) return;
    // (A sheet with a name of its own the same: its formulas read that one, and stay.)
    if (sheet == null && findName(wb, from, si)) return;
    const changes = [];
    for (const [r, line] of wb.sheet(si).cells) {
      for (const [c, cell] of line) {
        if (cell.f == null || cell.f === '' || !cell.f.toLowerCase().includes(lower)) continue;
        let toks;
        try { toks = tokenize(cell.f); } catch { continue; }
        const hits = toks.filter((t) => t.type === 'name' && t.text.toLowerCase() === lower);
        if (!hits.length) continue;
        let text = '';
        let at = 0;
        for (const t of hits) { text += cell.f.slice(at, t.start) + to; at = t.end; }
        changes.push({ r, c, formula: text + cell.f.slice(at), ...(cell.dynamic ? {} : { legacy: true }) });
      }
    }
    if (changes.length) wb.apply(si, changes);
  });
}

/** After names change, the formulas that use them worked out again (a new name stops being #NAME?). */
function namesChanged(wb, names) {
  wb.resetCalc();
  const calc = wb.calc;
  calc.build();
  const lower = names.map((n) => n.toLowerCase());
  const also = [];
  for (const e of calc.entries.values()) {
    const text = String(e.text || '').toLowerCase();
    if (!lower.some((n) => text.includes(n))) continue;
    let toks;
    try { toks = tokenize(e.text); } catch { continue; }
    if (toks.some((t) => t.type === 'name' && lower.includes(t.text.toLowerCase()))) also.push({ si: e.si, r: e.r, c: e.c });
  }
  if (also.length) calc.recalc([], also);
}

/** The names a person sees (not Excel's hidden or built-in ones), with what each refers to, its scope's name and its
 *  value as the Name Manager lists them: [{ name, sheet, scope, refersTo, comment }]. */
export function listNames(wb) {
  return wb.definedNames.filter((d) => !d.name.startsWith('_xlnm.') && d.el?.getAttribute('hidden') !== '1').sort(before).map((d) => ({
    name: d.name, sheet: d.sheet ?? null, scope: d.sheet == null ? 'Workbook' : wb.sheets[d.sheet]?.name ?? '', refersTo: `=${bookNames(wb, d.text)}`,
    comment: d.el?.getAttribute('comment') || '',
  }));
}

/** The names sheet si sees, as Excel's Paste Names (F3), Use in Formula and Paste List list them (measured over COM:
 *  ListNames): the workbook's and the sheet's own (one of its own hiding the workbook's of the same name), not hidden
 *  ones, a built-in one by its own name (Print_Area); sorted by name, capitals aside: [{ name, text }]. */
export function namesOnSheet(wb, si) {
  const seen = wb.definedNames.filter((d) => d.el?.getAttribute('hidden') !== '1' && (d.sheet == null || d.sheet === si));
  const own = (d) => d.sheet == null && seen.some((x) => x.sheet === si && x.name.toLowerCase() === d.name.toLowerCase());
  const key = (n) => n.toLowerCase();
  return seen.filter((d) => !own(d)).map((d) => ({ name: d.name.replace(/^_xlnm\./i, ''), text: d.text }))
    .sort((a, b) => (key(a.name) < key(b.name) ? -1 : key(a.name) > key(b.name) ? 1 : 0));
}

/** The defined names after cells were cut and pasted (tokens.js moveRefs: `from` the cells cut, `to` where they went),
 *  as Excel's follow (measured over COM, B2:C3 cut to E5): a name wholly in the cut cells moves with them ($B$2:$C$3:
 *  $E$5:$F$6; to another sheet, Data!$E$5:$F$6), one only partly in them as a formula's range does ($B$2:$B$5 cut
 *  to another sheet: $B$4:$B$5), a single cell pasted over is Sheet1!#REF!; a name with relative references (B2), a
 *  constant and names apart from the cut stay. Returns how many changed. */
export function moveNamesForCut(wb, from, to) {
  let n = 0;
  for (const d of [...wb.definedNames]) {
    let toks;
    try { toks = tokenize(d.text); } catch { continue; }
    const refs = toks.filter((t) => t.type === 'ref' && t.area);
    const abs = (a) => a.kind === 'cols' ? a.ac0 && a.ac1 : a.kind === 'rows' ? a.ar0 && a.ar1 : a.ar0 && a.ac0 && (a.kind === 'cell' || (a.ar1 && a.ac1));
    if (!refs.length || !refs.every((t) => abs(t.area))) continue;
    const next = moveRefs(d.text, { here: from.sheet, from, to });
    if (next === d.text) continue;
    editName(wb, { name: d.name, sheet: d.sheet ?? null }, { refersTo: `=${next}` });
    n++;
  }
  return n;
}
