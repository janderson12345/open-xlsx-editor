// What the editors on one page share, as Excel's open workbooks share one clipboard: the last cells copied in any of
// them (so another can paste them, or Paste Link to them: an external reference, extlinks.js), and the workbooks
// open, by file name (so a link to one of them is updated from it, as Excel updates a link to an open workbook, and
// Home > Cell Styles > Merge Styles can take the styles of one).

let last = null;
const open = new Set();

/** The last copy (clipops.js snapshot, its `text` what went on the system clipboard) made in an editor here. */
export const lastCopy = () => last;
/** Keep `clip` as the last copy (null: let `clip` go, if it is still the last). */
export function setLastCopy(clip, drop = null) {
  if (drop) { if (last === drop) last = null; return; }
  last = clip;
}

/** An editor's workbook offered to the others while it is mounted: `entry` { workbook(), fileName() }; returns the
 *  call that takes it back. */
export function offerBook(entry) {
  open.add(entry);
  return () => open.delete(entry);
}

/** The workbook open here under file name `name` (capitals aside, any folder ignored), or null. */
export function openBook(name) {
  const want = String(name || '').replace(/^.*[\\/]/, '').toLowerCase();
  if (!want) return null;
  for (const e of open) {
    const n = String(e.fileName() || '').replace(/^.*[\\/]/, '').toLowerCase();
    if (n === want) return e.workbook();
  }
  return null;
}

/** The workbooks open here: [{ name, workbook }] (name: the file's, any folder left out; null for one not named). */
export function openBooks() {
  return [...open].map((e) => ({ name: e.fileName() ? String(e.fileName()).replace(/^.*[\\/]/, '') : null, workbook: e.workbook() })).filter((b) => b.workbook);
}
