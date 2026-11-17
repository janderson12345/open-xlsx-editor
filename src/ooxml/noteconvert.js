// Review > Notes > Convert to Comments, as Excel 365's (measured over COM, Workbook.ConvertComments:
// tools/excel-compare/excel-notes-visible.ps1, test/notes-visible-excel.json). Every note in the workbook, on every
// sheet, becomes a threaded comment and the note goes:
// - the comment's text is the note's, its first line dropped when it is the note's author's name and a colon (bold or
//   not, as Excel's New Note starts it; another name, or a bold "Important:", stays), its line breaks CR LF;
// - its author is a person by the note's author's name (userId the name, providerId None), one for each name, made
//   when the workbook has no such person (one of the same name with another userId isn't used); a note with no author
//   is written with the empty id {00000000-...} and no person (Excel then reads it as "Author");
// - it has no date (no dT), as Excel writes it;
// - the threads and their placeholders for older versions of Excel ("[Threaded comment]...", author tc={id}) are in
//   the order of their cells, row by row, the threads already there kept; the comments part's authors are those its
//   comments use, in their order;
// - each note's box in the VML drawing goes and its thread's placeholder box is added after the shapes there, at
//   Excel's place for a new thread's (a shown note's comes hidden), its id after the highest the drawing had, and
//   every shape's z-index numbered again in order.
// A protected sheet's notes stay (Excel's ConvertComments left them, the other sheets converted).
import { readNotes } from './notes.js';
import { refName, parseRef } from './refs.js';
import { notesPartsFor, noteShapeCell } from './noteedit.js';
import { newGuid, crlf, placeholderText, placeholderShape, threadsPart, personsPart } from './threadedit.js';

const XR = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"';
/** The person id Excel gives a note with no author. */
export const NO_PERSON = '{00000000-0000-0000-0000-000000000000}';
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unesc = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
/** Text as Excel writes it in a threadedComment or a placeholder: escaped, its quotes left as they are. */
const textXml = (s) => esc(s).replace(/&quot;/g, '"');

/** A note's text as its comment's: without the first line when that is its author's name and a colon. */
export function noteCommentText(note) {
  const text = note.text || '';
  const head = `${note.author}:`;
  if (!note.author || !text.startsWith(head)) return text;
  const rest = text.slice(head.length);
  return rest === '' ? '' : /^\r?\n/.test(rest) ? rest.replace(/^\r?\n/, '') : text;
}

/** Cell `ref`'s place, row by row: a number to sort by. */
const order = (ref) => { const at = parseRef(ref); return at ? at.r * 16384 + at.c : Infinity; };

/** `el` put into `xml` before the first of `items` (matches of the elements in order, with their ref) on a later cell,
 *  else before `close` (or into a root written empty, `<name .../>`). */
function insertByCell(xml, items, el, ref, close, rootName) {
  const at = order(ref);
  const next = items.find((m) => order(m.ref) > at);
  if (next) return xml.slice(0, next.index) + el + xml.slice(next.index);
  const empty = new RegExp(`<${rootName}\\b([^>]*)/>`);
  if (!xml.includes(close) && empty.test(xml)) return xml.replace(empty, (m, a) => `<${rootName}${a}>${el}${close}`);
  const i = xml.lastIndexOf(close);
  return xml.slice(0, i) + el + xml.slice(i);
}

/** The person whose id a converted note's comment takes, by its author's name: the workbook's (or this conversion's,
 *  `people.ids`) with that name as its userId and provider None, else a new one, put in `people.fresh` too. */
function personFor(wb, name, people) {
  if (!name) return NO_PERSON;
  if (people.ids.has(name)) return people.ids.get(name);
  const text = wb.pkg.rels.some((r) => /\/person$/.test(r.type)) ? wb.pkg.text(personsPart(wb)) : '';
  const at = (el, k) => { const m = new RegExp(`\\s${k}="([^"]*)"`).exec(el); return m ? unesc(m[1]) : ''; };
  const found = [...text.matchAll(/<person\b[^>]*>/g)].find((m) => at(m[0], 'displayName') === name && at(m[0], 'userId') === name && at(m[0], 'providerId') === 'None');
  const id = found ? at(found[0], 'id') : newGuid();
  people.ids.set(name, id);
  if (!found) people.fresh.push([name, id]);
  return id;
}

/**
 * Convert to Comments: every note in the workbook made a threaded comment, as Excel's. `skip(si)` true for a sheet to
 * leave (a protected one). Returns how many notes were converted.
 */
export function convertNotes(wb, { skip = () => false } = {}) {
  const pkg = wb.pkg;
  const people = { ids: new Map(), fresh: [] };
  let count = 0;
  wb.sheets.forEach((info, si) => {
    if (skip(si)) return;
    const notes = [...readNotes(wb, si)].filter(([, n]) => n.kind === 'note').map(([key, n]) => {
      const [r, c] = key.split(',').map(Number);
      return { r, c, ref: refName(r, c), note: n };
    }).sort((a, b) => a.r - b.r || a.c - b.c);
    if (!notes.length) return;
    const made = notes.map((x) => ({ ...x, id: newGuid(), text: noteCommentText(x.note), person: personFor(wb, x.note.author, people) }));

    // (The threads, each before the first thread on a later cell.)
    const threads = threadsPart(wb, si);
    let tx = pkg.text(threads);
    for (const m of made) {
      const tops = [...tx.matchAll(/<threadedComment\b[^>]*>/g)].filter((t) => !/\sparentId="/.test(t[0])).map((t) => ({ index: t.index, ref: /\sref="([^"]+)"/.exec(t[0])?.[1] }));
      // (The person's id may be the workbook's own, read from its persons part: escaped again, as any text from a file.)
      const el = `<threadedComment ref="${m.ref}" personId="${esc(m.person)}" id="${m.id}"><text>${textXml(crlf(m.text))}</text></threadedComment>`;
      tx = insertByCell(tx, tops, el, m.ref, '</ThreadedComments>', 'ThreadedComments');
    }
    pkg.setText(threads, tx);

    // (The comments part: the notes out, the placeholders in by cell, its authors those used, in order.)
    const { comments, vml } = notesPartsFor(wb, si);
    let cx = pkg.text(comments);
    if (!/xmlns:xr=/.test(cx)) cx = cx.replace(/<comments\b([^>]*?)>/, `<comments$1 ${XR}>`);
    for (const m of made) cx = cx.replace(new RegExp(`<comment\\b[^>]*\\sref="${m.ref}"[^>]*>[\\s\\S]*?</comment>`), '');
    for (const m of made) {
      const list = [...cx.matchAll(/<comment\b[^>]*>/g)].map((t) => ({ index: t.index, ref: /\sref="([^"]+)"/.exec(t[0])?.[1] }));
      const el = `<comment ref="${m.ref}" authorId="tc=${m.id}" shapeId="0" xr:uid="${m.id}"><text><t>${textXml(placeholderText([{ text: m.text }]))}</t></text></comment>`;
      cx = insertByCell(cx, list, el, m.ref, '</commentList>', 'commentList');
    }
    pkg.setText(comments, renumberAuthors(cx));

    // (The VML drawing: the notes' boxes out, their threads' in after the rest, ids on from the highest.)
    let vx = pkg.text(vml);
    const block = Number(/<o:idmap\b[^>]*\sdata="(\d+)/.exec(vx)?.[1] || 1);
    let id = Math.max(block * 1024, ...[...vx.matchAll(/<v:shape\b[^>]*\sid="_x0000_s(\d+)"/g)].map((t) => Number(t[1])));
    const gone = new Set(made.map((m) => `${m.r},${m.c}`));
    vx = vx.replace(/<v:shape\b[\s\S]*?<\/v:shape>/g, (s) => (gone.has(noteShapeCell(s)) ? '' : s));
    const boxes = made.map((m) => placeholderShape(++id, 0, m.r, m.c)).join('');
    vx = vx.replace(/<\/xml>\s*$/, () => `${boxes}</xml>`);
    let z = 0;
    vx = vx.replace(/<v:shape\b[^>]*>/g, (s) => s.replace(/z-index:\s*-?\d+/, () => `z-index:${++z}`));
    pkg.setText(vml, vx);
    count += made.length;
  });
  // (New people first in the persons part, in the order first used, as Excel put them.)
  if (people.fresh.length) {
    const path = personsPart(wb);
    const list = people.fresh.map(([name, id]) => `<person displayName="${esc(name)}" id="${id}" userId="${esc(name)}" providerId="None"/>`).join('');
    const text = pkg.text(path);
    pkg.setText(path, /<personList\b[^>]*\/>/.test(text) ? text.replace(/<personList\b([^>]*)\/>/, (m, a) => `<personList${a}>${list}</personList>`) : text.replace(/<personList\b[^>]*>/, (m) => m + list));
  }
  if (count) wb.edited = true;
  return count;
}

/** A comments part's authors those its comments use, in the order they first do, and each comment's authorId
 *  numbered so (an authorId "tc={id}" is a new placeholder's author by name). */
function renumberAuthors(xml) {
  const authors = [...xml.matchAll(/<author>([\s\S]*?)<\/author>/g)].map((m) => m[1]);
  const used = [];
  const out = xml.replace(/(<comment\b[^>]*\sauthorId=")([^"]*)"/g, (m, a, id) => {
    const name = /^\d+$/.test(id) ? authors[Number(id)] ?? '' : id;
    let k = used.indexOf(name);
    if (k < 0) k = used.push(name) - 1;
    return `${a}${k}"`;
  });
  const list = `<authors>${used.map((a) => `<author>${a}</author>`).join('')}</authors>`;
  return /<authors\s*\/>/.test(out) ? out.replace(/<authors\s*\/>/, () => list) : out.replace(/<authors>[\s\S]*?<\/authors>/, () => list);
}
