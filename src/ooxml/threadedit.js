// Threaded comments written as Excel 365 writes them (Review > New Comment, a reply, an edit; measured over COM:
// tools/excel-compare/excel-threads.ps1, test/threads-excel.json): each comment a threadedComment in the sheet's
// threadedComments part (a reply's parentId its thread's first), its author a person in the workbook's persons part,
// and beside it, for older versions of Excel, a placeholder note in the comments part ("[Threaded comment]..." with
// every comment and reply, its author tc={id}) with a hidden box in the VML drawing, as a note has. The parts, their
// relationships and content types are made when the workbook or sheet has none. Deleting a thread is noteedit.js's.
import { NS } from './xml.js';
import { refName } from './refs.js';
import { notesPartsFor } from './noteedit.js';

const THREADED = 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment';
const PERSON = 'http://schemas.microsoft.com/office/2017/10/relationships/person';
const TC_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';
const CT_THREADS = 'application/vnd.ms-excel.threadedcomments+xml';
const CT_PERSONS = 'application/vnd.ms-excel.person+xml';
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const XR = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"';
/** The placeholder's opening words, as Excel writes them. */
export const PLACEHOLDER = '[Threaded comment]\r\n\r\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\r\n\r\n';

/** Line breaks as Excel keeps them in comments: CR LF. */
const crlf = (s) => String(s).replace(/\r?\n/g, '\r\n');
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unesc = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
const hex = (n) => Array.from({ length: n }, () => Math.floor(Math.random() * 16).toString(16).toUpperCase()).join('');
/** A new id, as Excel writes them: {8-4-4-4-12} in capitals. */
export const newGuid = () => `{${hex(8)}-${hex(4)}-4${hex(3)}-${'89AB'[Math.floor(Math.random() * 4)]}${hex(3)}-${hex(12)}}`;
/** A comment's time as Excel writes it: local, to hundredths of a second. */
export const commentTime = (d = new Date()) => {
  const p = (n, w = 2) => String(n).padStart(w, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}.${p(Math.floor(d.getMilliseconds() / 10))}`;
};
const relOf = (pkg, path, type) => pkg.readRels(path).find((r) => r.type === type);
const freePath = (pkg, base, ext) => { for (let n = 1; ; n++) if (!pkg.has(`${base}${n}.${ext}`)) return `${base}${n}.${ext}`; };

/** The person `who` (a name, or { name, userId, providerId } from the host's directory) writes or (`mentioned`) is
 *  mentioned as: the workbook's own with that name, else one added. An author is userId the name and provider None, as
 *  Excel names someone not signed in. Someone mentioned is a person of their own, provider PeoplePicker, even when an
 *  author has the same name: Excel, reading a mention made here, gave it such a person and kept the author apart
 *  (excel-mentions-check.ps1, test/mentions-excel.json). Returns their id. */
function personId(wb, who, mentioned = false) {
  const given = typeof who === 'string' ? { name: who } : who;
  const name = given.name;
  const userId = given.userId || name;
  // (Excel gave the mentioned person PeoplePicker whatever provider the file had for them: AD became PeoplePicker.)
  const providerId = mentioned ? 'PeoplePicker' : given.providerId || 'None';
  const pkg = wb.pkg;
  let path = pkg.rels.find((r) => r.type === PERSON)?.target;
  if (!path) {
    path = 'xl/persons/person.xml';
    pkg.setText(path, `${DECL}<personList xmlns="${TC_NS}" xmlns:x="${NS.x}"></personList>`);
    pkg.addOverride(path, CT_PERSONS);
    pkg.addRel(PERSON, 'persons/person.xml');
  }
  const text = pkg.text(path);
  const picked = (el) => /\sproviderId="PeoplePicker"/.test(el);
  const found = [...text.matchAll(/<person\b[^>]*>/g)].find((m) => unesc(/displayName="([^"]*)"/.exec(m[0])?.[1] || '') === name && picked(m[0]) === (providerId === 'PeoplePicker'));
  if (found) return /\sid="([^"]+)"/.exec(found[0])[1];
  const id = newGuid();
  const person = `<person displayName="${esc(name)}" id="${id}" userId="${esc(userId)}" providerId="${esc(providerId)}"/>`;
  // (In Excel's order: by name, then provider, None before PeoplePicker, as Excel saved them back.)
  const key = (n, p) => `${n.toLowerCase()}\u0000${p}`;
  const after = [...text.matchAll(/<person\b[^>]*>/g)].find((m) => key(unesc(/displayName="([^"]*)"/.exec(m[0])?.[1] || ''), /providerId="([^"]*)"/.exec(m[0])?.[1] || '') > key(name, providerId));
  if (after) pkg.setText(path, text.slice(0, after.index) + person + text.slice(after.index));
  else pkg.setText(path, /<personList\b[^>]*\/>/.test(text) ? text.replace(/<personList\b([^>]*)\/>/, `<personList$1>${person}</personList>`) : text.replace('</personList>', `${person}</personList>`));
  return id;
}

/** The workbook's people (its persons part): [{ id, name, userId, providerId }], for @mentions. */
export function workbookPeople(wb) {
  const path = wb.pkg.rels.find((r) => r.type === PERSON)?.target;
  const text = path && wb.pkg.has(path) ? wb.pkg.text(path) : '';
  const at = (el, k) => { const m = new RegExp(`\\s${k}="([^"]*)"`).exec(el); return m ? unesc(m[1]) : ''; };
  return [...text.matchAll(/<person\b[^>]*>/g)].map((m) => ({ id: at(m[0], 'id'), name: at(m[0], 'displayName'), userId: at(m[0], 'userId'), providerId: at(m[0], 'providerId') }));
}

/**
 * A comment's text and its mentions as a threadedComment's children: `<text>`, then `<mentions>` with a `mention`
 * for each "@Name" of the people mentioned ([{ name, userId, providerId }]) found in the text, in the order they come:
 * the person's id (added to the persons part when new), a new id of its own, where it starts in the text and its
 * length with the "@" ([MS-XLSX] CT_Mention, as Excel 365 writes a mention picked from its list).
 */
function textAndMentions(wb, text, mentions = []) {
  const body = crlf(text);
  const found = [];
  for (const who of mentions) {
    const tag = `@${who.name}`;
    for (let i = body.indexOf(tag); i >= 0; i = body.indexOf(tag, i + tag.length)) {
      // (Not inside a longer name: "@Ann" in "@Anna".)
      if (/[\p{L}\p{N}]/u.test(body[i + tag.length] || '')) continue;
      if (!found.some((f) => i < f.start + f.length && f.start < i + tag.length)) found.push({ who, start: i, length: tag.length });
    }
  }
  found.sort((a, b) => a.start - b.start);
  const list = found.map((f) => `<mention mentionpersonId="${personId(wb, f.who, true)}" mentionId="${newGuid()}" startIndex="${f.start}" length="${f.length}"/>`).join('');
  return `<text>${esc(body).replace(/&quot;/g, '"')}</text>${list ? `<mentions>${list}</mentions>` : ''}`;
}

/** Sheet si's threadedComments part, made (empty) with its relationship when it has none. */
function threadsPart(wb, si) {
  const pkg = wb.pkg;
  const sheet = wb.sheets[si].path;
  const have = relOf(pkg, sheet, THREADED)?.target;
  if (have) return have;
  const path = freePath(pkg, 'xl/threadedComments/threadedComment', 'xml');
  pkg.setText(path, `${DECL}<ThreadedComments xmlns="${TC_NS}" xmlns:x="${NS.x}"></ThreadedComments>`);
  pkg.addOverride(path, CT_THREADS);
  pkg.addRel(THREADED, `../${path.slice(3)}`, false, sheet);
  return path;
}

/** The thread on cell `ref` of a threadedComments part: [{ id, el (its text), text }], first then replies. */
function threadAt(xml, ref) {
  const all = [...xml.matchAll(/<threadedComment\b[^>]*>[\s\S]*?<\/threadedComment>/g)].map((m) => ({
    el: m[0], ref: /\sref="([^"]+)"/.exec(m[0])?.[1], id: /\sid="([^"]+)"/.exec(m[0])?.[1], parent: /\sparentId="([^"]+)"/.exec(m[0])?.[1] || null,
    text: unesc(/<text>([\s\S]*?)<\/text>/.exec(m[0])?.[1] || ''),
  }));
  const top = all.find((x) => x.ref === ref && !x.parent);
  return top ? [top, ...all.filter((x) => x.parent === top.id)] : [];
}

/** The placeholder note of a thread for older versions of Excel: every comment and reply, indented. */
function placeholderText(items) {
  return PLACEHOLDER + items.map((x, i) => `${i ? 'Reply' : 'Comment'}:\r\n    ${crlf(x.text)}`).join('\r\n');
}

/** The thread's placeholder in the comments part, and its box in the VML drawing (as a new thread's: Excel's place for
 *  it, a column right, three wide, four tall). */
function writePlaceholder(wb, si, r, c, top, items) {
  const pkg = wb.pkg;
  const { comments, vml } = notesPartsFor(wb, si);
  const ref = refName(r, c);
  let xml = pkg.text(comments);
  if (!/xmlns:xr=/.test(xml)) xml = xml.replace(/<comments\b([^>]*?)>/, `<comments$1 ${XR}>`);
  const author = `tc=${top.id}`;
  const authors = [...xml.matchAll(/<author>([\s\S]*?)<\/author>/g)].map((m) => m[1]);
  let authorId = authors.indexOf(author);
  if (authorId < 0) {
    authorId = authors.length;
    xml = /<authors\s*\/>/.test(xml) ? xml.replace(/<authors\s*\/>/, `<authors><author>${author}</author></authors>`) : xml.replace('</authors>', `<author>${author}</author></authors>`);
  }
  const comment = `<comment ref="${ref}" authorId="${authorId}" shapeId="0" xr:uid="${top.id}"><text><t>${esc(placeholderText(items)).replace(/&quot;/g, '"')}</t></text></comment>`;
  const existing = new RegExp(`<comment\\b[^>]*\\sref="${ref}"[^>]*>[\\s\\S]*?</comment>`).exec(xml);
  if (existing) xml = xml.replace(existing[0], () => comment);
  else xml = /<commentList\s*\/>/.test(xml) ? xml.replace(/<commentList\s*\/>/, `<commentList>${comment}</commentList>`) : xml.replace('</commentList>', `${comment}</commentList>`);
  pkg.setText(comments, xml);
  let shapes = pkg.text(vml);
  const has = [...shapes.matchAll(/<v:shape\b[\s\S]*?<\/v:shape>/g)].some((m) => /ObjectType="Note"/.test(m[0]) && new RegExp(`<x:Row>${r}</x:Row>`).test(m[0]) && new RegExp(`<x:Column>${c}</x:Column>`).test(m[0]));
  if (!has) {
    const block = Number(/<o:idmap\b[^>]*\sdata="(\d+)/.exec(shapes)?.[1] || 1);
    const ids = [...shapes.matchAll(/<v:shape\b[^>]*\sid="_x0000_s(\d+)"/g)].map((m) => Number(m[1]));
    const id = Math.max(block * 1024, ...ids) + 1;
    const top0 = Math.max(0, r - 1);
    // (Excel's box for B2 at 107.25pt, 7.5pt; for F6 at 299.25pt, 67.5pt: 48pt a column, 15pt a row of the defaults.)
    const shape = `<v:shape id="_x0000_s${id}" type="#_x0000_t202" style='position:absolute;margin-left:${59.25 + c * 48}pt;margin-top:${Math.max(0, r * 15 - 7.5)}pt;width:108pt;height:59.25pt;z-index:${ids.length + 1};visibility:hidden' fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto">` +
      `<v:fill color2="infoBackground [80]"/><v:shadow color="none [81]" obscured="t"/><v:path o:connecttype="none"/><v:textbox style='mso-direction-alt:auto'><div style='text-align:left'></div></v:textbox>` +
      `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${c + 1}, 15, ${top0}, 10, ${c + 3}, 31, ${top0 + 4}, 9</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${r}</x:Row><x:Column>${c}</x:Column></x:ClientData></v:shape>`;
    shapes = shapes.replace(/<\/xml>\s*$/, `${shape}</xml>`);
    pkg.setText(vml, shapes);
  }
}

/** A new thread on cell (r, c) of sheet si by `author` (Review > New Comment). Refused when the cell has a note or a
 *  thread already: Excel's words. Returns its id. */
export function addThread(wb, si, r, c, text, author = 'Author', { when = new Date(), mentions = [] } = {}) {
  const pkg = wb.pkg;
  const ref = refName(r, c);
  const had = relOf(pkg, wb.sheets[si].path, THREADED)?.target;
  if (had && threadAt(pkg.text(had), ref).length) throw new Error('This cell has a comment already: reply to it.');
  const notes = relOf(pkg, wb.sheets[si].path, `${NS.r}/comments`)?.target;
  if (notes && new RegExp(`<comment\\b[^>]*\\sref="${ref}"`).test(pkg.text(notes) || '')) throw new Error('This cell has a note already: delete it first.');
  const path = threadsPart(wb, si);
  let xml = pkg.text(path);
  const id = newGuid();
  const el = `<threadedComment ref="${ref}" dT="${commentTime(when)}" personId="${personId(wb, author)}" id="${id}">${textAndMentions(wb, text, mentions)}</threadedComment>`;
  xml = /<ThreadedComments\b[^>]*\/>/.test(xml) ? xml.replace(/<ThreadedComments\b([^>]*)\/>/, `<ThreadedComments$1>${el}</ThreadedComments>`) : xml.replace('</ThreadedComments>', `${el}</ThreadedComments>`);
  pkg.setText(path, xml);
  writePlaceholder(wb, si, r, c, { id }, [{ text }]);
  wb.edited = true;
  return id;
}

/** A reply to the thread on cell (r, c) of sheet si. */
export function replyThread(wb, si, r, c, text, author = 'Author', { when = new Date(), mentions = [] } = {}) {
  const pkg = wb.pkg;
  const ref = refName(r, c);
  const path = relOf(pkg, wb.sheets[si].path, THREADED)?.target;
  const items = path ? threadAt(pkg.text(path), ref) : [];
  if (!items.length) throw new Error('No comment on this cell to reply to.');
  // (A resolved thread takes no reply until it is reopened: Excel's AddReply on one left the file without it.)
  if (/\sdone="1"/.test(items[0].el.slice(0, items[0].el.indexOf('>')))) return null;
  // (Excel orders replies of the same time by id, and numbers each reply after the first from the one before:
  // {EFE96F7A-5C52-4835-...} then {EFE96F7A-5C52-4836-...}.)
  const prev = items.length > 1 ? items.at(-1).id : null;
  const bump = prev && /^\{([0-9A-F]{8}-[0-9A-F]{4})-([0-9A-F]{4})-([0-9A-F]{4}-[0-9A-F]{12})\}$/.exec(prev);
  const id = bump && bump[2] !== 'FFFF' ? `{${bump[1]}-${(parseInt(bump[2], 16) + 1).toString(16).toUpperCase().padStart(4, '0')}-${bump[3]}}` : newGuid();
  const el = `<threadedComment ref="${ref}" dT="${commentTime(when)}" personId="${personId(wb, author)}" id="${id}" parentId="${items[0].id}">${textAndMentions(wb, text, mentions)}</threadedComment>`;
  const last = items.at(-1).el;
  pkg.setText(path, pkg.text(path).replace(last, () => last + el));
  writePlaceholder(wb, si, r, c, items[0], [...items, { text }]);
  wb.edited = true;
  return id;
}

/** Comment k (0 the first, then the replies) of the thread on cell (r, c) given new text, or (text null) a reply
 *  deleted. The placeholder written again. */
export function editThread(wb, si, r, c, k, text, { mentions = [] } = {}) {
  const pkg = wb.pkg;
  const ref = refName(r, c);
  const path = relOf(pkg, wb.sheets[si].path, THREADED)?.target;
  const items = path ? threadAt(pkg.text(path), ref) : [];
  const item = items[k];
  if (!item) throw new Error('No such comment.');
  if (text == null && k === 0) throw new Error('Delete the whole thread instead.');
  const next = text == null ? '' : item.el.replace(/<text>[\s\S]*?<\/text>(?:<mentions>[\s\S]*?<\/mentions>|<mentions\/>)?/, () => textAndMentions(wb, text, mentions));
  pkg.setText(path, pkg.text(path).replace(item.el, () => next));
  const left = text == null ? items.filter((_, i) => i !== k) : items.map((x, i) => (i === k ? { ...x, text } : x));
  writePlaceholder(wb, si, r, c, items[0], left);
  wb.edited = true;
}

/** The thread on cell (r, c) of sheet si resolved (`on`) or reopened: done="1" on its first comment, after its id, as
 *  Excel writes it (measured over COM: CommentThreaded.Resolved). */
export function resolveThread(wb, si, r, c, on = true) {
  const pkg = wb.pkg;
  const ref = refName(r, c);
  const path = relOf(pkg, wb.sheets[si].path, THREADED)?.target;
  const items = path ? threadAt(pkg.text(path), ref) : [];
  if (!items.length) throw new Error('No comment on this cell.');
  const top = items[0].el;
  const start = top.slice(0, top.indexOf('>')).replace(/\sdone="[^"]*"/, '');
  const next = (on ? `${start} done="1"` : start) + top.slice(top.indexOf('>'));
  pkg.setText(path, pkg.text(path).replace(top, () => next));
  wb.edited = true;
}
