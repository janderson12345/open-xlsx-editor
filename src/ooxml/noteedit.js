// Notes written as Excel's New Note writes them (Review > Notes, Shift+F2, a cell's menu): a comment in the sheet's
// comments part, its text the author's name in bold ("Avery Chen:") and the note's own text after a line break, both
// in Tahoma 9; its box a shape in the sheet's VML drawing, hidden until the pointer rests on the cell; the parts,
// relationships, content types and the sheet's legacyDrawing made when the sheet has none. Deleting a note removes its
// comment and box (and a threaded comment's thread with it). Show/Hide Note and Show All Notes write a note's shape
// shown or hidden as Excel saves it (setNoteVisible, setAllNotesVisible).
import { NS } from './xml.js';
import { refName } from './refs.js';
import { readNotes } from './notes.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const THREADED = 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment';
const CT_COMMENTS = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
const CT_VML = 'application/vnd.openxmlformats-officedocument.vmlDrawing';
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const VML_NS = 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"';
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const run = (t, bold) => `<r><rPr>${bold ? '<b/>' : ''}<sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t xml:space="preserve">${esc(t)}</t></r>`;
/** The sheet elements that come after legacyDrawing (CT_Worksheet's order). */
const AFTER_LEGACY = ['legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

const relOf = (wb, path, type) => wb.pkg.readRels(path).find((r) => r.type === type);

/** A new part's path, `base` followed by the first number not taken: xl/comments2.xml. */
function freePath(pkg, base, ext) {
  for (let n = 1; ; n++) if (!pkg.has(`${base}${n}.${ext}`)) return `${base}${n}.${ext}`;
}

/** The comments part and VML drawing of sheet si, made (empty) when it has none: { comments, vml }. */
function notesParts(wb, si) {
  const pkg = wb.pkg;
  const info = wb.sheets[si];
  let comments = relOf(wb, info.path, `${REL}/comments`)?.target;
  if (!comments) {
    comments = freePath(pkg, 'xl/comments', 'xml');
    pkg.setText(comments, `${DECL}<comments xmlns="${NS.x}"><authors></authors><commentList></commentList></comments>`);
    pkg.addOverride(comments, CT_COMMENTS);
    pkg.addRel(`${REL}/comments`, `../${comments.slice(3)}`, false, info.path);
  }
  let vml = relOf(wb, info.path, `${REL}/vmlDrawing`)?.target;
  if (!vml) {
    vml = freePath(pkg, 'xl/drawings/vmlDrawing', 'vml');
    // (Each drawing's shape ids from its own block of 1,024: the first block no other VML drawing uses.)
    const used = new Set(Object.keys(pkg.files).filter((k) => /\.vml$/i.test(k)).flatMap((k) => [...(pkg.text(k) || '').matchAll(/<o:idmap\b[^>]*\sdata="([^"]+)"/g)].flatMap((m) => m[1].split(',').map(Number))));
    let block = 1;
    while (used.has(block)) block++;
    pkg.setText(vml, `<xml ${VML_NS}><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${block}"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype></xml>`);
    const ct = pkg.text('[Content_Types].xml');
    if (!/<Default\b[^>]*Extension="vml"/i.test(ct)) pkg.setText('[Content_Types].xml', ct.replace(/<Default\b/, `<Default Extension="vml" ContentType="${CT_VML}"/><Default`));
    const id = pkg.addRel(`${REL}/vmlDrawing`, `../${vml.slice(3)}`, false, info.path);
    // The sheet's legacyDrawing, in its place among the sheet's elements.
    const { doc } = wb._part(si);
    const root = doc.documentElement;
    const el = doc.createElementNS(root.namespaceURI, root.prefix ? `${root.prefix}:legacyDrawing` : 'legacyDrawing');
    el.setAttributeNS(REL, 'r:id', id);
    const before = [...root.childNodes].find((n) => n.nodeType === 1 && AFTER_LEGACY.includes(n.localName));
    root.insertBefore(el, before || null);
    wb.dirty.add(info.path);
  }
  return { comments, vml };
}

/**
 * Give cell (r, c) of sheet si a note with `text` (a new one under `author`, or the one there with its text replaced,
 * keeping its author's name), as Excel's New Note and Edit Note do.
 */
export function setNote(wb, si, r, c, text, author = 'Author') {
  const pkg = wb.pkg;
  const { comments, vml } = notesParts(wb, si);
  const ref = refName(r, c);
  let xml = pkg.text(comments);
  const existing = new RegExp(`<comment\\b[^>]*\\sref="${ref}"[^>]*>[\\s\\S]*?</comment>`).exec(xml);
  // (The note's author: the one it has, else a new author in the list.)
  const authors = [...xml.matchAll(/<author>([\s\S]*?)<\/author>/g)].map((m) => m[1]);
  let authorId = existing ? Number(/\sauthorId="(\d+)"/.exec(existing[0])?.[1] || 0) : authors.findIndex((a) => a === esc(author));
  if (authorId < 0) {
    authorId = authors.length;
    xml = /<authors\s*\/>/.test(xml) ? xml.replace(/<authors\s*\/>/, `<authors><author>${esc(author)}</author></authors>`) : xml.replace('</authors>', `<author>${esc(author)}</author></authors>`);
  }
  const name = existing ? (/<r>(?:(?!<\/r>)[\s\S])*?<b\/>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>/.exec(existing[0])?.[1] ?? null) : `${esc(author)}:`;
  const decode = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
  const body = `<text>${name != null ? run(decode(name), true) + run(`\n${text}`, false) : run(text, false)}</text>`;
  const comment = `<comment ref="${ref}" authorId="${authorId}" shapeId="0">${body}</comment>`;
  if (existing) xml = xml.replace(existing[0], comment);
  else xml = /<commentList\s*\/>/.test(xml) ? xml.replace(/<commentList\s*\/>/, `<commentList>${comment}</commentList>`) : xml.replace('</commentList>', `${comment}</commentList>`);
  pkg.setText(comments, xml);
  // Its box, when the cell has none: beside the cell, as Excel's new note's (a column right, three wide, four tall).
  let shapes = pkg.text(vml);
  const has = [...shapes.matchAll(/<v:shape\b[\s\S]*?<\/v:shape>/g)].some((m) => /ObjectType="Note"/.test(m[0]) && new RegExp(`<x:Row>${r}</x:Row>`).test(m[0]) && new RegExp(`<x:Column>${c}</x:Column>`).test(m[0]));
  if (!has) {
    const block = Number(/<o:idmap\b[^>]*\sdata="(\d+)/.exec(shapes)?.[1] || 1);
    const ids = [...shapes.matchAll(/<v:shape\b[^>]*\sid="_x0000_s(\d+)"/g)].map((m) => Number(m[1]));
    const id = Math.max(block * 1024, ...ids) + 1;
    const z = ids.length + 1;
    const shape = `<v:shape id="_x0000_s${id}" type="#_x0000_t202" style='position:absolute;margin-left:${60 + c * 48}pt;margin-top:${Math.max(0, r - 1) * 15 + 1.5}pt;width:108pt;height:59.25pt;z-index:${z};visibility:hidden' fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto">` +
      `<v:fill color2="infoBackground [80]"/><v:shadow color="none [81]" obscured="t"/><v:path o:connecttype="none"/><v:textbox style='mso-direction-alt:auto'><div style='text-align:left'></div></v:textbox>` +
      `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${c + 1}, 15, ${Math.max(0, r - 1)}, 10, ${c + 3}, 31, ${Math.max(0, r - 1) + 4}, 9</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${r}</x:Row><x:Column>${c}</x:Column></x:ClientData></v:shape>`;
    shapes = shapes.replace(/<\/xml>\s*$/, `${shape}</xml>`);
    pkg.setText(vml, shapes);
  }
  wb.edited = true;
}

/** Delete cell (r, c)'s note (or threaded comment, with its replies) on sheet si; false when it has none. */
export function deleteNote(wb, si, r, c) {
  const pkg = wb.pkg;
  const path = wb.sheets[si].path;
  const ref = refName(r, c);
  let found = false;
  const drop = (target, re) => {
    if (!target) return;
    const text = pkg.text(target);
    const next = text.replace(re, () => { found = true; return ''; });
    if (next !== text) pkg.setText(target, next);
  };
  const comments = relOf(wb, path, `${REL}/comments`)?.target;
  drop(comments, new RegExp(`<comment\\b[^>]*\\sref="${ref}"[^>]*>[\\s\\S]*?</comment>`, 'g'));
  // (A deleted thread's placeholder author, tc={id}, goes from the list too and the rest are numbered again, as
  // Excel's saved comments part has it: excel-threads.ps1.)
  if (comments && found) pkg.setText(comments, pruneThreadAuthors(pkg.text(comments)));
  drop(relOf(wb, path, THREADED)?.target, new RegExp(`<threadedComment\\b[^>]*\\sref="${ref}"[^>]*?(?:/>|>[\\s\\S]*?</threadedComment>)`, 'g'));
  const vml = relOf(wb, path, `${REL}/vmlDrawing`)?.target;
  if (vml) {
    const text = pkg.text(vml);
    const next = text.replace(/<v:shape\b[\s\S]*?<\/v:shape>/g, (m) => (/ObjectType="Note"/.test(m) && new RegExp(`<x:Row>${r}</x:Row>`).test(m) && new RegExp(`<x:Column>${c}</x:Column>`).test(m) ? '' : m));
    if (next !== text) pkg.setText(vml, next);
  }
  if (found) wb.edited = true;
  return found;
}


/** Whether a VML shape is the box of a note (or thread placeholder) on cell (r, c). */
export const noteShapeAt = (shape, r, c) => /ObjectType="Note"/.test(shape) && new RegExp(`<x:Row>\\s*${r}\\s*</x:Row>`).test(shape) && new RegExp(`<x:Column>\\s*${c}\\s*</x:Column>`).test(shape);

/** A note's VML shape shown always (`on`) or hidden until the pointer rests on its cell, as Excel writes it (measured
 *  over COM, Comment.Visible: excel-notes-visible.ps1, test/notes-visible-excel.json): shown, `visibility:visible` in
 *  its style and `<x:Visible/>` after its x:Column, on that line's indent; hidden, `visibility:hidden` and no
 *  x:Visible, which is the shape Excel had before it was shown. Nothing else of it changes. */
function shapeShown(shape, on) {
  const want = on ? 'visible' : 'hidden';
  let out = /visibility:\s*(?:visible|hidden|inherit|collapse)/.test(shape)
    ? shape.replace(/visibility:\s*(?:visible|hidden|inherit|collapse)/, `visibility:${want}`)
    : shape.replace(/(<v:shape\b[^>]*\sstyle=)(['"])([\s\S]*?)\2/, (m, a, q, s) => `${a}${q}${s}${s ? ';' : ''}visibility:${want}${q}`);
  const has = /<x:Visible\b/.test(out);
  if (!on && has) out = out.replace(/\s*<x:Visible\s*\/>|\s*<x:Visible>[^<]*<\/x:Visible>/g, '');
  else if (on && has && !/<x:Visible\s*\/>|<x:Visible>\s*(?:True|1)?\s*<\/x:Visible>/i.test(out)) out = out.replace(/<x:Visible>[^<]*<\/x:Visible>/, '<x:Visible/>');
  else if (on && !has) {
    const col = /(\s*)<x:Column>[^<]*<\/x:Column>/.exec(out);
    out = col ? out.slice(0, col.index + col[0].length) + `${col[1]}<x:Visible/>` + out.slice(col.index + col[0].length) : out.replace(/(\s*)<\/x:ClientData>/, '<x:Visible/>$1</x:ClientData>');
  }
  return out;
}

/** Show/Hide Note: cell (r, c)'s note on sheet si shown always (`on`) or hidden until the pointer rests on its cell,
 *  written in its VML shape as Excel writes it. False when the cell has no note's shape, or it is so already. */
export function setNoteVisible(wb, si, r, c, on) {
  const vml = relOf(wb, wb.sheets[si].path, `${REL}/vmlDrawing`)?.target;
  const text = vml && wb.pkg.text(vml);
  if (!text) return false;
  const next = text.replace(/<v:shape\b[\s\S]*?<\/v:shape>/g, (m) => (noteShapeAt(m, r, c) ? shapeShown(m, on) : m));
  if (next === text) return false;
  wb.pkg.setText(vml, next);
  wb.edited = true;
  return true;
}

/** Show All Notes: every note on every sheet shown (`on`) or every one hidden, each written as Show/Hide Note writes
 *  it. Excel's Show All Notes (Application.DisplayCommentIndicator = xlCommentAndIndicator, measured over COM) saves
 *  every note shown while it is on, and turning it off hides every note, one shown before too. The threads'
 *  placeholders are left as they are. Returns how many notes changed. */
export function setAllNotesVisible(wb, on) {
  let n = 0;
  wb.sheets.forEach((_, si) => {
    for (const [key, note] of readNotes(wb, si)) {
      if (note.kind !== 'note') continue;
      const [r, c] = key.split(',').map(Number);
      if (setNoteVisible(wb, si, r, c, on)) n++;
    }
  });
  return n;
}

/** The comments part and VML drawing of sheet si, made when it has none (threadedit.js's placeholders go there too). */
export { notesParts as notesPartsFor };

/** A comments part without the thread placeholders' authors (tc={id}) no comment uses any more, the comments'
 *  authorIds numbered again to match. */
function pruneThreadAuthors(xml) {
  const authors = [...xml.matchAll(/<author>([\s\S]*?)<\/author>/g)].map((m) => m[1]);
  const used = new Set([...xml.matchAll(/<comment\b[^>]*\sauthorId="(\d+)"/g)].map((m) => Number(m[1])));
  const keep = authors.map((a, i) => used.has(i) || !a.startsWith('tc='));
  if (keep.every(Boolean)) return xml;
  const renumber = new Map();
  let n = 0;
  keep.forEach((k, i) => { if (k) renumber.set(i, n++); });
  const list = authors.filter((_, i) => keep[i]).map((a) => `<author>${a}</author>`).join('');
  return xml.replace(/<authors>[\s\S]*?<\/authors>/, `<authors>${list}</authors>`)
    .replace(/(<comment\b[^>]*\sauthorId=")(\d+)"/g, (m, a, id) => `${a}${renumber.get(Number(id)) ?? id}"`);
}
