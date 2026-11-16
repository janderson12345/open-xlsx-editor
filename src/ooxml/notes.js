// A sheet's notes and threaded comments, read as Excel 365 writes them (docs/NOTES.md): a note's text in the comments
// part (its runs, the author's name in bold first as Excel's Insert Note puts it), and a threaded comment with its
// replies in the threadedComments part, its people in the workbook's persons part. A threaded comment also has a
// placeholder note for older versions of Excel ("[Threaded comment]..."), which Excel 365 doesn't show: nor do we.
// Read only; the parts are saved as they are.
import { kids, kid, attr, NS } from './xml.js';
import { parseRef } from './refs.js';

const THREADED = 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment';
const PERSON = 'http://schemas.microsoft.com/office/2017/10/relationships/person';
const TC_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';

/** The text of an element's t children, in order (a run's, or the comment's own). */
const textOf = (el, ns) => kids(el, 't', ns).map((t) => t.textContent).join('');

/**
 * Sheet i's notes and threaded comments by cell: Map("r,c" => { kind: 'note', author, runs: [{ text, bold }], text }
 * or { kind: 'thread', items: [{ author, text, date, mentions: [{ start, length, name }] }] (the first, then its
 * replies), text }).
 */
export function readNotes(wb, i) {
  const out = new Map();
  const path = wb.sheets[i]?.path;
  if (!path) return out;
  const rels = wb.pkg.readRels(path);
  const people = new Map();
  const personPart = wb.pkg.rels.find((r) => r.type === PERSON)?.target;
  const personDoc = personPart && wb.pkg.xml(personPart);
  if (personDoc) for (const p of kids(personDoc.documentElement, 'person', TC_NS)) people.set(attr(p, 'id', null), attr(p, 'displayName', null));

  const threadPart = rels.find((r) => r.type === THREADED)?.target;
  const threadDoc = threadPart && wb.pkg.xml(threadPart);
  if (threadDoc) {
    const all = kids(threadDoc.documentElement, 'threadedComment', TC_NS).map((el) => ({
      ref: attr(el, 'ref', null), id: attr(el, 'id', null), parent: attr(el, 'parentId', null), done: attr(el, 'done', null) === '1',
      author: people.get(attr(el, 'personId', null)) || '', date: attr(el, 'dT', null) || '', text: kid(el, 'text', TC_NS)?.textContent ?? '',
      // (@mentions: where each is in the text, and whom it names.)
      mentions: kids(kid(el, 'mentions', TC_NS), 'mention', TC_NS).map((m) => ({
        start: Number(attr(m, 'startIndex', null) || 0), length: Number(attr(m, 'length', null) || 0), name: people.get(attr(m, 'mentionpersonId', null)) || '',
      })),
    }));
    for (const top of all.filter((x) => !x.parent)) {
      const at = parseRef(top.ref);
      if (!at) continue;
      const items = [top, ...all.filter((x) => x.parent === top.id)].map(({ author, text, date, mentions }) => ({ author, text, date, mentions }));
      // (Resolved: done="1" on the thread's first comment, as Excel writes it.)
      out.set(`${at.r},${at.c}`, { kind: 'thread', items, resolved: top.done, text: items.map((x) => `${x.author}: ${x.text}`).join('\n') });
    }
  }

  const notePart = rels.find((r) => r.type === `${NS.r}/comments`)?.target;
  const noteDoc = notePart && wb.pkg.xml(notePart);
  if (noteDoc) {
    const root = noteDoc.documentElement;
    const ns = root.namespaceURI;
    const authors = kids(kid(root, 'authors', ns), 'author', ns).map((a) => a.textContent);
    for (const c of kids(kid(root, 'commentList', ns), 'comment', ns)) {
      const at = parseRef(attr(c, 'ref'));
      if (!at || out.has(`${at.r},${at.c}`)) continue;
      const body = kid(c, 'text', ns);
      const runs = [];
      if (body) {
        const own = textOf(body, ns);
        if (own) runs.push({ text: own, bold: false });
        for (const r of kids(body, 'r', ns)) runs.push({ text: textOf(r, ns), bold: !!kid(kid(r, 'rPr', ns), 'b', ns) });
      }
      const author = authors[Number(attr(c, 'authorId') || 0)] || '';
      out.set(`${at.r},${at.c}`, { kind: 'note', author: author.startsWith('tc=') ? '' : author, runs, text: runs.map((x) => x.text).join('') });
    }
  }
  // (Each note's box as its VML shape has it: where its corners are, and whether it is always shown.)
  for (const [key, shape] of noteShapes(wb, rels)) { const n = out.get(key); if (n?.kind === 'note') Object.assign(n, shape); }
  return out;
}

/**
 * The notes' shapes in the sheet's VML drawing, by cell: Map("r,c" => { anchor: { c0, dx0, r0, dy0, c1, dx1, r1, dy1 }
 * (x:Anchor: the box's top-left and bottom-right corners, a column or row and a pixel offset into it), visible }).
 */
function noteShapes(wb, rels) {
  const out = new Map();
  const vml = rels.find((r) => r.type === `${NS.r}/vmlDrawing`)?.target;
  const text = vml && wb.pkg.text(vml);
  if (!text) return out;
  for (const m of text.matchAll(/<v:shape\b[\s\S]*?<\/v:shape>/g)) {
    const shape = m[0];
    if (!/ObjectType="Note"/.test(shape)) continue;
    const row = /<x:Row>(\d+)<\/x:Row>/.exec(shape);
    const col = /<x:Column>(\d+)<\/x:Column>/.exec(shape);
    if (!row || !col) continue;
    const a = /<x:Anchor>([^<]*)<\/x:Anchor>/.exec(shape)?.[1].split(',').map((x) => Number(x.trim()));
    const anchor = a?.length === 8 && a.every(Number.isFinite) ? { c0: a[0], dx0: a[1], r0: a[2], dy0: a[3], c1: a[4], dx1: a[5], r1: a[6], dy1: a[7] } : null;
    out.set(`${row[1]},${col[1]}`, { anchor, visible: /<x:Visible\s*\/>|<x:Visible>\s*(True|1)?\s*<\/x:Visible>/i.test(shape) });
  }
  return out;
}
