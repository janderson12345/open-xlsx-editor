// The tables in HTML that another program put on the clipboard (Excel, Word, Google Sheets, a browser): each cell's
// text, its spans and its CSS worked out as a browser would (the style block's rules by specificity, then the
// inline style; inherited properties from the table, row and the elements wrapping all of the cell's text). Read
// without a DOM, so it runs in Node too: a small tag scanner that tolerates Excel's and Word's unquoted, unclosed
// HTML. Used by cliphtml.js.

const VOID = new Set(['br', 'col', 'meta', 'link', 'img', 'hr', 'input', 'wbr', 'area', 'base', 'source']);
// (Properties that pass from an element to what's inside it: CSS's inherited ones, and text-decoration, which Excel
// takes from an element around the text, as <s>$3.25</s> struck through, measured.)
const INHERITED = new Set(['color', 'font-family', 'font-size', 'font-weight', 'font-style', 'text-align', 'white-space',
  'text-decoration', 'text-decoration-line', 'text-underline-style', 'mso-font-charset', 'mso-char-indent-count']);
// (What an element means without any CSS, as a browser's own style sheet gives it.)
const IMPLIED = {
  b: { 'font-weight': '700' }, strong: { 'font-weight': '700' }, i: { 'font-style': 'italic' }, em: { 'font-style': 'italic' },
  u: { 'text-decoration': 'underline' }, ins: { 'text-decoration': 'underline' }, s: { 'text-decoration': 'line-through' },
  strike: { 'text-decoration': 'line-through' }, del: { 'text-decoration': 'line-through' },
  th: { 'font-weight': '700', 'text-align': 'center' }, td: {},
};
const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: String.fromCharCode(160) };

/** Text with its character references decoded. */
export function decodeEntities(text) {
  if (!text.includes('&')) return text;
  return text.replace(/&(#x[0-9a-f]{1,6}|#[0-9]{1,7}|[a-z]{2,8});/gi, (all, e) => {
    if (e[0] === '#') {
      const n = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : Number(e.slice(1));
      return n > 0 && n <= 0x10ffff ? String.fromCodePoint(n) : all;
    }
    return ENTITIES[e.toLowerCase()] ?? all;
  });
}

/** A CSS value with its escapes undone: \0022 is ", \# is # (Excel's mso-number-format is written so). */
export function cssUnescape(text) {
  let out = '';
  for (let i = 0; i < text.length; i++) {
    if (text[i] !== '\\' || i === text.length - 1) { out += text[i]; continue; }
    let j = i + 1;
    while (j < text.length && j - i <= 6 && /[0-9a-f]/i.test(text[j])) j++;
    if (j > i + 1) {
      out += String.fromCodePoint(Math.min(parseInt(text.slice(i + 1, j), 16), 0x10ffff));
      // (One space after a hex escape ends it.)
      if (text[j] === ' ') j++;
      i = j - 1;
    } else {
      out += text[i + 1];
      i++;
    }
  }
  return out;
}

/** CSS declarations ("a:b; c:'d;e'") as a map, names lower case, values trimmed and unquoted where a whole value. */
export function parseDecls(text) {
  const out = {};
  let i = 0;
  while (i < text.length) {
    let j = i;
    let quote = null;
    while (j < text.length && (quote || text[j] !== ';')) {
      if (quote) { if (text[j] === quote) quote = null; } else if (text[j] === '"' || text[j] === "'") quote = text[j];
      j++;
    }
    const part = text.slice(i, j);
    const colon = part.indexOf(':');
    if (colon > 0) {
      const name = part.slice(0, colon).trim().toLowerCase();
      // (!important cut off with endsWith, not /\s*!important$/: that regex is quadratic on a long run of spaces in
      // clipboard text from any web page.)
      let value = part.slice(colon + 1).trim();
      if (value.toLowerCase().endsWith('!important')) value = value.slice(0, -'!important'.length).trimEnd();
      if (name) out[name] = value;
    }
    i = j + 1;
  }
  return out;
}

/** A style block's rules: [{ tag, cls, specificity, order, decls }] for simple selectors (td, .xl63, p.MsoNormal). */
function parseRules(css, rules) {
  // (A style block's <!-- and --> are CSS's CDO and CDC tokens, dropped; --!> as an HTML comment's end. The text is
  // only read for declarations, never put into the page.)
  const text = css.replace(/<!--|--!?>/g, ' ').replace(/\/\*[\s\S]*?\*\//g, ' ');
  let i = 0;
  while (i < text.length) {
    const open = text.indexOf('{', i);
    if (open < 0) break;
    const close = text.indexOf('}', open);
    if (close < 0) break;
    const selectors = text.slice(i, open).trim();
    const decls = parseDecls(text.slice(open + 1, close));
    for (const sel of selectors.split(',')) {
      const m = /^([a-z][a-z0-9]*)?(?:\.([\w-]+))?$/i.exec(sel.trim());
      if (!m || (!m[1] && !m[2])) continue;
      rules.push({ tag: m[1]?.toLowerCase() || null, cls: m[2] || null, specificity: (m[1] ? 1 : 0) + (m[2] ? 10 : 0), order: rules.length, decls });
    }
    i = close + 1;
  }
}

/** A tag's attributes ("a=1 b='x y' c") as a map, names lower case, values decoded. */
function parseAttrs(text) {
  const out = {};
  let i = 0;
  while (i < text.length) {
    while (i < text.length && /[\s/]/.test(text[i])) i++;
    let j = i;
    while (j < text.length && !/[\s=/>]/.test(text[j])) j++;
    const name = text.slice(i, j).toLowerCase();
    if (!name) { i = j + 1; continue; }
    while (j < text.length && /\s/.test(text[j])) j++;
    if (text[j] !== '=') { out[name] = ''; i = j; continue; }
    j++;
    while (j < text.length && /\s/.test(text[j])) j++;
    let value;
    if (text[j] === '"' || text[j] === "'") {
      const end = text.indexOf(text[j], j + 1);
      const stop = end < 0 ? text.length : end;
      value = text.slice(j + 1, stop);
      j = stop + 1;
    } else {
      let k = j;
      while (k < text.length && !/\s/.test(text[k])) k++;
      value = text.slice(j, k);
      j = k;
    }
    out[name] = decodeEntities(value);
    i = j;
  }
  return out;
}

/** HTML as a tree of { tag, attrs, kids, up } and { text } nodes, with the style blocks' rules. */
function parseTree(html) {
  const root = { tag: '#root', attrs: {}, kids: [], up: null, depth: 0 };
  const rules = [];
  let top = root;
  const close = (tags) => { while (top !== root && tags.includes(top.tag)) top = top.up; };
  let i = 0;
  while (i < html.length) {
    const lt = html.indexOf('<', i);
    const end = lt < 0 ? html.length : lt;
    if (end > i) top.kids.push({ text: decodeEntities(html.slice(i, end)), up: top });
    if (lt < 0) break;
    if (html.startsWith('<!--', lt)) {
      const e = html.indexOf('-->', lt + 4);
      i = e < 0 ? html.length : e + 3;
      continue;
    }
    if (html[lt + 1] === '!' || html[lt + 1] === '?') {
      const e = html.indexOf('>', lt);
      i = e < 0 ? html.length : e + 1;
      continue;
    }
    // (A < not starting a tag is text, as a browser reads it; checked before looking for the tag's end, so a run of
    // them can't make the reading slow.)
    if (!/[a-z]/i.test(html[lt + 1] === '/' ? html[lt + 2] || '' : html[lt + 1] || '')) { top.kids.push({ text: '<', up: top }); i = lt + 1; continue; }
    // (The tag's end: the first > outside quotes; with none, the rest is let go, as a browser does.)
    let j = lt + 1;
    let quote = null;
    while (j < html.length && (quote || html[j] !== '>')) {
      if (quote) { if (html[j] === quote) quote = null; } else if ((html[j] === '"' || html[j] === "'") && /[=\s]/.test(html[j - 1])) quote = html[j];
      j++;
    }
    if (j >= html.length) break;
    const inner = html.slice(lt + 1, j);
    i = j + 1;
    const m = /^(\/?)([a-z][\w:.-]*)/i.exec(inner);
    const tag = m[2].toLowerCase();
    if (m[1]) {
      // (An end tag closes up to its element, if it is open; a stray one is let go.)
      let at = top;
      while (at !== root && at.tag !== tag) at = at.up;
      if (at !== root) top = at.up;
      continue;
    }
    if (tag === 'style' || tag === 'script') {
      const e = html.toLowerCase().indexOf(`</${tag}`, i);
      const body = html.slice(i, e < 0 ? html.length : e);
      if (tag === 'style') parseRules(body, rules);
      const gt = e < 0 ? -1 : html.indexOf('>', e);
      i = gt < 0 ? html.length : gt + 1;
      continue;
    }
    // (Tags HTML leaves open: a cell ends at the next cell or row, a row at the next row, a paragraph at the next.)
    if (tag === 'td' || tag === 'th') close(['td', 'th', 'p', 'span', 'font', 'b', 'i', 'u', 's']);
    if (tag === 'tr') close(['td', 'th', 'tr', 'p', 'span', 'font', 'b', 'i', 'u', 's']);
    if (tag === 'tbody' || tag === 'thead' || tag === 'tfoot') close(['td', 'th', 'tr', 'tbody', 'thead', 'tfoot', 'p']);
    if (tag === 'p') close(['p']);
    const el = { tag, attrs: parseAttrs(inner.slice(m[0].length)), kids: [], up: top, depth: top.depth + 1 };
    top.kids.push(el);
    // (Elements more than 200 deep are read as siblings: the tree is walked by recursion.)
    if (!VOID.has(tag) && !/\/\s*$/.test(inner) && el.depth < 200) top = el;
  }
  return { root, rules };
}

/** An element's own CSS: its implied style, the matching rules by specificity and order, then its style attribute. */
function ownStyle(el, rules) {
  const out = { ...(IMPLIED[el.tag] || {}) };
  if (el.tag === 'font') {
    const inline = parseDecls(el.attrs.style || '');
    // (Excel's own copy marks a font element's colour as one to ignore: mso-ignore:color, measured.)
    if (el.attrs.color && !/color/.test(inline['mso-ignore'] || '')) out.color = el.attrs.color;
    if (el.attrs.face) out['font-family'] = el.attrs.face;
  }
  const classes = (el.attrs.class || '').split(/\s+/).filter(Boolean);
  const hits = rules.filter((r) => (!r.tag || r.tag === el.tag) && (!r.cls || classes.includes(r.cls)) && (r.tag || r.cls))
    .sort((a, b) => a.specificity - b.specificity || a.order - b.order);
  for (const r of hits) Object.assign(out, r.decls);
  Object.assign(out, parseDecls(el.attrs.style || ''));
  if (el.attrs.bgcolor && !out.background && !out['background-color']) out['background-color'] = el.attrs.bgcolor;
  return out;
}

/** The inherited properties at `el`, from the root down (computed once per element). */
function inheritedAt(el, rules, cache) {
  if (!el || el.tag === '#root') return {};
  if (cache.has(el)) return cache.get(el);
  const up = inheritedAt(el.up, rules, cache);
  const own = ownStyle(el, rules);
  const out = { ...up };
  for (const [k, v] of Object.entries(own)) if (INHERITED.has(k) && v !== 'inherit') out[k] = v;
  cache.set(el, out);
  return out;
}

/** A cell's text: whitespace run together as a browser shows it, a line break for <br> and between paragraphs. With
 *  the text nodes, for the formatting all of the text shares. */
function cellText(td) {
  // (Character by character, each with the text node it came from, so the runs of rich text follow the trimming.)
  const chars = [];
  const owners = [];
  const nodes = [];
  const nbsp = String.fromCharCode(160);
  const walk = (n) => {
    for (const k of n.kids) {
      if (k.text != null) {
        if (k.text.trim()) nodes.push(k);
        for (const ch of k.text) { chars.push(/[ \t\r\n\f]/.test(ch) ? ' ' : ch); owners.push(k); }
        continue;
      }
      if (k.tag === 'br') { chars.push('\n'); owners.push(null); continue; }
      const block = k.tag === 'p' || k.tag === 'div' || k.tag === 'li';
      if (block) { chars.push('\n'); owners.push(null); }
      walk(k);
      if (block) { chars.push('\n'); owners.push(null); }
    }
  };
  walk(td);
  // (Spaces run together and taken off each line's ends; empty lines dropped; a non-breaking space kept as a space.)
  const keep = [];
  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i];
    const prev = keep.length ? chars[keep[keep.length - 1]] : '\n';
    if (ch === ' ' && (prev === ' ' || prev === '\n')) continue;
    if (ch === '\n') {
      while (keep.length && chars[keep[keep.length - 1]] === ' ') keep.pop();
      if (prev === '\n') continue;
    }
    keep.push(i);
  }
  while (keep.length && /[ \n]/.test(chars[keep[keep.length - 1]])) keep.pop();
  const text = keep.map((i) => (chars[i] === nbsp ? ' ' : chars[i])).join('');
  // (The pieces of the text by the node each came from: a line break goes with the piece before it.)
  const pieces = [];
  for (const i of keep) {
    const owner = owners[i] || pieces[pieces.length - 1]?.node || null;
    const ch = chars[i] === nbsp ? ' ' : chars[i];
    if (pieces.length && pieces[pieces.length - 1].node === owner) pieces[pieces.length - 1].text += ch;
    else pieces.push({ text: ch, node: owner });
  }
  return { text, nodes, pieces };
}

/**
 * The first table in `html` (the CF_HTML header, if any, before it): { rows: [[cell]], cols: [{ width, userset }] },
 * each cell { text, pieces, colspan, rowspan, attrs, css } where `css` is the cell's own properties with the inherited ones
 * the table, its row and the elements around all of its text give it, and `pieces` the text by
 * the elements it is in ([{ text, css }], the inherited properties there). A fragment without a table (a line of text
 * from a page) is one cell. Null when there's nothing.
 */
export function readHtmlTable(html) {
  const lt = html.indexOf('<');
  if (lt < 0) return null;
  const { root, rules } = parseTree(html.slice(lt));
  const cache = new Map();
  const find = (n, tag) => {
    for (const k of n.kids || []) {
      if (k.tag === tag) return k;
      const f = k.tag ? find(k, tag) : null;
      if (f) return f;
    }
    return null;
  };
  const table = find(root, 'table');
  const cssOf = (td) => {
    const css = { ...inheritedAt(td.up, rules, cache), ...ownStyle(td, rules) };
    const { text, nodes, pieces } = cellText(td);
    // (What the elements inside give every piece of the text: <b>Avon</b> is a bold cell; "Both <b>columns</b>" isn't.)
    if (nodes.length) {
      const each = nodes.map((n) => inheritedAt(n.up, rules, cache));
      for (const k of INHERITED) {
        const v = each[0][k];
        if (v !== undefined && each.every((x) => x[k] === v)) css[k] = v;
      }
    }
    // (Each piece of the text with what is inherited where it is, for rich text.)
    return { text, css, pieces: pieces.map((p) => ({ text: p.text, css: p.node ? inheritedAt(p.node.up, rules, cache) : css })) };
  };
  if (!table) {
    const body = find(root, 'body') || root;
    const { text, css, pieces } = cssOf(body);
    return text ? { rows: [[{ text, colspan: 1, rowspan: 1, attrs: {}, css, pieces }]], cols: [] } : null;
  }
  const rows = [];
  const cols = [];
  const walk = (n) => {
    for (const k of n.kids) {
      if (!k.tag || k.tag === 'table') continue;
      if (k.tag === 'col') {
        const css = ownStyle(k, rules);
        const span = Math.max(1, Number(k.attrs.span) || 1);
        const width = Number(k.attrs.width) || (/px$/.test(css.width || '') ? parseFloat(css.width) : null);
        for (let s = 0; s < span; s++) cols.push({ width: width || null, userset: /userset/.test(css['mso-width-source'] || '') });
        continue;
      }
      if (k.tag === 'tr') {
        const cells = [];
        for (const td of k.kids) {
          if (td.tag !== 'td' && td.tag !== 'th') continue;
          const { text, css, pieces } = cssOf(td);
          cells.push({ text, pieces, colspan: Math.max(1, Number(td.attrs.colspan) || 1), rowspan: Math.max(1, Number(td.attrs.rowspan) || 1), attrs: td.attrs, css });
        }
        rows.push(cells);
        continue;
      }
      walk(k);
    }
  };
  walk(table);
  return { rows, cols };
}
