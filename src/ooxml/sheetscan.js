// A worksheet part's cells read from its bytes, without parsing the XML into a tree: sheetData is most of a large
// sheet (a customer's 174 MB part took DOMParser 22 s and the walk over its tree 12 s more, and ran Node out of
// memory), and it is regular: rows of cells, each a few attributes, a formula, a value or an inline string. The rest
// of the part (views, columns, merges, conditional formats, the extLst...) is small and still parsed (sheet.js
// readSheetBytes). What isn't understood here (a document type, another prefix for the sheet's own elements, a
// namespace declared inside sheetData, CDATA or an element inside a value) gives null: the part is then parsed as before.
import { parseRef, MAX_ROWS, MAX_COLS } from './refs.js';
import { decodeXstring } from './xml.js';
import { SheetCells, makeCell, withRuns, rowEntry, formulaKind, blankLine } from './sheetcells.js';

const LT = 0x3c;
const GT = 0x3e;
const SLASH = 0x2f;
const COLON = 0x3a;
const EQ = 0x3d;
const QUOT = 0x22;
const APOS = 0x27;
const BANG = 0x21;
const QUERY = 0x3f;
const DASH = 0x2d;
const AMP = 0x26;
const space = (b) => b === 0x20 || b === 0x09 || b === 0x0a || b === 0x0d;
/** After a tag's name: a space, '>' or '/'. */
const ends = (b) => b === GT || b === SLASH || space(b);
const bytes = (s) => Uint8Array.from(s, (ch) => ch.charCodeAt(0));
const decoder = new TextDecoder();
const ONE = Array.from({ length: 128 }, (_, k) => String.fromCharCode(k));

/** Thrown inside a scan for XML it doesn't read (the caller parses the part instead), and for rows out of order. */
const BAIL = { bail: true };
const UNORDERED = { unordered: true };

/** The index after the comment, processing instruction or CDATA section starting at i ('<!' or '<?' there). */
function afterMarkup(u8, i) {
  if (u8[i + 1] === QUERY) {
    for (let j = u8.indexOf(GT, i + 2); j >= 0; j = u8.indexOf(GT, j + 1)) if (u8[j - 1] === QUERY) return j + 1;
    throw BAIL;
  }
  if (u8[i + 2] === DASH && u8[i + 3] === DASH) {
    for (let j = u8.indexOf(GT, i + 6); j >= 0; j = u8.indexOf(GT, j + 1)) if (u8[j - 1] === DASH && u8[j - 2] === DASH) return j + 1;
    throw BAIL;
  }
  if (u8[i + 2] === 0x5b) { // (<![CDATA[ ... ]]>)
    for (let j = u8.indexOf(GT, i + 9); j >= 0; j = u8.indexOf(GT, j + 1)) if (u8[j - 1] === 0x5d && u8[j - 2] === 0x5d) return j + 1;
    throw BAIL;
  }
  throw BAIL; // (a document type: entities of its own)
}

/** The index of the '>' ending the tag that starts at i, quoted attribute values (which may hold '>') passed over. */
function tagEnd(u8, i) {
  for (let j = i + 1; j < u8.length; j++) {
    const b = u8[j];
    if (b === GT) return j;
    if (b === QUOT || b === APOS) {
      j = u8.indexOf(b, j + 1);
      if (j < 0) break;
    }
  }
  throw BAIL;
}

/** The index after the element whose start tag is at i, whatever it holds. */
function afterElement(u8, i) {
  let gt = tagEnd(u8, i);
  if (u8[gt - 1] === SLASH) return gt + 1;
  let depth = 1;
  let j = gt + 1;
  while (depth > 0) {
    const lt = u8.indexOf(LT, j);
    if (lt < 0) throw BAIL;
    const b = u8[lt + 1];
    if (b === SLASH) { depth--; gt = u8.indexOf(GT, lt); if (gt < 0) throw BAIL; j = gt + 1; }
    else if (b === BANG || b === QUERY) j = afterMarkup(u8, lt);
    else { gt = tagEnd(u8, lt); if (u8[gt - 1] !== SLASH) depth++; j = gt + 1; }
  }
  return j;
}

/** Whether the bytes at i are `name` (bytes) followed by a tag name's end. */
function named(u8, i, name) {
  for (let k = 0; k < name.length; k++) if (u8[i + k] !== name[k]) return false;
  return ends(u8[i + name.length]);
}

const SHEET_DATA = bytes('sheetData');

/**
 * Where a worksheet part's sheetData is: { start (its '<'), inner (after its start tag), innerEnd (its closing tag's
 * '<'), end (after the element), prefix (the sheet's elements' prefix with its colon, '' as Excel writes them),
 * rootOpen and rootClose (the root's start tag with its namespace declarations and its closing tag, to parse a piece
 * of the part alone) }, or null when the part isn't read this way.
 */
export function findSheetData(u8) {
  try {
    let i = u8.indexOf(LT);
    // (The declaration, comments and processing instructions before the root.)
    while (i >= 0 && (u8[i + 1] === BANG || u8[i + 1] === QUERY)) i = u8.indexOf(LT, afterMarkup(u8, i));
    if (i < 0) return null;
    let j = i + 1;
    while (j < u8.length && !ends(u8[j])) j++;
    const colon = u8.indexOf(COLON, i);
    const prefix = colon > 0 && colon < j ? u8.slice(i + 1, colon + 1) : new Uint8Array(0);
    const rootName = decoder.decode(u8.subarray(i + 1, j));
    const rootGt = tagEnd(u8, i);
    if (u8[rootGt - 1] === SLASH) return null;
    const rootOpen = decoder.decode(u8.subarray(i, rootGt + 1));
    const tag = new Uint8Array(prefix.length + SHEET_DATA.length);
    tag.set(prefix);
    tag.set(SHEET_DATA, prefix.length);
    for (j = rootGt + 1; ;) {
      const lt = u8.indexOf(LT, j);
      if (lt < 0 || u8[lt + 1] === SLASH) return null;
      if (u8[lt + 1] === BANG || u8[lt + 1] === QUERY) { j = afterMarkup(u8, lt); continue; }
      if (!named(u8, lt + 1, tag)) { j = afterElement(u8, lt); continue; }
      const gt = tagEnd(u8, lt);
      // (A namespace declared on sheetData itself could change what its elements' prefixes mean.)
      if (decoder.decode(u8.subarray(lt, gt)).includes('xmlns')) return null;
      const empty = u8[gt - 1] === SLASH;
      const end = empty ? gt + 1 : closeOf(u8, gt + 1, tag);
      return { start: lt, inner: gt + 1, innerEnd: empty ? gt + 1 : u8.lastIndexOf(LT, end - 1), end, prefix: decoder.decode(prefix), rootOpen, rootClose: `</${rootName}>` };
    }
  } catch (e) {
    if (e === BAIL) return null;
    throw e;
  }
}

/** The index after sheetData's closing tag, looked for from the end of the part (nothing after it holds one, and the
 *  rows between aren't walked twice). */
function closeOf(u8, from, tag) {
  for (let lt = u8.lastIndexOf(LT); lt >= from; lt = u8.lastIndexOf(LT, lt - 1)) {
    if (u8[lt + 1] === SLASH && named(u8, lt + 2, tag)) {
      const gt = u8.indexOf(GT, lt);
      if (gt < 0) break;
      return gt + 1;
    }
  }
  throw BAIL;
}

const ENTITIES = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };
/** A decoded string's entities and character references. */
function unescape(s) {
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|\w+);?/g, (m, name) => {
    if (m[m.length - 1] !== ';') throw BAIL;
    if (name[0] === '#') {
      const code = name[1] === 'x' ? parseInt(name.slice(2), 16) : parseInt(name.slice(1), 10);
      if (!(code <= 0x10ffff)) throw BAIL;
      return String.fromCodePoint(code);
    }
    if (!(name in ENTITIES)) throw BAIL;
    return ENTITIES[name];
  });
}

const POW10 = [1, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9, 1e10, 1e11, 1e12, 1e13, 1e14, 1e15];
const T_STR = bytes('str');
const T_INLINE = bytes('inlineStr');

/**
 * The cells of a worksheet part's sheetData (`at`: findSheetData's answer) as readSheet's model has them:
 * { cells, rows, arrays, maxRow, maxCol }, or null for XML not read here. `compact`: rows of formatted empty cells
 * only are kept as bands (sheetcells.js SheetCells) when there are at least that many such cells in all; Infinity,
 * never. `parseIs(xml)` gives an inline string's `is` element parsed, for the ones that aren't a single `t`.
 */
export function scanSheetData(u8, at, sst, { compact = 50000, parseIs }) {
  for (const banded of compact === Infinity ? [false] : [true, false]) {
    try {
      const data = scan(u8, at, sst, banded, parseIs);
      if (banded) {
        if (data.cells.bandCells < compact) data.cells = new Map(data.cells);
      }
      return data;
    } catch (e) {
      if (e === BAIL) return null;
      if (e !== UNORDERED) throw e;
    }
  }
  return null;
}

function scan(u8, at, sst, banded, parseIs) {
  const P = at.prefix.length;
  const pfx = bytes(at.prefix);
  const end = at.innerEnd;
  const cells = banded ? new SheetCells() : new Map();
  const rows = new Map();
  const shared = new Map();
  const arrays = [];
  const sstRuns = sst.runs;
  let maxRow = -1;
  let maxCol = -1;

  // An attribute read by `attribute`: its name [n0, n1) and value [v0, v1).
  let n0 = 0;
  let n1 = 0;
  let v0 = 0;
  let v1 = 0;
  /** The attribute at j (after the tag's name or the attribute before): the index after it, or -1 at the tag's end
   *  (`tagClose` is then the index after the tag, `selfClosed` whether it was `/>`). */
  let tagClose = 0;
  let selfClosed = false;
  const attribute = (j) => {
    let b = u8[j];
    while (space(b)) b = u8[++j];
    if (b === GT) { tagClose = j + 1; selfClosed = false; return -1; }
    if (b === SLASH) { if (u8[j + 1] !== GT) throw BAIL; tagClose = j + 2; selfClosed = true; return -1; }
    n0 = j;
    while (b !== EQ && !space(b)) { if (b === undefined || b === GT || b === LT) throw BAIL; b = u8[++j]; }
    n1 = j;
    while (space(b)) b = u8[++j];
    if (b !== EQ) throw BAIL;
    b = u8[++j];
    while (space(b)) b = u8[++j];
    if (b !== QUOT && b !== APOS) throw BAIL;
    v0 = j + 1;
    v1 = u8.indexOf(b, v0);
    if (v1 < 0) throw BAIL;
    // (A namespace declared in here, or an attribute with the sheet's own prefix, isn't read by name alone.)
    if (n1 - n0 >= 5 && u8[n0] === 0x78 && u8[n0 + 1] === 0x6d && u8[n0 + 2] === 0x6c && u8[n0 + 3] === 0x6e && u8[n0 + 4] === 0x73) throw BAIL;
    if (P && n1 - n0 > P) { let k = 0; while (k < P && u8[n0 + k] === pfx[k]) k++; if (k === P) throw BAIL; }
    return v1 + 1;
  };

  /** Bytes [a, b) as text: an attribute's value (`inAttr`: tabs and line breaks are spaces there) or an element's. */
  const text = (a, b, inAttr) => {
    const n = b - a;
    let plain = true;
    for (let k = a; k < b; k++) { const ch = u8[k]; if (ch >= 0x80 || ch === AMP || ch < 0x20) { plain = false; break; } }
    if (plain) {
      if (n === 0) return '';
      if (n === 1) return ONE[u8[a]];
      if (n <= 24) { let s = ONE[u8[a]]; for (let k = a + 1; k < b; k++) s += ONE[u8[k]]; return s; }
      return decoder.decode(u8.subarray(a, b));
    }
    let s = decoder.decode(u8.subarray(a, b));
    if (s.includes('\r')) s = s.replace(/\r\n?/g, '\n');
    if (inAttr) s = s.replace(/[\t\n]/g, ' ');
    if (s.includes('<')) throw BAIL;
    return s.includes('&') ? unescape(s) : s;
  };
  /** Whether the attribute read is named `name` (a string). */
  const isName = (name) => {
    if (n1 - n0 !== name.length) return false;
    for (let k = 0; k < name.length; k++) if (u8[n0 + k] !== name.charCodeAt(k)) return false;
    return true;
  };
  /** The value read as a whole number of up to nine digits, or -1 when it isn't written as one. */
  const digits = () => {
    const n = v1 - v0;
    if (n < 1 || n > 9) return -1;
    let x = 0;
    for (let k = v0; k < v1; k++) { const d = u8[k] - 0x30; if (d < 0 || d > 9) return -1; x = x * 10 + d; }
    return x;
  };
  /** Bytes [a, b) as the number Number() makes of them, read in place when they are plain digits (at most fifteen,
   *  with or without a sign and a decimal point: the division is then exact to the last bit). */
  const number = (a, b) => {
    let k = a;
    const neg = u8[k] === DASH;
    if (neg) k++;
    let m = 0;
    let count = 0;
    let frac = -1;
    for (; k < b; k++) {
      const d = u8[k] - 0x30;
      if (d >= 0 && d <= 9) { m = m * 10 + d; count++; if (frac >= 0) frac++; } else if (u8[k] === 0x2e && frac < 0) frac = 0; else { count = 99; break; }
    }
    if (count < 1 || count > 15 || (frac === 0)) return Number(text(a, b, false));
    const x = frac > 0 ? m / POW10[frac] : m;
    return neg ? -x : x;
  };
  /** The element whose start tag ends before j: its text's bytes as [a, b) in t0, t1, and the index after it. Text
   *  alone: an element or CDATA inside isn't read here. */
  let t0 = 0;
  let t1 = 0;
  const content = (j) => {
    const lt = u8.indexOf(LT, j);
    if (lt < 0 || u8[lt + 1] !== SLASH) throw BAIL;
    t0 = j;
    t1 = lt;
    const gt = u8.indexOf(GT, lt);
    if (gt < 0) throw BAIL;
    return gt + 1;
  };
  /** Whether the tag at j ('<' there) is the sheet's element `ch` (one letter). */
  const tag1 = (j, ch) => {
    for (let k = 0; k < P; k++) if (u8[j + 1 + k] !== pfx[k]) return false;
    return u8[j + 1 + P] === ch && ends(u8[j + 2 + P]);
  };

  /** An inline string's text and runs from its `is` element, bytes [a, b): [text, runs]. */
  const inline = (a, b) => {
    // (As nearly all are written: <is><t>text</t></is>.)
    let j = a + 1 + P + 2;
    if (u8[j] === GT) {
      j++;
      while (space(u8[j])) j++;
      if (u8[j] === LT && tag1(j, 0x74)) {
        let k = attribute(j + 2 + P);
        while (k >= 0) k = attribute(k);
        if (!selfClosed) {
          const after = content(tagClose);
          const a0 = t0;
          const a1 = t1;
          k = after;
          while (space(u8[k])) k++;
          if (u8[k] === LT && u8[k + 1] === SLASH && u8.indexOf(GT, k) + 1 === b) return [decodeXstring(text(a0, a1, false)), null];
        }
      }
    }
    return parseIs(decoder.decode(u8.subarray(a, b)));
  };

  let i = at.inner;
  let r = -1;
  // (The row being read while it has formatted empty cells only: its runs [c0, c1, s, ...], `runN` entries used.)
  const run = [];
  for (;;) {
    let lt = u8.indexOf(LT, i);
    if (lt < 0 || lt >= end) break;
    const b1 = u8[lt + 1];
    if (b1 === BANG || b1 === QUERY) { i = afterMarkup(u8, lt); continue; }
    if (b1 === SLASH) throw BAIL;
    // (`row`: anything else in sheetData is passed over, as the tree's walk passed it.)
    let isRow = u8[lt + 1 + P] === 0x72 && u8[lt + 2 + P] === 0x6f && u8[lt + 3 + P] === 0x77 && ends(u8[lt + 4 + P]);
    for (let k = 0; isRow && k < P; k++) if (u8[lt + 1 + k] !== pfx[k]) isRow = false;
    if (!isRow) { i = afterElement(u8, lt); continue; }

    let rAt = null;
    let ht = null;
    let hidden = null;
    let customFormat = null;
    let sAt = null;
    let level = null;
    let collapsed = null;
    let thickBot = null;
    let thickTop = null;
    let customHeight = null;
    let rowR = -2;
    let any = false;
    for (let j = attribute(lt + 4 + P); j >= 0; j = attribute(j)) {
      const n = n1 - n0;
      if (n === 1) {
        if (u8[n0] === 0x72) { rowR = digits(); if (rowR < 1) { rowR = -2; rAt = text(v0, v1, true); } } else if (u8[n0] === 0x73) sAt = text(v0, v1, true);
      } else if (n === 2) { if (isName('ht')) { ht = text(v0, v1, true); any = true; } } else if (n === 6) { if (isName('hidden')) { hidden = text(v0, v1, true); any = true; } } else if (n === 8) { if (isName('thickBot')) { thickBot = text(v0, v1, true); any = true; } else if (isName('thickTop')) { thickTop = text(v0, v1, true); any = true; } } else if (n === 9) { if (isName('collapsed')) { collapsed = text(v0, v1, true); any = true; } } else if (n === 12) {
        if (isName('customFormat')) { customFormat = text(v0, v1, true); any = true; } else if (isName('outlineLevel')) { level = text(v0, v1, true); any = true; } else if (isName('customHeight')) customHeight = text(v0, v1, true);
      }
    }
    const before = r;
    r = rowR >= 1 ? rowR - 1 : rAt ? Number(rAt) - 1 : r + 1;
    // (Bands take rows in order, each once.)
    if (banded && !(r > before && Number.isInteger(r))) throw UNORDERED;
    if (any) {
      const entry = rowEntry(ht, hidden, customFormat, sAt, level, collapsed, thickBot, customHeight, thickTop);
      if (entry) rows.set(r, entry);
    }
    i = tagClose;
    if (selfClosed) continue;

    let c = -1;
    let line = null;
    let runN = 0;
    let blankRow = banded;
    for (;;) {
      lt = u8.indexOf(LT, i);
      if (lt < 0 || lt >= end) throw BAIL;
      const b = u8[lt + 1];
      if (b === SLASH) { i = u8.indexOf(GT, lt) + 1; if (i <= 0) throw BAIL; break; }
      if (b === BANG || b === QUERY) { i = afterMarkup(u8, lt); continue; }
      if (!tag1(lt, 0x63)) { i = afterElement(u8, lt); continue; }

      // A cell: its attributes.
      let col = -1;
      let s = 0;
      let type = 0; // 0 a number (as any type not named below), else the type's first letter; 1 'str', 2 'inlineStr'
      let cm = false;
      // (As nearly every cell starts, ` r="B12" s="3"`, read in place; the rest of its attributes, and any cell written
      // another way, by the loop below.)
      let from = lt + 2 + P;
      if (u8[from] === 0x20 && u8[from + 1] === 0x72 && u8[from + 2] === EQ && u8[from + 3] === QUOT) {
        const k0 = from + 4;
        let k = k0;
        let x = 0;
        let ch = u8[k];
        while (ch >= 0x41 && ch <= 0x5a && k - k0 < 3) { x = x * 26 + (ch - 0x40); ch = u8[++k]; }
        const d0 = k;
        let row = 0;
        while (ch >= 0x30 && ch <= 0x39 && k - d0 < 7) { row = row * 10 + (ch - 0x30); ch = u8[++k]; }
        if (d0 > k0 && k > d0 && ch === QUOT) {
          col = row >= 1 && row <= MAX_ROWS && x <= MAX_COLS ? x - 1 : -1;
          from = k + 1;
          if (u8[from] === 0x20 && u8[from + 1] === 0x73 && u8[from + 2] === EQ && u8[from + 3] === QUOT) {
            const s0 = from + 4;
            let sv = 0;
            k = s0;
            ch = u8[k];
            while (ch >= 0x30 && ch <= 0x39 && k - s0 < 9) { sv = sv * 10 + (ch - 0x30); ch = u8[++k]; }
            if (k > s0 && ch === QUOT) { s = sv; from = k + 1; }
          }
        }
      }
      for (let j = attribute(from); j >= 0; j = attribute(j)) {
        const n = n1 - n0;
        const a = u8[n0];
        if (n === 1 && a === 0x72) {
          // ("B12", as nearly all are written; anything else as parseRef reads it.)
          let k = v0;
          let x = 0;
          let ch = u8[k];
          while (ch >= 0x41 && ch <= 0x5a && k - v0 < 3) { x = x * 26 + (ch - 0x40); ch = u8[++k]; }
          const letters = k - v0;
          let row = 0;
          let ok = letters > 0 && v1 - k >= 1 && v1 - k <= 7;
          for (; ok && k < v1; k++) { const d = u8[k] - 0x30; if (d < 0 || d > 9) ok = false; else row = row * 10 + d; }
          if (ok) col = row >= 1 && row <= MAX_ROWS && x <= MAX_COLS ? x - 1 : -1;
          else { const p = parseRef(text(v0, v1, true)); col = p ? p.c : -1; }
        } else if (n === 1 && a === 0x73) {
          s = digits();
          if (s < 0) s = Number(text(v0, v1, true) || 0);
        } else if (n === 1 && a === 0x74) {
          const m = v1 - v0;
          const t = u8[v0];
          if (m === 1 && (t === 0x73 || t === 0x62 || t === 0x65 || t === 0x64)) type = t;
          else if (m === 3 && named3(u8, v0, T_STR)) type = 1;
          else if (m === 9 && named3(u8, v0, T_INLINE)) type = 2;
          else if (m > 1 && u8.indexOf(AMP, v0) < v1 && u8.indexOf(AMP, v0) >= v0) {
            // (Written with a character reference: read as text.)
            const name = text(v0, v1, true);
            type = name === 's' || name === 'b' || name === 'e' || name === 'd' ? name.charCodeAt(0) : name === 'str' ? 1 : name === 'inlineStr' ? 2 : 0;
          }
        } else if (n === 2 && a === 0x63 && u8[n0 + 1] === 0x6d) cm = true;
      }
      c = col >= 0 ? col : c + 1;
      if (c > maxCol) maxCol = c;
      i = tagClose;

      if (selfClosed && type === 0 && blankRow && Number.isInteger(s) && s >= 0 && (runN === 0 || c > run[runN - 2])) {
        // A formatted empty cell of a row that has had nothing else: one more of its last run, or a new run.
        if (runN && run[runN - 2] === c - 1 && run[runN - 1] === s) run[runN - 2] = c;
        else { run[runN] = c; run[runN + 1] = c; run[runN + 2] = s; runN += 3; }
        continue;
      }
      if (!line) {
        line = blankRow && runN ? blankLine(run, runN) : new Map();
        cells.set(r, line);
        blankRow = false;
      }
      let f = null;
      let fT = null;
      let fSi = null;
      let fRef = null;
      let is0 = -1;
      let is1 = -1;
      let hasV = false;
      let val0 = 0;
      let val1 = 0;
      if (!selfClosed) {
        // Its children: the first f, v and is; anything else (an extLst) passed over.
        for (;;) {
          lt = u8.indexOf(LT, i);
          if (lt < 0 || lt >= end) throw BAIL;
          const b2 = u8[lt + 1];
          if (b2 === SLASH) { i = u8.indexOf(GT, lt) + 1; if (i <= 0) throw BAIL; break; }
          if (b2 === BANG || b2 === QUERY) { i = afterMarkup(u8, lt); continue; }
          if (tag1(lt, 0x76) && !hasV) {
            let j = attribute(lt + 2 + P);
            while (j >= 0) j = attribute(j);
            hasV = true;
            if (selfClosed) { val0 = val1 = 0; i = tagClose; } else { i = content(tagClose); val0 = t0; val1 = t1; }
          } else if (tag1(lt, 0x66) && f === null) {
            for (let j = attribute(lt + 2 + P); j >= 0; j = attribute(j)) {
              if (isName('t')) fT = text(v0, v1, true); else if (isName('si')) fSi = text(v0, v1, true); else if (isName('ref')) fRef = text(v0, v1, true);
            }
            if (selfClosed) { f = ''; i = tagClose; } else { i = content(tagClose); f = text(t0, t1, false); }
          } else if (is0 < 0 && u8[lt + 1 + P] === 0x69 && u8[lt + 2 + P] === 0x73 && ends(u8[lt + 3 + P]) && tag1p(lt)) {
            is0 = lt;
            i = is1 = afterElement(u8, lt);
          } else i = afterElement(u8, lt);
        }
      }
      let cell;
      if (type === 0) cell = makeCell('n', !hasV || val0 === val1 ? null : number(val0, val1), s, f);
      else if (type === 0x73) {
        if (!hasV) cell = makeCell('s', null, s, f);
        else {
          const k = number(val0, val1);
          cell = withRuns(makeCell('s', sst[k] ?? '', s, f), sstRuns?.get(k));
        }
      } else if (type === 2) {
        if (is0 < 0) cell = makeCell('s', '', s, f);
        else { const [str, runs] = inline(is0, is1); cell = withRuns(makeCell('s', str, s, f), runs); }
      } else {
        const v = hasV ? text(val0, val1, false) : null;
        if (type === 0x62) cell = makeCell('b', v == null ? null : v === '1' || v === 'true', s, f);
        else cell = makeCell(type === 1 ? 'str' : type === 0x65 ? 'e' : 'd', v, s, f);
      }
      if (f !== null) formulaKind(cell, fT, fSi, fRef, r, c, shared, arrays, cm);
      line.set(c, cell);
    }
    if (blankRow && runN) cells.addBlankRow(r, run, runN);
    if ((line || runN) && r > maxRow) maxRow = r;
  }
  return { cells, rows, arrays, maxRow, maxCol };

  /** Whether the tag at j has the sheet's prefix (its name looked at by the caller). */
  function tag1p(j) {
    for (let k = 0; k < P; k++) if (u8[j + 1 + k] !== pfx[k]) return false;
    return true;
  }
}

/** Whether the bytes from i are exactly `name`. */
function named3(u8, i, name) {
  for (let k = 0; k < name.length; k++) if (u8[i + k] !== name[k]) return false;
  return true;
}
