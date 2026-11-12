// A formula between the file's form and what a person types and sees: newer functions' _xlfn. prefixes, Excel 365's
// dynamic-array parts (_xlfn.ANCHORARRAY(D1) for D1#, _xlfn.SINGLE for @, _xlpm. names, _xleta. functions) and trim
// references (_xlfn._TRO_TRAILING(A1:E8) for A1:.E8). Split from tokens.js, whose tokens it reads.
import { tokenize, mapRefs, areaText } from './tokens.js';
import { fileName } from './prefixed.js';
import { functionInfo } from './catalog.js';

/** Where the bracket opened at token k (a "(" punct, or the one after a func) closes: the index of its ")". */
function closing(toks, k) {
  let depth = 0;
  for (let j = k; j < toks.length; j++) {
    if (toks[j].type === 'punct' && toks[j].text === '(') depth++;
    else if (toks[j].type === 'punct' && toks[j].text === ')' && --depth === 0) return j;
  }
  return toks.length - 1;
}
/** The text with some tokens' spans replaced ([{ from, to (token indexes, inclusive), text }]). */
function splice(text, toks, edits) {
  let out = '';
  let at = 0;
  for (const e of edits.sort((a, b) => toks[a.from].start - toks[b.from].start)) {
    out += text.slice(at, toks[e.from].start) + e.text;
    at = toks[e.to].end;
  }
  return out + text.slice(at);
}

/** A typed formula's dynamic-array parts in the file's form, as Excel 365 writes them: D1# as _xlfn.ANCHORARRAY(D1),
 *  @x as _xlfn.SINGLE(x), and LET's and LAMBDA's names as _xlpm.name. */
function dynamicFileForm(typed) {
  const toks = tokenize(typed);
  const edits = [];
  const params = new Map(); // token index of a name -> its file form
  toks.forEach((t, k) => {
    const name = t.type === 'func' ? t.text.toUpperCase().replace(/^(_XL(FN|WS)\.)+/, '') : null;
    if (name !== 'LET' && name !== 'LAMBDA') return;
    // (Its arguments, split at its own commas: LET's names are every other one but the last, LAMBDA's all but the last.)
    const end = closing(toks, k + 1);
    const args = [[]];
    let depth = 0;
    for (let j = k + 2; j < end; j++) {
      const x = toks[j];
      if (x.type === 'punct' && (x.text === '(' || x.text === '{')) depth++;
      if (x.type === 'punct' && (x.text === ')' || x.text === '}')) depth--;
      if (depth === 0 && x.type === 'punct' && x.text === ',') args.push([]);
      else args[args.length - 1].push(j);
    }
    const names = args.slice(0, -1).filter((_, a) => name === 'LAMBDA' || a % 2 === 0)
      .map((a) => (a.length === 1 && toks[a[0]].type === 'name' ? toks[a[0]].text.replace(/^_xlpm\./i, '').toLowerCase() : null)).filter(Boolean);
    for (let j = k + 2; j < end; j++) if (toks[j].type === 'name' && names.includes(toks[j].text.replace(/^_xlpm\./i, '').toLowerCase())) params.set(j, `_xlpm.${toks[j].text.replace(/^_xlpm\./i, '')}`);
  });
  for (const [k, text] of params) edits.push({ from: k, to: k, text });
  // (A function named on its own where a LAMBDA goes, an eta LAMBDA, =MAP(A1:A3,SQRT): _xleta.SQRT in the file, as
  // Excel 365 writes it; HSTACK(SUM,AVERAGE) there too. Anywhere else a name is a defined name, Rate.)
  const eta = (k) => { const t = toks[k]; if (t && t.type === 'name' && !params.has(k) && !/^_xl/i.test(t.text) && functionInfo(t.text.toUpperCase())) edits.push({ from: k, to: k, text: `_xleta.${t.text.toUpperCase()}` }); };
  const argsOf = (k) => {
    const end = closing(toks, k + 1);
    const out = [[]];
    let depth = 0;
    for (let j = k + 2; j < end; j++) {
      const x = toks[j];
      if (x.type === 'punct' && (x.text === '(' || x.text === '{')) depth++;
      if (x.type === 'punct' && (x.text === ')' || x.text === '}')) depth--;
      if (depth === 0 && x.type === 'punct' && x.text === ',') out.push([]);
      else out[out.length - 1].push(j);
    }
    return out;
  };
  toks.forEach((t, k) => {
    if (t.type !== 'func') return;
    const name = t.text.toUpperCase().replace(/^(_XL(FN|WS)\.)+/, '');
    const at = { MAP: -1, REDUCE: -1, SCAN: -1, BYROW: 1, BYCOL: 1, MAKEARRAY: 2, GROUPBY: 2, PIVOTBY: 3 }[name];
    if (at === undefined) return;
    const args = argsOf(k);
    const arg = args[at < 0 ? args.length - 1 : at] || [];
    if (arg.length === 1) eta(arg[0]);
    else if (arg.length && toks[arg[0]].type === 'func' && /^(_xlfn\.)?(HSTACK|VSTACK)$/i.test(toks[arg[0]].text)) for (const a of argsOf(arg[0])) if (a.length === 1) eta(a[0]);
  });
  toks.forEach((t, k) => {
    if (t.type === 'ref' && t.spill) edits.push({ from: k, to: k, text: `_xlfn.ANCHORARRAY(${t.text.slice(0, -1)})` });
    if (t.type === 'op' && t.text === '@' && toks[k + 1]) {
      const next = toks[k + 1];
      const end = next.type === 'func' ? closing(toks, k + 2) : next.type === 'punct' && next.text === '(' ? closing(toks, k + 1) : k + 1;
      // (The operand as typed, inner names and # converted too.)
      const inner = dynamicFileForm(typed.slice(next.start, toks[end].end));
      edits.push({ from: k, to: end, text: `_xlfn.SINGLE(${inner})` });
    }
  });
  // (An @ that covers tokens another edit also covers: the @'s operand was converted whole.)
  const kept = edits.filter((e) => !edits.some((o) => o !== e && o.from < e.from && o.to >= e.to));
  return kept.length ? splice(typed, toks, kept) : typed;
}

/** The file's dynamic-array parts as Excel shows them: D1#, @x, and LET's and LAMBDA's names without _xlpm. */
function dynamicShownForm(text) {
  const toks = tokenize(text);
  const edits = [];
  toks.forEach((t, k) => {
    const name = t.type === 'func' ? t.text.toUpperCase().replace(/^(_XL(FN|WS)\.)+/, '') : null;
    if (name === 'ANCHORARRAY' && toks[k + 2]?.type === 'ref' && toks[k + 3]?.text === ')') edits.push({ from: k, to: k + 3, text: `${toks[k + 2].text}#` });
    else if (name === 'SINGLE') {
      const end = closing(toks, k + 1);
      edits.push({ from: k, to: end, text: `@${dynamicShownForm(text.slice(toks[k + 2].start, toks[end - 1].end))}` });
    } else if (t.type === 'name' && /^_xl(pm|eta)\./i.test(t.text)) edits.push({ from: k, to: k, text: t.text.replace(/^_xl(pm|eta)\./i, '') });
  });
  const kept = edits.filter((e) => !edits.some((o) => o !== e && o.from <= e.from && o.to >= e.to && (o.from !== e.from || o.to !== e.to)));
  return kept.length ? splice(text, toks, kept) : text;
}

// Trim references, Excel 365's A1:.E8 (empty rows and columns at the end trimmed), A1.:E8 (at the start) and A1.:.E8
// (both): in the file _xlfn._TRO_TRAILING(A1:E8), _TRO_LEADING and _TRO_ALL (measured).
const END = String.raw`(?:\$?[A-Za-z]{1,3}\$?[0-9]{1,7}|\$?[A-Za-z]{1,3}|\$?[0-9]{1,7})`;
const RE_TRIM = new RegExp(String.raw`((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(${END})(\.?):(\.?)(${END})(?![\w.(!])`, 'g');
const TRO = { '.:': 'LEADING', ':.': 'TRAILING', '.:.': 'ALL' };
/** Outside its strings, a typed formula's trim references in the file's form. */
function trimFileForm(text) {
  if (!/\.:|:\./.test(text)) return text;
  return text.split(/("(?:[^"]|"")*")/).map((part, k) => (k % 2 ? part : part.replace(RE_TRIM, (m, sheet, a, lead, trail, b) => {
    if (!lead && !trail) return m;
    return `_xlfn._TRO_${TRO[`${lead}:${trail}`]}(${sheet || ''}${a}:${b})`;
  }))).join('');
}
/** The file's trim references as typed: _xlfn._TRO_TRAILING(A1:E8) as A1:.E8. */
function trimShownForm(text) {
  return text.replace(/(?:_xlfn\.)?_TRO_(LEADING|TRAILING|ALL)\(([^()]*?):([^():]*)\)/gi, (m, kind, a, b) => {
    const k = kind.toUpperCase();
    return `${a}${k === 'TRAILING' ? '' : '.'}:${k === 'LEADING' ? '' : '.'}${b}`;
  });
}

/** A formula from the file as Excel shows it: without the _xlfn. and _xlws. prefixes of newer functions, and with
 *  Excel 365's D1#, @ and LET's names as typed. */
export function displayFormula(text) {
  try {
    const shown = /_xlfn\.(ANCHORARRAY|SINGLE)|_xlpm\.|_xleta\./i.test(text) ? dynamicShownForm(text) : text;
    // (Trim references last: A1:.E8 isn't a reference to read.)
    return trimShownForm(mapRefs(shown, (t) => (t.type === 'func' && /^_xl(fn|ws)\./i.test(t.text) ? t.text.replace(/^(_xl(fn|ws)\.)+/i, '') : null)));
  } catch {
    return text;
  }
}

/** A typed formula (without "=") in the file's form, as Excel takes an entry: function names, references and TRUE and
 *  FALSE in capitals, newer functions prefixed, missing closing brackets added. Throws a FormulaError when it doesn't
 *  read as a formula. */
export function normalizeFormula(entry) {
  const typed = dynamicFileForm(trimFileForm(entry));
  const toks = tokenize(typed);
  let depth = 0;
  for (const t of toks) {
    if (t.type === 'punct' && t.text === '(') depth++;
    if (t.type === 'func') depth++;
    if (t.type === 'punct' && t.text === ')') depth--;
  }
  // (A function token doesn't include its "(", which comes as its own token: count it once.)
  depth -= toks.filter((t) => t.type === 'func').length;
  let text = mapRefs(typed, (t) => {
    if (t.type === 'func') return fileName(t.text.toUpperCase().replace(/^(_XL(FN|WS)\.)+/, ''));
    if (t.type === 'bool') return t.text.toUpperCase();
    if (t.type === 'ref' && t.area) return t.prefix + areaText(t.area);
    if (t.type === 'err') return t.value;
    return null;
  });
  if (depth > 0) text += ')'.repeat(depth);
  // (A formula that is all @, =@A1:A3, is saved as Excel 2019 wrote it, A1:A3: its implicit intersection is the old one.)
  if (/^_xlfn\.SINGLE\(/i.test(text)) {
    const t = tokenize(text);
    if (closing(t, 1) === t.length - 1) text = text.slice(t[2].start, t[t.length - 1].start);
  }
  return text;
}
