// Text-function formulas over emoji and other characters past U+FFFF, all built with UNICHAR so no encoding stands
// between, for measuring how Excel counts them (docs/CALC.md): a fixed set over ten strings (an emoji, a ZWJ family,
// a variation selector, a skin tone, a flag, a combining accent, a CJK Extension B character), some edge cases, and
// 600 random ones over other characters as a held-out check. The same seed gives the same cases. Excel works them out
// with excel-formulas.ps1, in a workbook without the function compatibility version and in one with setVersion 2;
// emoji-excel.mjs puts both answers into test/emoji-excel.json.
// usage: node tools/excel-compare/emoji-cases.mjs <formulas.json>
import fs from 'node:fs';

const U = (...cps) => cps.map((c) => `UNICHAR(${c})`).join('&');
const E = U(128512);
const strs = [E, `"a"&${E}&"b"`, U(128512, 128513), U(128104, 8205, 128105, 8205, 128103), `"x"&${U(10084, 65039)}&"y"`,
  `${U(128075, 127995)}&"z"`, U(127482, 127480), `"e"&${U(769)}&"f"`, `"q"&${U(131072)}&"r"`,
  `"ab"&${E}&"cd"&${U(128104, 8205, 128105)}&"e"`];
const out = [];
for (const s of strs) {
  out.push(`=LEN(${s})`, `=LENB(${s})`, `=UNICODE(${s})`, `=CODE(${s})`);
  for (const n of [0, 1, 2, 3, 4, 5]) out.push(`=LEFT(${s},${n})`, `=RIGHT(${s},${n})`, `=LEFTB(${s},${n})`, `=RIGHTB(${s},${n})`);
  for (const st of [1, 2, 3, 4, 5]) for (const n of [1, 2, 3]) out.push(`=MID(${s},${st},${n})`);
  for (const st of [1, 2, 3, 4]) out.push(`=MIDB(${s},${st},2)`, `=REPLACE(${s},${st},1,"#")`, `=REPLACEB(${s},${st},2,"#")`);
  for (const st of [1, 2, 3, 4]) out.push(`=UNICODE(MID(${s},${st},1))`);
  for (const w of ['"a"', '"b"', '"e"', '"y"', '"z"', '"r"', '"f"', '"d"']) {
    out.push(`=FIND(${w},${s})`, `=SEARCH(${w},${s})`, `=FINDB(${w},${s})`, `=SEARCHB(${w},${s})`, `=FIND(${w},${s},2)`, `=SEARCH(${w},${s},3)`);
  }
  out.push(`=FIND(${E},${s})`, `=SEARCH(${E},${s})`, `=SEARCH("?",${s})`, `=SEARCH("??",${s})`,
    `=SEARCH("?b",${s})`, `=SEARCH("a?b",${s})`, `=SEARCH("*b",${s})`, `=SEARCH("?",${s},2)`, `=FIND(${s},${s})`,
    `=FIND("",${s},3)`, `=SEARCH("",${s},4)`, `=FIND("a",${s},LEN(${s})+1)`, `=FIND("",${s},LEN(${s})+1)`,
    `=FIND("",${s},LEN(${s})+2)`, `=SUBSTITUTE(${s},${E},"#")`, `=EXACT(${s},${s})`, `=UPPER(${s})`,
    `=TEXTBEFORE(${s},"b")`, `=TEXTAFTER(${s},"a")`, `=REPT(${s},2)`, `=LEN(TRIM(${s}))`, `=LEN(CLEAN(${s}))`,
    `=LEN(PROPER(${s}))`, `=LEN(TEXTJOIN("",,${s},${s}))`, `=TEXTSPLIT(${s},"b")`, `=LEN(VALUETOTEXT(${s}))`);
}
// (Edges: where FIND and SEARCH may start, half an emoji alone, letters past U+FFFF in UPPER and LOWER.)
const AE = `"a"&${E}&"b"`;
const D = U(66560);
const d = U(66600);
out.push('=SEARCH("c","abc",4)', '=SEARCH("c","abc",3)', '=SEARCH("","abc",3)', '=SEARCH("","abc",4)', '=FIND("","abc",4)',
  '=SEARCH("","")', '=FIND("","")', '=SEARCH("","",1)', '=SEARCH("*","abc",4)', '=SEARCH("*","")',
  `=SEARCH("b",${AE},4)`, `=SEARCH("b",${AE},3)`, `=SEARCH("*b",${E}&"b")`, `=SEARCH("?"&UNICHAR(56832),"a"&${E})`,
  `=SEARCH(${E},${AE},2)`, `=FIND(${E},${AE},3)`, `=SEARCH(${E},${AE},3)`, `=UNICODE(MIDB(${E}&"a",2,2))`,
  `=UNICODE(MIDB(${E}&"a",1,1))`, `=UNICODE(MIDB(${E}&"a",2,1))`, `=LEN(MIDB(${E}&"a",2,2))`, `=LEN(MIDB(${E}&"a",1,1))`,
  `=MID(MIDB("x"&${E},3,1)&"y",1,1)`, `=FIND("y",MIDB("x"&${E},3,1)&"y")`, `=FIND("y",MIDB("x"&${E},2,1)&"y")`,
  `=LEN(MIDB("x"&${E},2,1)&"y")`, `=LEFT(${U(128104, 8205, 128105)},2)`, `=RIGHT(${U(128104, 8205, 128105)},2)`,
  `=LEN(LEFT(${U(128104, 8205, 128105)},2))`, `=SEARCH("~?","a?"&${E})`, `=SEARCH("a*"&${E},"xa"&${E})`,
  `=REPLACE(${AE},2,0,"#")`, `=REPLACE(${AE},5,1,"#")`, `=MID(${AE},3,0)`, `=MID(${AE},9,1)`, `=SEARCH("B",${AE})`,
  `=SEARCH(${D},${d})`, `=FIND(${D},${d})`, `=LEN(UPPER(${d}))`, `=UNICODE(UPPER(${d}))`, `=UNICODE(LOWER(${D}))`,
  `=EXACT(${D},${d})`);
// (Held out: random strings of other characters, random functions and positions.)
let seed = 20261116;
const rnd = (n) => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed % n; };
const pieces = ['"a"', '"B"', '"?"', '"*"', '" "', U(127881), U(9731, 65039), `"1"&${U(65039, 8419)}`,
  U(127988, 917607, 917602, 917631), U(119808), U(128105, 127997, 8205, 128187), U(65533), U(1488), U(26085), U(194560),
  `MIDB(${E},1,1)`, `MIDB(${E},2,1)`];
const str = () => { const n = 1 + rnd(4); const p = []; for (let i = 0; i < n; i++) p.push(pieces[rnd(pieces.length)]); return p.join('&'); };
const needles = ['"a"', '"b"', '"?"', '"~?"', '"*a"', '"?a"', '"a?"', '""', U(127881), U(119808), U(65039), U(8205)];
for (let i = 0; i < 600; i++) {
  const s = str(); const k = 1 + rnd(6); const m = rnd(4);
  switch (rnd(9)) {
    case 0: out.push(`=LEN(${s})`); break;
    case 1: out.push(`=UNICODE(MID(${s},${k},1))`); break;
    case 2: out.push(`=LEN(MID(${s},${k},${m}))`); break;
    case 3: out.push(`=FIND(${needles[rnd(needles.length)]},${s},${k})`); break;
    case 4: out.push(`=SEARCH(${needles[rnd(needles.length)]},${s},${k})`); break;
    case 5: out.push(`=REPLACE(${s},${k},${m},"#")`); break;
    case 6: out.push(`=LEFT(${s},${k})`); break;
    case 7: out.push(`=RIGHT(${s},${k})`); break;
    default: out.push(`=MID(${s},${k},${m})`);
  }
}
fs.writeFileSync(process.argv[2], JSON.stringify(out));
console.log(`${out.length} formulas`);
