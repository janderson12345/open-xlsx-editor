// REGEXTEST, REGEXEXTRACT and REGEXREPLACE over the PCRE2 features JavaScript's regular expressions haven't got, for
// measuring what Excel does with each (docs/CALC.md): \K, recursion, conditionals, \G, \X, \C, options partway
// through a pattern, caseless \p{Lu} and [[:upper:]], backtracking verbs, lookbehind of varying length, branch reset
// and the rest, with the errors Excel gives for what it refuses. Excel works them out with excel-formulas.ps1.
// usage: node tools/excel-compare/regex-cases.mjs <formulas.json>
import fs from 'node:fs';

const q = (s) => `"${s.replace(/"/g, '""')}"`;
const out = [];
const X = (s, p, mode = 0, cs) => out.push(`=REGEXEXTRACT(${q(s)},${q(p)},${mode}${cs === undefined ? '' : `,${cs}`})`);
const T = (s, p, cs) => out.push(`=REGEXTEST(${q(s)},${q(p)}${cs === undefined ? '' : `,${cs}`})`);
const R = (s, p, r, occ, cs) => out.push(`=REGEXREPLACE(${q(s)},${q(p)},${q(r)}${occ === undefined ? '' : `,${occ}`}${cs === undefined ? '' : `,,${cs}`})`);
const XA = (s, p) => out.push(`=ARRAYTOTEXT(REGEXEXTRACT(${q(s)},${q(p)},1))`, `=ARRAYTOTEXT(REGEXEXTRACT(${q(s)},${q(p)},2))`);

// \K
X('foobar', 'foo\\Kbar'); X('price: 42', 'price: \\K\\d+'); XA('a1b2c3', '[a-z]\\K\\d'); R('aaa', 'a\\Ka', 'X');
R('foobar foobaz', 'foo\\K(bar|baz)', '[$1]'); X('abc', 'a\\Kb\\Kc'); X('abc', '(?=ab\\K)'); X('abc', '(?<=a\\K)b');
X('abab', '(a\\K|b)b'); XA('abab', '(a\\K|b)b'); R('ab ab', '(?:a\\K)b', '#'); T('xy', 'x\\Ky'); X('xy', 'x\\K');
R('a1a2', 'a\\K', '-'); X('aab', 'a+\\Kb'); R('abc', 'b\\K', '#'); X('ab', '(a)\\Kb', 2);
// Recursion
X('x(a(b)c)y', '\\((?:[^()]|(?R))*\\)'); X('((()))', '\\((?R)?\\)'); X('(()', '\\((?R)?\\)'); X('abcba', '^((.)(?:(?1)|.?)\\2)$');
X('racecar!', '\\b(\\w)(?:(?R)|\\w?)\\1\\b'); X('{a{b}c}', '(\\{(?:[^{}]|(?1))*\\})'); X('{a{b}c}', '(?<br>\\{(?:[^{}]|(?&br))*\\})');
X('{a{b}c}', '(?P<br>\\{(?:[^{}]|(?P>br))*\\})'); X('{a{b}c}', '(\\{(?:[^{}]|\\g<1>)*\\})'); X('{a{b}c}', "(\\{(?:[^{}]|\\g'1')*\\})");
X('aab', '(a)(?-1)b'); X('aab', '(?+1)(a)b'); X('abab', '(ab)(?1)'); X('1+2*(3+4)', '(\\d|\\((?1)(?:[+*](?1))*\\))'); XA('(a)(b(c))', '\\((?:[^()]|(?R))*\\)');
X('aaa', '(?R)'); X('a', 'a(?R)?'); X('aaaa', '(a(?1)?)'); X('((x))', '\\(((?R)|x)\\)', 2); X('(((((((((((x)))))))))))', '\\((?R)\\)|x');
// Conditionals
X('<a>', '(<)?a(?(1)>|)'); X('a>', '(<)?a(?(1)>|)'); X('a', '(<)?a(?(1)>)'); X('ab', '(?<o>a)?(?(<o>)b|c)'); X('c', "(?<o>a)?(?('o')b|c)");
X('ab', '(a)?(?(1)b|c)'); X('xc', '(a)?x(?(1)b|c)'); X('ab', '(?(?=a)ab|cd)'); X('cd', '(?(?=a)ab|cd)'); X('cd', '(?(?!a)cd|ab)');
X('ab', '(?(?<=x)ab|a)'); X('12-34', '(?(DEFINE)(?<n>\\d\\d))(?&n)-(?&n)'); X('aa', '(?(R)a|a(?R)?)'); X('ab', '(a)(?(1)b|c|d)');
X('ab', '(?(2)a)'); X('ab', '(?(1a)b)'); X('ab', '(a)?(?(+1)b|c)(b)'); X('abb', '(a)(?(-1)b)b');
// \G
R('aaab', '\\Ga', 'X'); XA('aaba', '\\Ga'); X('ba', '\\Ga'); T('aa', '\\Ga'); R('1,2,3;4', '\\G(\\d),?', '<$1>'); R('abc', '\\G', '-');
// \X and \C
X('éx', '\\X'); XA('áb', '\\X'); X('\u{1F468}‍\u{1F469}‍\u{1F467}!', '^\\X'); X('\u{1F1FA}\u{1F1F8}x', '\\X');
X('\r\nx', '\\X'); X('abc', 'a\\Cc'); T('\u{1F600}', '^\\C\\C$'); X('xy', '\\X\\X');
// Options partway through
X('aB', 'a(?i)b'); X('AB', 'a(?i)b'); T('aB', 'a(?i)b'); X('aBc', 'a(?i:b)c'); X('aBC', 'a(?i:b)c'); X('AbC', '(?i)a(?-i)b(?i)c');
X('Ab', '(?i)a(?-i:B)'); X('Abc', '(?i)a(?-i:b)C'); X('a\nb', 'a(?s).b'); X('a\nb', 'a(?s:.)b'); X('x\nab', 'x\\n(?m)^ab');
X('ab', 'a(?x) b'); X('ab', '(?x:a b)'); X('aB', '(a(?i)b)'); X('Ab', '(a(?i)b)'); X('aBa', '(?:a(?i)b)a'); X('aBA', '(?:a(?i)b)a');
X('ab|AB', 'a(?i)b|c'); X('C', 'a(?i)b|c'); X('Ab', '(?i:a|B)b'); X('aB', '(?^)a(?i)b'); X('aB', 'a(?^i)b'); X('ab', '(?n)(a)(b)', 2);
X('aaa', '(?U)a+'); X('aaa', '(?U)a+?'); X('ab', '(?J)(?<x>a)(?<x>b)'); X('a b', '(?xx)a [ ] b'); X('AB', '(?i)a(?^)b');
X('ab', 'a(?i)'); X('Ba', '(?i)b(?-i)a', 0, 1); X('BA', '(?-i)ba', 0, 1); X('bA', 'b(?-i)a', 0, 1); X('BA', 'b(?-i:A)', 0, 1);
X('é É', '(?i)É'); R('aAaA', 'a(?i)a', '#');
// Caseless classes
T('a', '\\p{Lu}', 1); T('A', '\\p{Ll}', 1); T('a', '[[:upper:]]', 1); T('A', '[[:lower:]]', 1); T('a', '[A-Z]', 1); T('a', '\\p{Lt}', 1);
T('ǅ', '\\p{Lu}', 1); T('a', '[\\p{Lu}]', 1); T('a', '\\P{Ll}', 1); T('a', '[^\\p{Ll}]', 1); T('a', '(?i)\\p{Lu}'); T('A', '\\p{Lu}', 1);
T('ß', '\\p{Lu}', 1); T('a', '[[:^lower:]]', 1); T('a', '\\p{L&}', 1); T('a', '\\p{Lu}|x', 1); X('aB', '\\p{Lu}+', 0, 1);
// Verbs
X('aaabbb', 'a+(*SKIP)(*FAIL)|b+'); XA('aaabbb', 'a+(*SKIP)(*F)|b+'); X('ab', 'a(*PRUNE)b'); X('aac', 'a+(*COMMIT)b|c'); X('ab', 'a(*ACCEPT)b');
X('ab', '(*MARK:x)a'); X('ab', 'a(*THEN)x|ab'); X('ab', '(*UTF)a'); X('ab', '(*UCP)\\w'); X('a\r\nb', '(*CR)a$'); X('a\nb', '(*LF)a$');
X('ab', '(*NO_AUTO_POSSESS)a+b'); X('aaaa', '(*LIMIT_MATCH=10)a+'); X('ab', '(*FAIL)|b'); X('ab', '(*ACCEPT)');
// Lookbehind, branch reset, names, other syntax
X('aab', '(?<=a+)b'); X('aab', '(?<=a{1,3})b'); X('xab', '(?<=xa|b)b'); X('xaab', '(?<=x(?:a|aa))b'); X('ab', '(?<=\\Ka)b');
X('xb', '(?|(a)|(x))b', 2); X('b', '(?|(a)|(b))', 2); X('ab', '(?<n>a)(?<n>b)'); X('ab', '(?<n>a)|(?<n>b)'); X('A', '\\N{U+41}');
X('ab', '\\u0061'); X('ab', '\\U'); X('ab', '\\L'); X('ab', 'a\\E'); X('a.b', '\\Qa.b'); X('ab', '[\\K]'); X('ab', '(?C)a');
X('ab', '(?C1)a'); X('ab', '\\x{110000}'); X('aaa', 'a{2,1}'); X('ab', '(*'); X('ab', '(?'); X('ab', '(?<=a)'); X('ab', 'a{65536}');
X('Ωx', '\\p{Greek}'); X('a1_', '\\p{Xwd}+'); X('a1', '\\p{Xan}+'); X(' x', '\\p{Xsp}'); X('$', '\\p{Xuc}'); X('a', '\\p{Latin}');
X('日本', '\\p{Han}+'); X('ab', '\\p{Zzzz}'); X('a', '\\p{Lc}'); X('a', '\\p{Script=Latin}'); X('a', '\\p{sc=Latin}'); X('a', '\\p{IsLatin}');
X('a', '\\p{InBasicLatin}'); X('ab', '(?<a>.)\\k{a}'); X('aa', '(.)\\g{-1}'); X('a\u{85}b', 'a\\Rb'); X('a', '\\ca'); X('a', '\\o{141}');
X('a', '\\141'); X('ab', '\\8'); X('ab', '[[:word:]]+'); X('ab', '[[.a.]]'); X('ab', '[[=a=]]'); X('ab', '[[:foo:]]'); X('ab', '[\\d-z]+');
X('a-z', '[a-\\d]+'); X('ab', '(?#x)a'); X('aaa', 'a++'); X('aaa', '(?>a+)a'); X('aaab', 'a*+b'); X('aaa', 'a{2}+'); X('ab', 'a{,1}b');
X('a{,1}b', 'a{,1}b'); X('a{1,b', 'a{1,b'); X('ab', '\\Aab\\z'); X('ab\n', 'ab\\Z'); T('ab\n', 'ab\\z'); X('a\nb', '^b'); X('a\nb', '(?m)^b');
// Catastrophic backtracking: what Excel does at PCRE2's match limit
T(`${'a'.repeat(30)}b`, '^(a+)+$'); T(`${'a'.repeat(40)}!`, '^(a|a)*$'); T(`${'a'.repeat(28)}`, '^(a+)+$'); X(`${'a'.repeat(25)}x`, '(a*)*b');
R(`${'a'.repeat(30)}b`, '(a+)+$', '#'); T(`${'x'.repeat(200)}`, '(x+x+)+y'); T(`${'ab'.repeat(20)}c`, '^(ab|a|b)*$');
// \G next to empty matches: where each search after one starts.
for (const s of ['abc', 'xxa', 'céB', 'aéb']) {
  for (const p of ['\\G', '(?!\\G)', '(?<=(?!\\G))', '(?=\\G)', '(?<!\\G)', '\\Ga|b', 'a|\\G', 'x*', '\\Gx*', '(?!\\G)x*', '\\G(?!a)', '(?<=\\G)', '(?<=\\G)a', '(?<=a)']) {
    out.push(`=REGEXREPLACE(${q(s)},${q(p)},"<$0>")`, `=ARRAYTOTEXT(REGEXEXTRACT(${q(s)},${q(p)},1))`);
  }
}

// Held out: 600 random patterns over these features on random texts, each as every match, the groups of the first,
// and a replacement of every match.
let seed = 31337;
const rnd = (n) => { seed = (seed * 1103515245 + 12345) % 2147483648; return Math.floor(seed / 65536) % n; };
const ATOMS = ['a', 'b', 'c', 'A', '.', '\\w', '\\d', '\\s', '[ab]', '[^a]', '[[:upper:]]', '\\b', '^', '$', '\\K', '\\G', '\\1', '(?1)', '\\X', '\\p{Lu}'];
function rpat(depth) {
  const n = 1 + rnd(3);
  let s = '';
  for (let k = 0; k < n; k++) {
    let a;
    const r = rnd(10);
    if (depth < 2 && r < 3) {
      const open = ['(', '(', '(?:', '(?>', '(?=', '(?!', '(?<=', '(?i:', '(?<n>', '(?|'][rnd(10)];
      a = `${open}${rpat(depth + 1)}${rnd(3) ? '' : `|${rpat(depth + 1)}`})`;
    } else if (depth < 2 && r === 3) a = `(?(1)${rpat(depth + 1)}|${rpat(depth + 1)})`;
    else if (r === 4) a = ['(?i)', '(?-i)', '(?s)', '(?m)', '(*F)|', '(?R)?'][rnd(6)];
    else a = ATOMS[rnd(ATOMS.length)];
    if (!/^[(^$\\][bKGXbmsi(*]?$/.test(a) || a.startsWith('(')) a += ['', '', '', '*', '+', '?', '*?', '+?', '++', '{1,2}', '{2}'][rnd(11)];
    s += a;
  }
  return s;
}
const TEXT = ['a', 'b', 'c', 'A', 'B', '1', ' ', '\n', 'ab', 'é'];
for (let k = 0; k < 200; k++) {
  const p = rpat(0);
  let s = '';
  for (let j = 0, m = 2 + rnd(8); j < m; j++) s += TEXT[rnd(TEXT.length)];
  out.push(`=ARRAYTOTEXT(REGEXEXTRACT(${q(s)},${q(p)},1))`, `=ARRAYTOTEXT(REGEXEXTRACT(${q(s)},${q(p)},2))`, `=REGEXREPLACE(${q(s)},${q(p)},"<$0>")`);
}
fs.writeFileSync(process.argv[2], JSON.stringify(out));
console.log(`${out.length} formulas`);
