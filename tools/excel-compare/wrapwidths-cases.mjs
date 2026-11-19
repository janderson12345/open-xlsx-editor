// The cases excel-wrapwidths.ps1 measures in Excel, all made up here (no customer text): [{ text, font, size, bold }].
//   node tools/excel-compare/wrapwidths-cases.mjs <cases.json>             the scan cases (test/wrapbreaks-excel.json)
//   node tools/excel-compare/wrapwidths-cases.mjs <cases.json> narrowest   2 to 100 of one character, for -Mode narrowest
//   node tools/excel-compare/wrapwidths-cases.mjs <cases.json> italicruns  2 to 40 of five characters in 14 italic fonts,
//                                                                          for -Mode narrowest (test/wrapitalic-excel.json)
//   node tools/excel-compare/wrapwidths-cases.mjs <cases.json> italic      five zeros in each font's italic and bold
//                                                                          italic, 8 to 36pt, for -Mode narrowest
// The scan cases: sentences of 4 to 17 words drawn from a list with hyphens, digits and punctuation in it (a fixed
// seed, so the cases are the same every run), four in each of 16 fonts, and strings that show one rule each: what a
// hyphen, spaces at either end or a mark does to a line, and which empty lines count.
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';

const FONTS = [['Calibri', 11], ['Calibri', 11, 1], ['Calibri', 10], ['Calibri', 12, 1], ['Calibri', 14, 1], ['Calibri', 9], ['Arial', 10], ['Arial', 8], ['Arial', 12, 1], ['Times New Roman', 12], ['Book Antiqua', 12, 1], ['Tahoma', 10], ['Verdana', 9], ['Calibri', 16], ['Arial', 9, 1], ['Calibri', 8]];
const WORDS = 'We expect to finish the remaining work on this item by end of month, once parts arrive. Total Non-Stock Items 3% handling; 5% flat rate for all tasks Agreement Parts D-K: Part F.2 48.120-35 Report (Filled in by sender) Quote exact wording under review state-of-the-art follow-up 2740.118-64 FREIGHT INSIDE REGION $1,250.00 FY2025 Q3/Q4 e-mail re: A&B a I x - -- "quoted" it\'s (if any) 10/15/2024 ITEM #4 allocation Reallocation Amount (-$) invoices; approval? Yes! No. N/A TBD:'.split(' ');
const RULES = ['aaaa-bbbb', '--------------------', 'aaaa -', 'TENTATIVE -', 'aaaa - bbbb', 'aaaa bbbb ', 'aaaa  bbbb', ' aaaa', 'aaaa   bbbb cccc', 'aaaa/bbbb', 'aaaa,bbbb', 'aaaa.bbbb', 'aaaa(bbbb)', 'aaaa?bbbb', 'aaaa;bbbb', 'aaaa:bbbb', 'aaaa_bbbb', 'aaaa+bbbb', 'aaaa\\bbbb', 'aaaa|bbbb', 'aaaaaaaaaaaaaaaaaaaa bb', 'bb aaaaaaaaaaaaaaaaaaaa', 'bb aaaaaaaaaaaaaaaaaaaa cc', 'a-b-c-d-e-f-g-h', '12/31/2024-01/15/2025', 'well-known, state-of-the-art e-mail', 'TENTATIVE -\nStart Date', 'Unit Cost Source', 'Net Amount (-$)', 'Item # ', 'Cost ', 'x 1-2', 'a - b', 'ab -cd', '-ab -cd -ef', '0000000000', '00000000000000000000', 'iiiiiiiiii', 'WWWWW', 'abc\n', 'abc\n\n', 'abc\n \n', 'abc\n \n  ', 'abc\n  \n', ' \n', 'abc\n \n \n', 'abc\n \nd', 'abc\n\n \n', 'abc \n \n', 'abc\n \n\n'];

/** The scan cases. */
export function scanCases() {
  let seed = 12345;
  const rnd = () => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
  const cases = [];
  for (const [font, size, bold] of FONTS) {
    for (let k = 0; k < 4; k++) {
      const n = 4 + Math.floor(rnd() * 14);
      cases.push({ text: Array.from({ length: n }, () => WORDS[Math.floor(rnd() * WORDS.length)]).join(' '), font, size, bold: !!bold });
    }
  }
  for (const text of RULES) for (const bold of [false, true]) cases.push({ text, font: 'Calibri', size: 11, bold });
  return cases;
}

/** 2 to 100 of one character, in 9 fonts and sizes. */
export function narrowestCases() {
  const cases = [];
  for (const [font, size, bold] of [['Calibri', 11], ['Calibri', 11, 1], ['Calibri', 8], ['Calibri', 14], ['Calibri', 20, 1], ['Arial', 10], ['Arial', 12], ['Times New Roman', 12], ['Verdana', 10, 1], ['Calibri', 7], ['Arial', 6], ['Arial', 7, 1], ['Calibri', 6.5, 1]]) {
    for (const ch of ['0', 'i', 'W', 'e']) for (const n of [2, 3, 5, 8, 12, 15, 20, 30, 40, 60, 80, 100]) cases.push({ text: ch.repeat(n), font, size, bold: !!bold });
  }
  return cases;
}

/** Five zeros in italic and in bold italic, in each font of excel-fontmetrics.ps1 at each half point from 8 to 36pt:
 *  what an italic line takes beyond its characters' widths (charwidths-table.mjs). */
export function italicCases() {
  const fonts = ['Calibri', 'Calibri Light', 'Arial', 'Arial Narrow', 'Arial Black', 'Times New Roman', 'Aptos Narrow', 'Aptos', 'Tahoma', 'Book Antiqua', 'Cambria', 'Verdana', 'Lato', 'Trebuchet MS', 'Courier New', 'Segoe UI', 'Georgia', 'Century Gothic', 'Garamond', 'Consolas', 'Helv', 'Geneva', 'MS Sans Serif', 'Tms Rmn', 'Courier'];
  const cases = [];
  for (const font of fonts) for (let size = 8; size <= 36; size += 0.5) for (const bold of [false, true]) cases.push({ text: '00000', font, size, bold, italic: true });
  return cases;
}

/** 2 to 40 of 0, i, W, f and e in italic, in 14 fonts and sizes: real italics, slanted fonts (Tahoma, MS Sans Serif),
 *  bold italics and Small Fonts (Arial 7). */
export function italicRunCases() {
  const cases = [];
  for (const [font, size, bold] of [['Calibri', 11], ['Calibri', 11, 1], ['Tahoma', 10], ['Tahoma', 10, 1], ['Arial', 7], ['Arial', 7, 1], ['Book Antiqua', 10], ['Book Antiqua', 12, 1], ['Arial', 10], ['Times New Roman', 12], ['Verdana', 9], ['Calibri', 16], ['MS Sans Serif', 10], ['Courier New', 10]]) {
    for (const ch of ['0', 'i', 'W', 'f', 'e']) for (const n of [2, 3, 5, 10, 20, 40]) cases.push({ text: ch.repeat(n), font, size, bold: !!bold, italic: true });
  }
  return cases;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const cases = process.argv[3] === 'narrowest' ? narrowestCases() : process.argv[3] === 'italic' ? italicCases() : process.argv[3] === 'italicruns' ? italicRunCases() : scanCases();
  fs.writeFileSync(process.argv[2], JSON.stringify(cases));
  console.log(`${cases.length} cases`);
}
