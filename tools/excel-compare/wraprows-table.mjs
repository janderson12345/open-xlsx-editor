// A font's line and descent as Excel has them, for fontmetrics-table.mjs (src/ooxml/excelfonts.js).
//   node tools/excel-compare/wraprows-table.mjs cases <cases.json>
//     the cases for excel-wraprows.ps1: 6 and 7 lines of each font of excel-fontmetrics.ps1, regular, bold, italic
//     and bold italic, at each half point from 6 to 36pt. Run it with -NormalFont Calibri -NormalSize 36.
//   node tools/excel-compare/wraprows-table.mjs check-cases <cases.json>
//     1, 2, 3 and 5 lines of six fonts from 6 to 16pt, to run under several Normal fonts (Calibri 8, 11 and 14, Arial
//     10): the rule itself, apart from the table.
//   node tools/excel-compare/wraprows-table.mjs check <excel-wraprows.json>...
//     writes test/wraprows-excel.json of those runs: { excel, sets: [{ normal, standard, rows: [[font, size, lines,
//     pt], ...] }] }.
//   node tools/excel-compare/wraprows-table.mjs table <excel-wraprows.json>
//     writes test/fontlines-excel.json: { excel, normal, fonts: { name: { regular|bold|italic|boldItalic:
//     [[line px, descent px], ...] } } }, by half point from 6pt.
// Excel's row of n lines of a font is n x line + (the row's deepest descent - the font's), measured (docs/LAYOUT.md):
// so 7 lines less 6 is the line, and what 6 lines have over 6 x line is the Normal font's descent less the font's.
// Calibri 36's own descent is taken from Calibri 11's being 4px (Windows' tmDescent for both, 13 and 4, and what
// Excel's rows of 11pt text under larger and smaller Normal fonts give).
import fs from 'node:fs';

const FONTS = ['Calibri', 'Calibri Light', 'Arial', 'Arial Narrow', 'Arial Black', 'Times New Roman', 'Aptos Narrow', 'Aptos', 'Tahoma', 'Book Antiqua', 'Cambria', 'Verdana', 'Lato', 'Trebuchet MS', 'Courier New', 'Segoe UI', 'Georgia', 'Century Gothic', 'Garamond', 'Consolas', 'Helv', 'Geneva', 'MS Sans Serif', 'Tms Rmn', 'Courier'];
const STYLES = { regular: [false, false], bold: [true, false], italic: [false, true], boldItalic: [true, true] };
const styleOf = (c) => (c.bold ? (c.italic ? 'boldItalic' : 'bold') : c.italic ? 'italic' : 'regular');
const [mode, file, ...more] = process.argv.slice(2);
if (mode === 'cases') {
  const cases = [];
  for (const font of FONTS) for (let size = 6; size <= 36; size += 0.5) for (const [bold, italic] of Object.values(STYLES)) for (const lines of [6, 7]) cases.push({ font, size, bold, italic, lines });
  fs.writeFileSync(file, JSON.stringify(cases));
  console.log(`${cases.length} cases`);
} else if (mode === 'check-cases') {
  const cases = [];
  for (const font of ['Calibri', 'Arial', 'Verdana', 'Times New Roman', 'Tahoma', 'Book Antiqua']) for (let size = 6; size <= 16; size += 1) for (const lines of [1, 2, 3, 5]) cases.push({ font, size, bold: false, italic: false, lines });
  fs.writeFileSync(file, JSON.stringify(cases));
  console.log(`${cases.length} cases`);
} else if (mode === 'check') {
  const sets = [file, ...more].map((f) => { const d = JSON.parse(fs.readFileSync(f, 'utf8').trimStart()); return { excel: d.excel, normal: d.normal, standard: d.standard, rows: d.cases.map((c) => [c.font, c.size, c.lines, c.pt]) }; });
  fs.writeFileSync('test/wraprows-excel.json', `${JSON.stringify({ excel: sets[0].excel, sets: sets.map(({ normal, standard, rows }) => ({ normal, standard, rows })) })}\n`);
  console.log(`test/wraprows-excel.json: ${sets.length} Normal fonts`);
} else if (mode === 'table') {
  const d = JSON.parse(fs.readFileSync(file, 'utf8').trimStart());
  const px = new Map();
  for (const c of d.cases) px.set(`${c.font}|${c.size}|${styleOf(c)}|${c.lines}`, Math.round(c.pt / 0.75));
  const measure = (font, size, style) => {
    const six = px.get(`${font}|${size}|${style}|6`);
    const seven = px.get(`${font}|${size}|${style}|7`);
    // (546px is Excel's tallest row: a row stopped there says nothing.)
    if (six == null || seven == null || seven >= 546) throw new Error(`${font} ${size} ${style}: not measured`);
    return [seven - six, six - 6 * (seven - six)];
  };
  const deepest = measure('Calibri', 11, 'regular')[1] + 4;
  const fonts = {};
  for (const font of FONTS) {
    fonts[font] = {};
    for (const style of Object.keys(STYLES)) {
      fonts[font][style] = [];
      for (let size = 6; size <= 36; size += 0.5) { const [line, over] = measure(font, size, style); fonts[font][style].push([line, deepest - over]); }
    }
  }
  fs.writeFileSync('test/fontlines-excel.json', `${JSON.stringify({ excel: d.excel, normal: d.normal, fonts })}\n`);
  console.log(`test/fontlines-excel.json: ${FONTS.length} fonts, the Normal font's descent ${deepest}px`);
} else {
  console.error('usage: node tools/excel-compare/wraprows-table.mjs cases|check-cases <cases.json> | table <excel-wraprows.json> | check <excel-wraprows.json>...');
  process.exit(1);
}
