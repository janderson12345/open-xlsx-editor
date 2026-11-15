// Browser check of rich text while a cell is edited: the text drawn formatted as it is typed and as Ctrl+B, I, U and 5
// format part of it (celleditor.js, richview.js), typed text taking the font before the caret (richmarks.js, as
// Excel's own insert measured over COM: test/richtype-excel.json), the caret and a click placed by the drawn text,
// undo, Esc and an IME composition, a formula staying plain, and what was shown being what is entered and saved.
// Run after `npm run build`: node test/editing-rich.mjs [outDir]
import fs from 'node:fs';
import { editingPage } from './editing-kit.mjs';
import { openFixture } from './browser.mjs';

const { page, check, finish, openFixtureFrom } = await editingPage();
await openFixture(page, 'basic.xlsx');

/** The editor's drawn runs: [text, weight, style, decoration, colour, size, family] each, or null when not drawn. */
const drawn = () => page.evaluate(() => {
  const view = document.querySelector('.oxe-edit-view');
  if (!view?.isConnected) return null;
  return [...view.querySelectorAll('.oxe-edit-runs > span')].map((x) => {
    const st = getComputedStyle(x);
    return [x.textContent, st.fontWeight, st.fontStyle, st.textDecorationLine, st.color, st.fontSize, st.fontFamily.split(',')[0].replace(/"/g, '')];
  });
});
/** A cell's runs in the workbook, [text, bold, italic, underline, strike] each. */
const savedRuns = (r, c) => page.evaluate(([r, c]) => (window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(r)?.get(c)?.runs || null)
  ?.map((x) => [x.text, !!x.font?.b, !!x.font?.i, !!x.font?.u, !!x.font?.strike]), [r, c]);
/** The drawn runs as the same shape. */
const asRuns = (d) => d.map(([t, w, s, dec]) => [t, Number(w) >= 600, s === 'italic', /underline/.test(dec), /line-through/.test(dec)]);
const boxText = () => page.evaluate(() => document.querySelector('.oxe-editor:not(.oxe-edit-view)')?.value ?? null);

// Typed in E30, F2 to move the caret, the last word selected and Ctrl+B: drawn bold at once, before Enter.
await page.evaluate(() => window.oxe.select('E30'));
await page.focus('.oxe-grid');
await page.keyboard.type('Plain bold');
check(await drawn() === null && !(await page.evaluate(() => document.querySelector('.oxe-editor').classList.contains('oxe-editor-rich'))),
  'plain text typed: the text box’s own text, nothing drawn over it');
await page.keyboard.press('F2');
for (let k = 0; k < 4; k++) await page.keyboard.press('Shift+ArrowLeft');
await page.keyboard.press('Control+b');
let d = await drawn();
check(d?.length === 2 && d[0][0] === 'Plain ' && d[0][1] === '400' && d[1][0] === 'bold' && d[1][1] === '700', 'Ctrl+B on the selected word: drawn bold while still editing', JSON.stringify(d));
const marks = await page.evaluate(() => ({ sel: document.querySelectorAll('.oxe-edit-sel').length, rich: document.querySelector('.oxe-editor.oxe-editor-rich') && getComputedStyle(document.querySelector('.oxe-editor.oxe-editor-rich')).color }));
check(marks.sel >= 1 && marks.rich === 'rgba(0, 0, 0, 0)', 'the selection drawn over the drawn text, the box’s own text hidden', JSON.stringify(marks));

// Ctrl+I and Ctrl+U on the same word, then Ctrl+5 on "Pla": each drawn as toggled.
await page.keyboard.press('Control+i');
await page.keyboard.press('Control+u');
await page.keyboard.press('Home');
for (let k = 0; k < 3; k++) await page.keyboard.press('Shift+ArrowRight');
await page.keyboard.press('Control+5');
d = await drawn();
check(JSON.stringify(asRuns(d)) === JSON.stringify([['Pla', false, false, false, true], ['in ', false, false, false, false], ['bold', true, true, true, false]]),
  'Ctrl+I, Ctrl+U on the word and Ctrl+5 on "Pla": drawn italic, underlined and struck where toggled', JSON.stringify(d));

// Typed after the bold run (End): bold, the font before the caret; typed at the start (Home): the first character's.
await page.keyboard.press('End');
await page.keyboard.type('er');
await page.keyboard.press('Home');
await page.keyboard.type('A ');
d = await drawn();
check(JSON.stringify(asRuns(d)) === JSON.stringify([['A Pla', false, false, false, true], ['in ', false, false, false, false], ['bolder', true, true, true, false]]),
  'typed at the end of the bold run: bold; typed at the start: the first character’s font (struck)', JSON.stringify(d));
// (The caret drawn just after "A ", by the drawn characters.)
const caret = await page.evaluate(() => {
  const c = document.querySelector('.oxe-edit-caret')?.getBoundingClientRect();
  const span = document.querySelector('.oxe-edit-runs > span');
  const r = document.createRange();
  r.setStart(span.firstChild, 2);
  r.setEnd(span.firstChild, 2);
  return c ? { caret: c.left, text: r.getBoundingClientRect().left } : null;
});
check(caret && Math.abs(caret.caret - caret.text) < 1.5, 'the caret drawn where the drawn text puts it', JSON.stringify(caret));

// Typed inside the struck run, deleted inside the bold one, a line break (Alt+Enter) after "in": each stays in place.
await page.keyboard.press('ArrowRight');
await page.keyboard.type('x');
await page.keyboard.press('End');
await page.keyboard.press('Backspace');
await page.keyboard.press('Backspace');
d = await drawn();
check(JSON.stringify(asRuns(d)) === JSON.stringify([['A Pxla', false, false, false, true], ['in ', false, false, false, false], ['bold', true, true, true, false]]),
  'typed inside the struck run: struck; Backspace inside the bold run: the rest stays bold', JSON.stringify(d));

// Ctrl+Z in the editor: the text box's own undo, the formatting back with the text.
await page.keyboard.press('Control+z');
d = await drawn();
check(await boxText() === 'A Pxlain bolde' && d?.[2]?.[0] === 'bolde' && d[2][1] === '700', 'Ctrl+Z in the editor: the deleted letter back, bold as it was', JSON.stringify([await boxText(), d]));

// A click in the drawn text: the caret put by the drawn characters (between "b" and "o" of the bold run).
const target = await page.evaluate(() => {
  const span = document.querySelectorAll('.oxe-edit-runs > span')[2];
  const r = document.createRange();
  r.setStart(span.firstChild, 1);
  r.setEnd(span.firstChild, 1);
  const b = r.getBoundingClientRect();
  return { x: b.left + 1, y: b.top + b.height / 2 };
});
await page.mouse.click(target.x, target.y);
const clicked = await page.evaluate(() => document.querySelector('.oxe-editor:not(.oxe-edit-view)').selectionStart);
check(clicked === 'A Pxlain b'.length, 'a click in the drawn text: the caret where it was clicked', String(clicked));

// At 150% zoom too: a click between "o" and "l" of the bold run puts the caret there, drawn at that point.
await page.evaluate(() => window.oxe.command('zoom', 150));
const zoomed = await page.evaluate(() => {
  const span = document.querySelectorAll('.oxe-edit-runs > span')[2];
  const r = document.createRange();
  r.setStart(span.firstChild, 2);
  r.setEnd(span.firstChild, 2);
  const b = r.getBoundingClientRect();
  return { x: b.left + 1, y: b.top + b.height / 2 };
});
await page.mouse.click(zoomed.x, zoomed.y);
const zoomCaret = await page.evaluate(() => ({ at: document.querySelector('.oxe-editor:not(.oxe-edit-view)').selectionStart, x: document.querySelector('.oxe-edit-caret')?.getBoundingClientRect().left }));
await page.evaluate(() => window.oxe.command('zoom', 100));
check(zoomCaret.at === 'A Pxlain bo'.length && Math.abs(zoomCaret.x - (zoomed.x - 1)) < 2, 'at 150% zoom: a click puts the caret there and it is drawn there', JSON.stringify({ zoomCaret, zoomed }));

// Enter: the cell's runs are what was drawn, and the cell draws them so.
const shownBefore = asRuns(await drawn());
await page.keyboard.press('Enter');
const saved = await savedRuns(29, 4);
const cellDrawn = await page.evaluate(() => [...document.querySelectorAll('.oxe-row[aria-rowindex="31"] .oxe-cell[aria-colindex="6"] .oxe-runs > span')].map((x) => [x.textContent, getComputedStyle(x).fontWeight]));
check(JSON.stringify(saved) === JSON.stringify(shownBefore) && cellDrawn.length === 3 && cellDrawn[2][1] === '700',
  'entered: the cell’s runs are the ones drawn while editing, and the cell draws them', JSON.stringify({ saved, shownBefore, cellDrawn }));

// F2 on it again: drawn formatted from the start; Esc gives up a change, the runs kept.
await page.evaluate(() => window.oxe.select('E30'));
await page.focus('.oxe-grid');
await page.keyboard.press('F2');
d = await drawn();
check(JSON.stringify(asRuns(d)) === JSON.stringify(shownBefore), 'F2 on the rich cell: drawn formatted as soon as it opens', JSON.stringify(d));
await page.keyboard.type('zz');
await page.keyboard.press('Escape');
check(JSON.stringify(await savedRuns(29, 4)) === JSON.stringify(shownBefore) && await drawn() === null, 'Esc: the change given up, the cell’s runs as they were, nothing drawn');

// Everything selected and a formula typed over it: plain, the box's own text; entered as a formula without runs.
await page.keyboard.press('F2');
await page.keyboard.press('Control+a');
await page.keyboard.type('=1+2');
check(await drawn() === null, 'a formula typed over the rich text: not drawn formatted');
await page.keyboard.press('Control+b');
check(await drawn() === null, 'Ctrl+B in a formula: nothing formatted');
await page.keyboard.press('Enter');
const formula = await page.evaluate(() => { const c = window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(29).get(4); return { f: c.f, v: c.v, runs: c.runs || null }; });
check(formula.f === '1+2' && formula.v === 3 && formula.runs === null, 'entered: a formula, no runs', JSON.stringify(formula));

// A number typed in a cell after formatting part of it: entered as a number, no runs.
await page.evaluate(() => window.oxe.select('E31'));
await page.focus('.oxe-grid');
await page.keyboard.type('1234');
await page.keyboard.press('F2');
await page.keyboard.press('Shift+ArrowLeft');
await page.keyboard.press('Control+b');
await page.keyboard.press('Enter');
const num = await page.evaluate(() => { const c = window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(30).get(4); return { v: c.v, runs: c.runs || null }; });
check(num.v === 1234 && num.runs === null, 'a number with a digit made bold: entered as the number, no runs', JSON.stringify(num));

// A run of the file's own (italic, 16pt, red, Courier New, as Excel writes it): drawn in its font while edited, text
// typed after it in that font, and that is what is saved.
{
  const { Workbook } = await import('../src/ooxml/workbook.js');
  const { parseXml } = await import('../src/ooxml/xml.js');
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  wb.setCell(0, 'J12', 'x');
  const { doc } = wb._part(0);
  const ns = doc.documentElement.namespaceURI;
  const c = [...doc.getElementsByTagNameNS(ns, 'c')].find((e) => e.getAttribute('r') === 'J12');
  while (c.firstChild) c.removeChild(c.firstChild);
  c.setAttribute('t', 'inlineStr');
  const is = parseXml(`<is xmlns="${ns}"><r><t xml:space="preserve">Plain </t></r>`
    + `<r><rPr><i/><sz val="16"/><color rgb="FFFF0000"/><rFont val="Courier New"/><family val="3"/></rPr><t>red</t></r></is>`).documentElement;
  c.appendChild(doc.importNode(is, true));
  wb.dirty.add(wb.sheets[0].path);
  fs.writeFileSync('test/out/rich-edit.xlsx', wb.toBytes());
}
await page.evaluate(() => { document.querySelector('input[type=file]').value = ''; });
await openFixtureFrom(page, 'test/out/rich-edit.xlsx', 'rich-edit.xlsx');
await page.evaluate(() => window.oxe.select('J12'));
await page.focus('.oxe-grid');
await page.keyboard.press('F2');
d = await drawn();
check(d?.length === 2 && d[1][0] === 'red' && d[1][2] === 'italic' && d[1][4] === 'rgb(255, 0, 0)' && d[1][5] === '21.3333px' && d[1][6] === 'Courier New',
  'F2 on a cell with a red 16pt Courier New run: drawn in that font while edited', JSON.stringify(d));
await page.keyboard.type('der');
d = await drawn();
check(d?.[1]?.[0] === 'redder' && d[1][4] === 'rgb(255, 0, 0)', 'typed after the red run: red 16pt Courier New, the font before the caret', JSON.stringify(d));
// (An IME composition after the red run, as Chromium's IME input gives it: drawn as it is composed, in that font.)
const cdp = await page.context().newCDPSession(page);
await cdp.send('Input.imeSetComposition', { text: 'か', selectionStart: 1, selectionEnd: 1 });
const composing = await drawn();
await cdp.send('Input.insertText', { text: '漢' });
d = await drawn();
check(composing?.[1]?.[0] === 'redderか' && d?.[1]?.[0] === 'redder漢' && d[1][4] === 'rgb(255, 0, 0)' && await boxText() === 'Plain redder漢',
  'an IME composition: drawn as it is composed and once committed, in the run’s font', JSON.stringify({ composing, d }));
await page.keyboard.press('Enter');
const bytes = await page.evaluate(() => [...window.oxe.save()]);
const { Workbook } = await import('../src/ooxml/workbook.js');
const back = Workbook.fromBytes(new Uint8Array(bytes)).sheet(0).cells.get(11).get(9);
const red = back.runs?.[1];
check(back.v === 'Plain redder漢' && back.runs?.length === 2 && red.text === 'redder漢' && red.font.i && red.font.size === 16 && red.font.name === 'Courier New' && red.font.color?.rgb === 'FFFF0000',
  'saved and read back: the run as it was drawn, its typed text red 16pt Courier New italic', JSON.stringify(back.runs));

await finish();
