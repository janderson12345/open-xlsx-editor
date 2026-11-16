// Browser check of Home > Cell Styles' Duplicate..., Delete and Merge Styles... on a fresh basic.xlsx, from the
// gallery's own menus: the Style box on a copy named "Mine 2"; a style of the workbook's own deleted, its cells back
// to Normal but for their own font, Undo bringing it back; a built-in style deleted gone from the gallery; and Merge
// Styles from a second editor on the page, Excel's question answered Yes (stylesmore.test.js has Excel's cases).
// Run after `npm run build`: node test/editing-stylesmore.mjs [outDir]
import path from 'node:path';
import { openFixture } from './browser.mjs';
import { editingPage } from './editing-kit.mjs';

const { page, check, finish, ribbon, outDir } = await editingPage();
await openFixture(page, 'basic.xlsx');

const cellXf = (ref) => page.evaluate((ref) => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  const wb = window.oxe.workbook;
  const s = wb.sheet(window.oxe.sheetIndex).cells.get(Number(m[2]) - 1)?.get(m[1].charCodeAt(0) - 65)?.s || 0;
  const xf = wb.styles.cellXfs[s];
  const font = wb.styles.fonts[xf.fontId] || {};
  const style = new XMLSerializer().serializeToString(wb.stylesDoc).match(new RegExp(`<cellStyle name="([^"]*)" xfId="${xf.xfId || 0}"`))?.[1];
  return { s, b: !!font.b, i: !!font.i, fill: xf.fillId, style };
}, ref);
/** The ribbon's menu closed, as a click outside it closes it. */
const closePop = () => page.evaluate(() => document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })));
/** The gallery's style names, opened and closed again. */
const gallery = async () => {
  let names = [];
  await ribbon('Cell Styles', async () => {
    await page.waitForSelector('.oxe-pop button.oxe-cstyle');
    names = await page.evaluate(() => [...document.querySelectorAll('.oxe-pop button.oxe-cstyle')].map((b) => b.title));
    await closePop();
  });
  return names;
};
/** Style `name`'s menu in the gallery (a right click), and its item `item`. */
const styleMenu = (name, item) => ribbon('Cell Styles', async () => {
  await page.click(`.oxe-pop button.oxe-cstyle[title="${name}"]`, { button: 'right' });
  await page.click(`.oxe-pop >> text="${item}"`);
});

// Mine: H2 bold and yellow, made a style; given to H3 and H4, H4 then italic of its own.
await page.evaluate(() => {
  const o = window.oxe;
  o.select('H2'); o.command('format', { font: { b: true } }); o.command('format', { fill: { rgb: 'FFFFFF00' } });
  o.command('newCellStyle', 'Mine');
  o.select('H3'); o.command('format', { cellStyle: 'Mine' });
  o.select('H4'); o.command('format', { cellStyle: 'Mine' }); o.command('format', { font: { i: true } });
  o.select('A1');
});
const items = await ribbon('Cell Styles', async () => {
  await page.click('.oxe-pop button.oxe-cstyle[title="Mine"]', { button: 'right' });
}).then(() => page.evaluate(() => [...document.querySelectorAll('.oxe-pop .oxe-item')].map((b) => b.textContent)));
await closePop();
check(['Apply', 'Modify...', 'Duplicate...', 'Delete'].every((x) => items.includes(x)), 'a style\'s menu: Apply, Modify..., Duplicate..., Delete', JSON.stringify(items));

// Duplicate...: the Style box on "Mine 2", what Mine includes; OK makes it, under Custom.
await styleMenu('Mine', 'Duplicate...');
await page.waitForSelector('dialog.oxe-style-box[open]');
const box = await page.evaluate(() => ({ name: document.querySelector('.oxe-style-box input[type=text]').value, parts: [...document.querySelectorAll('.oxe-style-part')].map((x) => x.textContent),
  checked: [...document.querySelectorAll('.oxe-style-box .oxe-style-row input')].map((x) => x.checked) }));
await page.screenshot({ path: path.join(outDir, 'stylesmore-duplicate.png') });
await page.click('.oxe-style-box button:text("OK")');
const afterDup = await gallery();
check(box.name === 'Mine 2' && /Bold/.test(box.parts[2]) && box.parts[4] === 'Shaded' && box.checked.every(Boolean) && afterDup.includes('Mine 2') && afterDup.includes('Mine'),
  'Duplicate...: the Style box on "Mine 2" with Mine\'s parts; OK puts it in the gallery beside Mine', JSON.stringify({ box, afterDup: afterDup.slice(0, 3) }));

// Delete Mine: H3 back to Normal (not bold, no fill), H4 keeping its own bold italic font; Undo brings Mine back.
const before = { h3: await cellXf('H3'), h4: await cellXf('H4') };
await styleMenu('Mine', 'Delete');
const after = { h3: await cellXf('H3'), h4: await cellXf('H4'), names: await gallery() };
check(before.h3.b && before.h3.style === 'Mine' && !after.h3.b && after.h3.fill === 0 && after.h3.style === 'Normal' && after.h4.b && after.h4.i && after.h4.fill === 0 && after.h3.s === before.h3.s
  && !after.names.includes('Mine') && after.names.includes('Mine 2'),
  'Delete: Mine gone from the gallery, H3 back to Normal, H4 keeping its own font, the cells\' formats changed in place', JSON.stringify({ before, after: { ...after, names: after.names.slice(0, 3) } }));
await page.evaluate(() => window.oxe.command('undo'));
const undone = { h3: await cellXf('H3'), names: await gallery() };
check(undone.h3.b && undone.h3.style === 'Mine' && undone.names.includes('Mine'), 'Undo: Mine back, and H3 bold with it', JSON.stringify(undone.h3));
await page.evaluate(() => window.oxe.command('redo'));

// Delete a built-in style: Good on H6, deleted; gone from the gallery (kept hidden in the file), H6 plain.
await page.evaluate(() => { window.oxe.select('H6'); window.oxe.command('format', { cellStyle: 'Good' }); window.oxe.select('A1'); });
const good = await cellXf('H6');
await styleMenu('Good', 'Delete');
const gone = { h6: await cellXf('H6'), names: await gallery(), hidden: await page.evaluate(() => /<cellStyle name="Good" xfId="\d+" builtinId="26" hidden="1"\/>/.test(new XMLSerializer().serializeToString(window.oxe.workbook.stylesDoc))) };
check(good.style === 'Good' && good.fill > 0 && gone.h6.fill === 0 && gone.h6.style === 'Normal' && !gone.names.includes('Good') && gone.names.includes('Bad') && gone.hidden,
  'Delete Good: gone from the gallery, kept hidden as Excel keeps it, H6 back to Normal', JSON.stringify({ good, gone: { ...gone, names: gone.names.length } }));
const noNormalDelete = await ribbon('Cell Styles', async () => {
  await page.click('.oxe-pop button.oxe-cstyle[title="Normal"]', { button: 'right' });
}).then(() => page.evaluate(() => [...document.querySelectorAll('.oxe-pop .oxe-item')].map((b) => b.textContent)));
await closePop();
check(!noNormalDelete.includes('Delete') && noNormalDelete.includes('Duplicate...'), 'Normal\'s menu: no Delete', JSON.stringify(noNormalDelete));

// Merge Styles...: a second editor on the page, this workbook as saved now, its Mine 2 made italic and not bold and a
// style "Theirs" added; the box lists it, Excel's question answered Yes: Mine 2 put in place of ours, Theirs added.
await page.evaluate(() => {
  const host = document.createElement('div');
  host.style.cssText = 'position:fixed;left:-2000px;top:0;width:800px;height:500px';
  document.body.append(host);
  const other = window.mountXlsxEditor(host, { bytes: window.oxe.save(), fileName: 'Other.xlsx' });
  window.otherEditor = other;
  other.command('modifyCellStyle', 'Mine 2', { font: { b: false, i: true } });
  other.select('A1'); other.command('format', { font: { i: true } });
  other.command('newCellStyle', 'Theirs');
});
await ribbon('Cell Styles', () => page.click('.oxe-pop >> text=Merge Styles...'));
await page.waitForSelector('dialog[open] select.oxe-dialog-list');
const listed = await page.evaluate(() => ({ title: document.querySelector('dialog[open] .oxe-dialog-title')?.textContent, books: [...document.querySelectorAll('dialog[open] select.oxe-dialog-list option')].map((o) => o.textContent) }));
await page.screenshot({ path: path.join(outDir, 'stylesmore-merge.png') });
await page.click('dialog[open] button:text("OK")');
await page.waitForSelector('dialog[open][role=alertdialog]');
const question = await page.evaluate(() => document.querySelector('dialog[open] .oxe-dialog-text')?.textContent);
await page.click('dialog[open] button:text("Yes")');
const merged = await page.evaluate(() => {
  const wb = window.oxe.workbook;
  const text = new XMLSerializer().serializeToString(wb.stylesDoc);
  const fontOf = (name) => { const id = new RegExp(`<cellStyle name="${name}" xfId="(\\d+)"`).exec(text)?.[1]; return id == null ? null : wb.styles.fonts[wb.styles.styleXfs[Number(id)].fontId]; };
  return { mine2: fontOf('Mine 2'), theirs: fontOf('Theirs') };
});
check(listed.title === 'Merge Styles' && listed.books.join() === 'Other.xlsx' && question === 'Merge styles that have the same names?' && merged.mine2?.i && !merged.mine2?.b && merged.theirs?.i,
  'Merge Styles...: the other editor\'s workbook listed, Excel\'s question, Yes: Mine 2 italic as theirs, Theirs added', JSON.stringify({ listed, question, merged }));
await page.evaluate(() => window.oxe.command('undo'));
const unmerged = await page.evaluate(() => /<cellStyle name="Theirs"/.test(new XMLSerializer().serializeToString(window.oxe.workbook.stylesDoc)));
check(!unmerged, 'Undo: the merge taken back as one step');

// Saved and read back: the hidden Good, Mine and Mine 2 as written.
const bytes = await page.evaluate(() => [...window.oxe.save()]);
const { Workbook } = await import('../src/ooxml/workbook.js');
const back = Workbook.fromBytes(new Uint8Array(bytes));
const { customCellStyles, deletedBuiltins } = await import('../src/ooxml/cellstyleedit.js');
check(customCellStyles(back).join() === 'Mine 2' && [...deletedBuiltins(back)].join() === '26', 'saved and read back: Mine 2 of its own (Mine deleted), Good hidden', JSON.stringify(customCellStyles(back)));
await page.evaluate(() => { window.otherEditor.destroy(); });

await finish();
