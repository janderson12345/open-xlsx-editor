// Browser check of what moves with rows, columns and cells as Excel moves it (docs/ROWCOL.md), through the editor's
// commands: on test/fixtures/refmoves.xlsx, Circle Invalid Data's ovals gone after a row is inserted, the sparklines,
// validation rules and notes moved with it, and back on Ctrl+Z; notes going with their cells in a sort; on
// test/fixtures/slicers.xlsx, rows and columns inserted and deleted on a sheet with slicers and a timeline, a table
// slicer deleted with its column. Run after `npm run build`: node test/editing-refmoves.mjs [outDir]
import { editingPage } from './editing-kit.mjs';

const { page, check, finish, openFixtureFrom } = await editingPage();

await openFixtureFrom(page, 'test/fixtures/refmoves.xlsx', 'refmoves.xlsx');
const circlesNow = () => page.evaluate(() => document.querySelectorAll('.oxe-circles ellipse').length);
/** The sheet's sparklines (cell <- data), rules' ranges, and notes' cells, from the workbook as it is now. */
const parts = () => page.evaluate(() => {
  const wb = window.oxe.workbook;
  const xml = new XMLSerializer().serializeToString(wb._part(0).doc);
  return {
    sparklines: [...xml.matchAll(/<xm:f>([^<]*)<\/xm:f><xm:sqref>([^<]*)<\/xm:sqref>/g)].map((m) => `${m[2]}<-${m[1]}`).sort().join(' '),
    rules: [...xml.matchAll(/<dataValidation\b[^>]*\ssqref="([^"]*)"/g), ...xml.matchAll(/<xm:sqref>(E[^<]*)<\/xm:sqref>/g)].map((m) => m[1]).sort().join(' '),
    notes: [...(wb.pkg.text('xl/comments1.xml') || '').matchAll(/<comment ref="([^"]+)"/g)].map((m) => m[1]).sort().join(' '),
  };
});
const start = await parts();

// Data > Data Validation > Circle Invalid Data, then a row inserted at 1 through the row's heading menu's command.
await page.click('.oxe-rtab >> text=Data');
await page.click('.oxe-toolbar button.oxe-split-more[aria-label^="Data Validation"]');
await page.click('.oxe-pop >> text=Circle Invalid Data');
const circled = await circlesNow();
await page.evaluate(() => { window.oxe.command('selectLines', 'row', 0); window.oxe.command('lines', 'row', 1); });
const inserted = await parts();
const afterInsert = await circlesNow();
await page.evaluate(() => window.oxe.command('undo'));
const undone = await parts();
check(circled === 5 && afterInsert === 0, 'Circle Invalid Data: five cells circled (the rule of the Excel 2010 form isn\'t read yet); none after a row is inserted, as Excel\'s go', JSON.stringify({ circled, afterInsert }));
check(inserted.sparklines.includes('O3<-Sheet1!B3:M3') && inserted.sparklines.includes('Q3<-Data!B2:M2') && inserted.rules === 'B11:B16 C11:C16 D11:D16 E11:E16' && inserted.notes === 'B13 C14 F21',
  'a row inserted at 1: sparklines, validation rules (both forms) and notes a row down, the data of another sheet\'s sparkline as it was', JSON.stringify(inserted));
check(JSON.stringify(undone) === JSON.stringify(start), 'Ctrl+Z: all of them back', JSON.stringify({ start, undone }));

// Sort B10:E15 by C, largest first (the Sort box's answer): B12's note to B10 and C13's thread to C14, as in Excel; the
// rules stay; Ctrl+Z puts the notes back.
await page.evaluate(() => {
  window.oxe.select('B10');
  window.oxe.command('sort', { area: { r0: 9, c0: 1, r1: 14, c1: 4 }, keys: [{ at: 2, desc: true }], header: false });
});
const sorted = await parts();
const threads = await page.evaluate(() => [...(window.oxe.workbook.pkg.text('xl/threadedComments/threadedComment1.xml') || '').matchAll(/<threadedComment ref="([^"]+)"/g)].map((m) => m[1]).join(' '));
await page.evaluate(() => window.oxe.command('undo'));
const unsorted = await parts();
check(sorted.notes === 'B10 C14 F20' && threads === 'C14' && sorted.rules === start.rules && unsorted.notes === start.notes,
  'Sort: notes and the threaded comment go with their cells, the rules stay; Ctrl+Z puts them back', JSON.stringify({ sorted, threads, unsorted }));

// A sheet with slicers and a timeline: a row inserted at 1 (no longer refused), the slicers below it moving down with
// the table and the pivot; column C (the table's Region) deleted, the Region slicer goes with its cache and name.
await openFixtureFrom(page, 'test/fixtures/slicers.xlsx', 'slicers.xlsx');
const slicerState = () => page.evaluate(() => {
  const wb = window.oxe.workbook;
  const drawing = wb.pkg.text('xl/drawings/drawing1.xml') || '';
  return {
    item: /<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>(\d+)<\/xdr:col><xdr:colOff>0<\/xdr:colOff><xdr:row>(\d+)<\/xdr:row>[\s\S]*?name="Item"/.exec(drawing)?.slice(1).join(','),
    table: / ref="([^"]*)"/.exec(wb.pkg.text('xl/tables/table1.xml'))?.[1],
    region: drawing.includes('name="Region"') && wb.pkg.has('xl/slicerCaches/slicerCache2.xml') && wb.definedNames.some((d) => d.name === 'Slicer_Region'),
    alert: !!document.querySelector('dialog[open]'),
  };
});
await page.evaluate(() => { window.oxe.command('selectLines', 'row', 0); window.oxe.command('lines', 'row', 1); });
const rowIn = await slicerState();
await page.evaluate(() => window.oxe.command('undo'));
await page.evaluate(() => { window.oxe.command('selectLines', 'col', 2); window.oxe.command('lines', 'col', -1); });
const colOut = await slicerState();
check(!rowIn.alert && rowIn.item === '15,2' && rowIn.table === 'B3:E15' && rowIn.region,
  'slicers: a row inserted on their sheet, the Item slicer a row down (P3), the table B3:E15', JSON.stringify(rowIn));
check(!colOut.alert && colOut.item === '14,1' && colOut.table === 'B2:D14' && !colOut.region,
  'slicers: column C deleted, the Region slicer gone with its cache and name, the Item slicer a column left', JSON.stringify(colOut));

await finish();
