// Browser check of Data > Queries & Connections (docs/QUERIES.md): the pane lists the fixture's queries (made by Excel,
// test/queries-excel.xlsx) with what Excel says under each and its connections; Refresh All asks the host
// (onRefreshQuery: the demo hands it to window.oxeRefreshQuery) and writes its rows into the query's table, one step
// of the history; Alt+F5 in the table refreshes it; a host's failure and an editor without the option say Excel's
// words; the Advanced Editor changes a query's M, kept in the file saved. Run after `npm run build`:
// node test/queries.mjs [outDir]
import fs from 'node:fs';
import path from 'node:path';
import { Workbook } from '../src/ooxml/workbook.js';
import { readQueries, QUERY_MESSAGES } from '../src/ooxml/queries.js';
import { editingPage } from './editing-kit.mjs';

const { page, check, finish, openFixtureFrom, value, outDir, ribbon } = await editingPage();
await openFixtureFrom(page, 'test/queries-excel.xlsx', 'queries-excel.xlsx');
const tableRef = () => page.evaluate(() => window.oxe.workbook.pkg.text('xl/tables/table1.xml').match(/ ref="([^"]+)"/)[1]);

// The pane, from the Data tab: the three queries, what Excel says under each, where each loads.
await page.click('.oxe-rtab >> text=Data');
await ribbon('Queries & Connections');
await page.waitForSelector('.oxe-queriesbar:not([hidden])');
const items = () => page.$$eval('.oxe-queriesbar .oxe-query-item', (els) => els.map((e) => [...e.querySelectorAll('.oxe-query-name, .oxe-query-line')].map((x) => x.textContent)));
const queries = await items();
check(JSON.stringify(queries) === JSON.stringify([['Sales', '4 rows loaded.', 'Load to: Table Sales (Sheet1!A1:C5)'], ['Big', 'Connection only.', 'Load to: Connection only'], ['Model', '4 rows loaded.', 'Load to: Data model']]),
  'Queries & Connections: each query, what Excel says under it, where it loads', JSON.stringify(queries));
await page.click('.oxe-queriesbar button:has-text("Connections")');
const conns = (await items()).map((x) => x[0]);
check(JSON.stringify(conns) === JSON.stringify(['Query - Big', 'Query - Model', 'Query - Sales', 'ThisWorkbookDataModel']), 'its Connections tab: the workbook\'s connections', JSON.stringify(conns));
await page.click('.oxe-queriesbar button:has-text("Queries")');

// Refresh All: the host asked for Sales (its M, its connection, its table), its six rows written into the table.
await page.evaluate(() => {
  window.asked = [];
  window.oxeRefreshQuery = async (req) => {
    window.asked.push(req);
    return [['Region', 'Item', 'Qty'], ['North', 'kiwi', 30], ['South', 'apple', 10], ['East', 'pear', 20], ['West', 'date', 50], ['North', 'fig', 40], ['South', 'lime', 5]];
  };
});
await page.click('.oxe-queriesbar button:has-text("Refresh All")');
await page.waitForFunction(() => window.oxe.workbook.pkg.text('xl/tables/table1.xml').includes('ref="A1:C7"'));
const asked = await page.evaluate(() => window.asked);
check(asked.length === 1 && asked[0].name === 'Sales' && asked[0].kind === 'query' && asked[0].m.startsWith('let') && asked[0].connection.name === 'Query - Sales' && asked[0].target.ref === 'A1:C5',
  'Refresh All: the host asked for the query loaded to a table, with its M, connection and table', JSON.stringify(asked));
check(await value('A7') === 'South' && await value('C6') === 40, 'its rows written into the table, the table resized', JSON.stringify([await value('A7'), await value('C6')]));
const status = (await items())[0];
check(status[1] === '6 rows loaded.' && status[2].endsWith('(Sheet1!A1:C7)'), 'the pane drawn again after the refresh: the rows loaded, the table’s new range', JSON.stringify(status));

// Ctrl+Z takes it back; Alt+F5 in the table (Refresh) does it again, through the host.
await page.focus('.oxe-grid');
await page.keyboard.press('Control+z');
check(await tableRef() === 'A1:C5' && await value('A7') == null, 'Ctrl+Z: the refresh undone, one step', await tableRef());
await page.evaluate(() => window.oxe.select('B3'));
await page.focus('.oxe-grid');
await page.keyboard.press('Alt+F5');
await page.waitForFunction(() => window.asked.length === 2);
await page.waitForFunction(() => window.oxe.workbook.pkg.text('xl/tables/table1.xml').includes('ref="A1:C7"'));
check(true, 'Alt+F5 in the query\'s table: refreshed through the host');

// The cell menu of a query table's cell has Refresh.
const cell = await page.evaluate(() => { const r = document.querySelector('.oxe-row[aria-rowindex="3"] .oxe-cell[aria-colindex="3"]').getBoundingClientRect(); return { x: r.x + r.width / 2, y: r.y + r.height / 2 }; });
await page.mouse.click(cell.x, cell.y, { button: 'right' });
await page.waitForSelector('.oxe-menu >> text=Paste Special...');
const hasRefresh = await page.isVisible('.oxe-menu >> text=Refresh');
await page.keyboard.press('Escape');
check(hasRefresh, 'the cell menu on the query\'s table: Refresh');

// A host that fails: Excel for the web's words, the reason after them, nothing changed.
await page.evaluate(() => { window.oxeRefreshQuery = async () => { throw new Error('The server is down'); }; });
await page.click('.oxe-queriesbar .oxe-query-item:first-child button:has-text("Refresh")');
await page.waitForSelector('dialog.oxe-dialog[role="alertdialog"]');
const failed = await page.textContent('dialog.oxe-dialog[role="alertdialog"]');
await page.keyboard.press('Enter');
check(failed.includes('The following connections failed to refresh:') && failed.includes('Query - Sales') && failed.includes('The server is down') && await tableRef() === 'A1:C7',
  'a host that fails: Excel\'s words and the reason, the table as it was', failed);

// The Advanced Editor: Big's M changed, Done, the file saved holds it.
await page.click('.oxe-queriesbar .oxe-query-item:nth-child(2) button:has-text("Edit")');
await page.waitForSelector('dialog.oxe-query-editor textarea');
const shownM = await page.inputValue('dialog.oxe-query-editor textarea');
await page.fill('dialog.oxe-query-editor textarea', shownM.replace('[Qty] > 15', '[Qty] > 25'));
await page.click('dialog.oxe-query-editor button:has-text("Done")');
const bytes = await page.evaluate(() => [...window.oxe.save()]);
fs.mkdirSync(outDir, { recursive: true });
const saved = path.join(outDir, 'queries-edited.xlsx');
fs.writeFileSync(saved, new Uint8Array(bytes));
const big = readQueries(Workbook.fromBytes(fs.readFileSync(saved))).queries.find((q) => q.name === 'Big');
check(shownM.includes('[Qty] > 15') && big.m.includes('[Qty] > 25'), 'the Advanced Editor: the query\'s M shown, changed and saved in the DataMashup', big.m);

// An editor without onRefreshQuery: Ctrl+Alt+F5 says the connection can't be refreshed, in Excel's words.
await page.evaluate((b) => {
  const host = document.createElement('div');
  host.id = 'host2';
  host.style.height = '300px';
  document.body.append(host);
  window.oxe2 = window.mountXlsxEditor(host, { bytes: new Uint8Array(b), fileName: 'nohost.xlsx' });
}, [...fs.readFileSync('test/queries-excel.xlsx')]);
await page.evaluate(() => window.oxe2.select('A2'));
await page.focus('#host2 .oxe-grid');
await page.keyboard.press('Control+Alt+F5');
await page.waitForSelector('#host2 dialog.oxe-dialog[role="alertdialog"]');
const refused = await page.textContent('#host2 dialog.oxe-dialog[role="alertdialog"]');
await page.keyboard.press('Enter');
check(refused.includes(QUERY_MESSAGES.cant(['Query - Sales']).split('\n')[0]) && refused.includes('Query - Sales'), 'without the host option: Excel\'s words that the connection can\'t be refreshed', refused);

await finish();
