// Browser check of Show/Hide Note, Show All Notes and Convert to Comments (Review > Notes) as saved to the file:
// D4's note shown stays shown with the pointer away, and the workbook saved and opened again opens with it shown; Show
// All Notes shows both notes and again hides both; Convert to Comments makes notes.xlsx's two notes threads (purple
// marks, the note's text without the author's name, which is the thread's author) and Ctrl+Z puts the notes back.
// Run after `npm run build`: node test/notes-visible.mjs [outDir]
import fs from 'node:fs';
import { editingPage } from './editing-kit.mjs';
import { openFixture } from './browser.mjs';

const { page, check, finish, outDir, menu, openFixtureFrom, pointAt } = await editingPage();

/** The note boxes shown (not threads'), by the text of each. */
const boxes = () => page.evaluate(() => [...document.querySelectorAll('.oxe-note-box:not(.oxe-thread-box)')].filter((b) => b.getClientRects().length).map((b) => b.textContent));
const marks = () => page.evaluate(() => ({ red: document.querySelectorAll('.oxe-grid > * .oxe-note-mark:not(.oxe-thread-mark)').length, purple: document.querySelectorAll('.oxe-grid > * .oxe-thread-mark').length }));
const away = () => page.mouse.move(5, 5);

await openFixture(page, 'notes.xlsx');
await away();
await page.click('.oxe-rtab >> text="Review"');
await page.evaluate(() => window.oxe.select('D4'));
await menu('Notes', 'Show/Hide Note');
await page.waitForTimeout(100);
const shown = await boxes();
const vml = await page.evaluate(() => window.oxe.workbook.pkg.text('xl/drawings/vmlDrawing1.vml'));
const d4 = [...vml.matchAll(/<v:shape\b[\s\S]*?<\/v:shape>/g)].map((m) => m[0]).find((s) => /<x:Row>3<\/x:Row>/.test(s));
check(shown.length === 1 && /First line/.test(shown[0]) && /visibility:visible/.test(d4) && /<x:Visible\/>/.test(d4),
  'Show/Hide Note on D4: its box shown with the pointer away, written in its shape (visibility:visible, x:Visible)', JSON.stringify(shown));

// (Saved and opened again: D4's note opens shown.)
const file = `${outDir}/notes-shown.xlsx`;
fs.writeFileSync(file, Buffer.from(await page.evaluate(() => Array.from(window.oxe.workbook.toBytes()))));
await openFixtureFrom(page, file, 'notes-shown.xlsx');
await away();
await page.waitForSelector('.oxe-note-mark');
const reopened = await boxes();
check(reopened.length === 1 && /First line/.test(reopened[0]), 'saved and opened again: D4\'s note shown', JSON.stringify(reopened));

// (Show All Notes: both shown; again, both hidden, D4's too, as Excel's.)
await menu('Notes', 'Show All Notes');
await page.waitForTimeout(100);
const all = await boxes();
await menu('Notes', 'Show All Notes');
await page.waitForTimeout(100);
const none = await boxes();
check(all.length === 2 && none.length === 0, 'Show All Notes: both notes shown; again: both hidden', JSON.stringify({ all, none }));

// (Convert to Comments: the two notes threads, the author's name taken out of the text; Ctrl+Z: notes again.)
const before = await marks();
await menu('Notes', 'Convert to Comments');
await page.waitForTimeout(100);
const after = await marks();
const b2 = await pointAt(1, 1);
await page.mouse.move(b2.x, b2.y);
await page.waitForSelector('.oxe-thread-box', { timeout: 2000 }).catch(() => {});
const thread = await page.evaluate(() => document.querySelector('.oxe-thread-box')?.textContent || '');
await away();
const parts = await page.evaluate(() => { const p = window.oxe.workbook.pkg; return { threads: p.text('xl/threadedComments/threadedComment1.xml'), notes: (p.text('xl/comments1.xml').match(/<comment\b/g) || []).length }; });
check(before.red === 2 && before.purple === 1 && after.red === 0 && after.purple === 3 && /Avery Chen/.test(thread) && /Check this total/.test(thread) && !/Avery Chen:/.test(thread)
  && /<threadedComment ref="B2" personId="[^"]+" id="[^"]+"><text>Check this total<\/text>/.test(parts.threads) && parts.notes === 3,
'Convert to Comments: both notes threads (purple marks), B2\'s by Avery Chen without the name in its text, no date', JSON.stringify({ before, after, thread }));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+z');
await page.waitForTimeout(100);
check(JSON.stringify(await marks()) === JSON.stringify(before), 'Ctrl+Z: the notes back', JSON.stringify(await marks()));

await finish();
