import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { addThread, replyThread, editThread, workbookPeople } from '../src/ooxml/threadedit.js';
import { readNotes } from '../src/ooxml/notes.js';
import { mentionPeople, mentionQuery, matchPeople } from '../src/editor/mentions.js';

// @mentions written as [MS-XLSX] CT_Mention gives them and Excel 365 writes a mention picked from its list: the text
// keeps "@Name", and a mentions element after it gives each one's person, its own id, where it starts and its length
// with the "@". (Excel makes a mention only from a signed-in organisation's directory, which COM can't reach, so these
// are checked against the format, and Excel opens the result: docs/NOTES.md.)
const open = (f) => Workbook.fromBytes(new Uint8Array(fs.readFileSync(`test/fixtures/${f}`)));
const again = (wb) => Workbook.fromBytes(wb.toBytes());
const threadsXml = (wb) => wb.pkg.text(Object.keys(wb.pkg.files).find((p) => /threadedComments\/threadedComment\d+\.xml$/.test(p)));
const SAM = { name: 'Sam Lee', userId: 'sam.lee@example.com', providerId: 'AD' };

test('a comment with @mentions: the text as typed, a mention for each name, the people added to the persons part', () => {
  const wb = open('basic.xlsx');
  const text = '@Avery Chen please check, @Sam Lee too';
  addThread(wb, 0, 1, 1, text, 'Robin Park', { when: new Date(2026, 10, 16, 9, 30), mentions: [{ name: 'Avery Chen' }, SAM, { name: 'Nobody Here' }] });
  const xml = threadsXml(wb);
  const m = [...xml.matchAll(/<mention mentionpersonId="(\{[^"]+\})" mentionId="(\{[^"]+\})" startIndex="(\d+)" length="(\d+)"\/>/g)];
  assert.match(xml, /<text>@Avery Chen please check, @Sam Lee too<\/text><mentions><mention /);
  assert.deepEqual(m.map((x) => [Number(x[3]), Number(x[4])]), [[0, 11], [26, 8]]);
  assert.notEqual(m[0][2], m[1][2]);
  const people = workbookPeople(wb);
  const sam = people.find((p) => p.name === 'Sam Lee');
  // (Someone mentioned is a person of provider PeoplePicker, their user id kept: as Excel rewrote AD on reading.)
  assert.deepEqual([sam.userId, sam.providerId], ['sam.lee@example.com', 'PeoplePicker']);
  assert.equal(people.find((p) => p.name === 'Avery Chen').providerId, 'PeoplePicker');
  assert.ok(!people.some((p) => p.name === 'Nobody Here'), 'someone picked but not in the text: no mention, no person');
  assert.equal(m[1][1], sam.id);
  const back = readNotes(again(wb), 0).get('1,1');
  assert.equal(back.items[0].text, text);
  assert.deepEqual(back.items[0].mentions, [{ start: 0, length: 11, name: 'Avery Chen' }, { start: 26, length: 8, name: 'Sam Lee' }]);
});

test('mentions in a reply and an edit; a name inside a longer one is not a mention; no mentions element without any', () => {
  const wb = open('basic.xlsx');
  addThread(wb, 0, 2, 2, 'Plain comment', 'Robin Park');
  assert.doesNotMatch(threadsXml(wb), /<mentions/);
  replyThread(wb, 0, 2, 2, 'Over to @Anna Bell and @Ann', 'Robin Park', { mentions: [{ name: 'Ann' }, { name: 'Anna Bell' }] });
  let items = readNotes(wb, 0).get('2,2').items;
  assert.deepEqual(items[1].mentions.map((x) => [x.start, x.length, x.name]), [[8, 10, 'Anna Bell'], [23, 4, 'Ann']]);
  editThread(wb, 0, 2, 2, 1, 'Over to @Ann alone', { mentions: [{ name: 'Ann' }] });
  items = readNotes(wb, 0).get('2,2').items;
  assert.deepEqual(items[1].mentions.map((x) => [x.start, x.length]), [[8, 4]]);
  editThread(wb, 0, 2, 2, 1, 'No one now');
  assert.deepEqual(readNotes(wb, 0).get('2,2').items[1].mentions, []);
  assert.equal((threadsXml(wb).match(/<mentions>/g) || []).length, 0);
});

test('the people @ offers: the host’s and the workbook’s, each once; narrowed by what follows the @', () => {
  const people = mentionPeople([SAM, { name: 'Avery Chen' }], [{ name: 'avery chen', userId: 'x' }, { name: 'Robin Park' }]);
  assert.deepEqual(people.map((p) => p.name), ['Sam Lee', 'Avery Chen', 'Robin Park']);
  assert.deepEqual(mentionQuery('Hi @Sa', 6), { at: 3, query: 'Sa' });
  assert.deepEqual(mentionQuery('@', 1), { at: 0, query: '' });
  assert.equal(mentionQuery('mail a@b', 8), null, 'an @ inside a word is not a mention');
  assert.equal(mentionQuery('Hi @Sam  x', 10), null);
  assert.deepEqual(matchPeople(people, 'park').map((p) => p.name), ['Robin Park']);
  assert.deepEqual(matchPeople(people, 'a').map((p) => p.name), ['Avery Chen', 'Sam Lee', 'Robin Park'], 'a word starting so first');
  assert.deepEqual(matchPeople(people, 'example').map((p) => p.name), ['Sam Lee'], 'by user id too');
});

test('the same name as author and as mentioned: two people, as Excel keeps them', () => {
  const wb = open('basic.xlsx');
  addThread(wb, 0, 1, 1, '@Avery Chen please check', 'Robin Park', { mentions: [{ name: 'Avery Chen' }] });
  replyThread(wb, 0, 1, 1, 'Done', 'Avery Chen');
  replyThread(wb, 0, 1, 1, 'Thanks @Avery Chen', 'Robin Park', { mentions: [{ name: 'Avery Chen' }] });
  const avery = workbookPeople(wb).filter((p) => p.name === 'Avery Chen');
  assert.deepEqual(avery.map((p) => p.providerId).sort(), ['None', 'PeoplePicker']);
  const ids = [...threadsXml(wb).matchAll(/mentionpersonId="([^"]+)"/g)].map((m) => m[1]);
  assert.deepEqual(ids, [avery.find((p) => p.providerId === 'PeoplePicker').id, avery.find((p) => p.providerId === 'PeoplePicker').id]);
});

test('a workbook with mentions for Excel to read (test/out/mentions/mentions.xlsx, excel-mentions-check.ps1)', () => {
  const wb = open('basic.xlsx');
  addThread(wb, 0, 1, 1, '@Avery Chen please check, @Sam Lee too', 'Robin Park', { mentions: [{ name: 'Avery Chen' }, SAM] });
  replyThread(wb, 0, 1, 1, 'Done, @Robin Park', 'Avery Chen', { mentions: [{ name: 'Robin Park' }] });
  fs.mkdirSync('test/out/mentions', { recursive: true });
  fs.writeFileSync('test/out/mentions/mentions.xlsx', wb.toBytes());
  assert.equal(readNotes(again(wb), 0).get('1,1').items[1].mentions[0].name, 'Robin Park');
});

test("Excel's reading of that workbook: every mention kept, its people the same and in the same order", () => {
  // (test/mentions-excel.json: Excel opened test/out/mentions/mentions.xlsx and saved a copy; its mention and person
  // elements were the same as ours, ids and all. The ids are new each run, so the rest is compared here.)
  const X = JSON.parse(fs.readFileSync('test/mentions-excel.json', 'utf8'));
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/out/mentions/mentions.xlsx')));
  const strip = (el) => el.replace(/\s(id|mentionpersonId|mentionId)="[^"]*"/g, '');
  const ours = [...threadsXml(wb).matchAll(/<mention [^>]*\/>/g)].map((m) => strip(m[0]));
  assert.deepEqual(ours, X.mentionsKept.map(strip));
  const personsXml = wb.pkg.text('xl/persons/person.xml');
  assert.deepEqual([...personsXml.matchAll(/<person [^>]*\/>/g)].map((m) => strip(m[0])), X.persons.map(strip));
  assert.equal(X.threads[0].text, '@Avery Chen please check, @Sam Lee too');
  assert.equal(X.threads[0].replies[0].text, 'Done, @Robin Park');
});
