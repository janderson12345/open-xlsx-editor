import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { OpcPackage } from '../src/ooxml/package.js';
import { validationXml } from '../src/ooxml/dvedit.js';
import { parseDecls } from '../src/ooxml/htmltable.js';
import { Pattern, STEPS, TOTAL_STEPS } from '../src/calc/pcre-match.js';
import { E } from '../src/calc/values.js';

// Text a workbook, the clipboard or a pattern brings, handled so it can't change what the code does: the security
// review of the November merges.

test('removeRel takes a relationship id as text, not as a pattern', () => {
  const rels = '<Relationships><Relationship Id="rId1" Target="a.xml"/><Relationship Id="(a+)+" Target="b.xml"/></Relationships>';
  const files = { 'xl/drawings/_rels/drawing1.xml.rels': rels };
  const pkg = { text: (p) => files[p] ?? null, setText: (p, s) => { files[p] = s; } };
  OpcPackage.prototype.removeRel.call(pkg, 'xl/drawings/drawing1.xml', '(a+)+');
  assert.equal(files['xl/drawings/_rels/drawing1.xml.rels'], '<Relationships><Relationship Id="rId1" Target="a.xml"/></Relationships>');
  OpcPackage.prototype.removeRel.call(pkg, 'xl/drawings/drawing1.xml', 'rId.');
  assert.match(files['xl/drawings/_rels/drawing1.xml.rels'], /Id="rId1"/);
});

test('a validation rule read from a file writes only the error styles and types Excel has', () => {
  const ranges = [{ r0: 0, c0: 0, r1: 0, c1: 0 }];
  const xml = validationXml({ type: 'list', f1: 'a,b', error: { style: 'warning" evil="1' } }, ranges);
  assert.doesNotMatch(xml, /evil/);
  assert.doesNotMatch(xml, /errorStyle/);
  assert.match(validationXml({ type: 'whole', op: 'greater', f1: '1', error: { style: 'information' } }, ranges), /errorStyle="information"/);
  const odd = validationXml({ type: 'constructor', op: 'toString', f1: '1' }, ranges);
  assert.doesNotMatch(odd, /function|native/);
});

test('clipboard CSS with a long run of spaces is read at once, !important still cut off', () => {
  const started = Date.now();
  const decls = parseDecls(`color: red${' '.repeat(200000)}x; font-weight: 700 !important`);
  assert.ok(Date.now() - started < 500);
  assert.equal(decls['font-weight'], '700');
  assert.ok(decls.color.startsWith('red'));
});

test('all the searches of one REGEXREPLACE share a limit', () => {
  const p = new Pattern('a', false);
  let calls = 0;
  p.exec = (s, from) => { calls++; p.S.steps = STEPS; return { index: from, start: from, end: from + 1, caps: [] }; };
  assert.throws(() => p.all('a'.repeat(1000)), (e) => e === E.value);
  assert.equal(calls, TOTAL_STEPS / STEPS + 1);
});
