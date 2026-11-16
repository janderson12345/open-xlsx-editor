import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { readObjects, anchorRect, setObjectAnchor, deleteObjects } from '../src/ooxml/drawobjects.js';
import { presetPaths, arrowHead, customPaths } from '../src/ooxml/presets.js';
import { paintObject, frameOf, resolve } from '../src/editor/drawpaint.js';

// Pictures, shapes, icons and groups (docs/LAYOUT.md, Pictures, shapes and icons), read from test/fixtures/drawings.xlsx
// (drawings.mjs) and drawn as Excel's PDF of it has them.
const book = (name = 'drawings.xlsx') => Workbook.fromBytes(new Uint8Array(fs.readFileSync(`test/fixtures/${name}`)));
// (The fixture's screen sizes: columns 64px, C 120px, G hidden; rows 20px, row 6 hidden.)
const axis = (size) => ({ size, start: (i) => { let x = 0; for (let k = 0; k < i; k++) x += size(k); return x; } });
const cols = axis((c) => (c === 6 ? 0 : c === 2 ? 120 : 64));
const rows = axis((r) => (r === 5 ? 0 : 20));
const THEME = { fonts: { major: 'Calibri', minor: 'Calibri' }, colors: ['000000', 'FFFFFF', '1F3864', 'E7E6E6', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'], lines: [6350, 6350, 6350] };

test('drawings.xlsx: every anchor read, in the drawing\'s order, with what it holds', () => {
  const list = readObjects(book(), 0);
  assert.deepEqual(list.map((o) => `${o.k} ${o.kind} ${o.name}`), ['0 picture Quarters', '1 picture Cropped', '2 picture Turned', '3 picture Icon 1', '4 shape Rectangle 5', '5 shape Rounded 6', '6 shape Oval 7',
    '7 shape Triangle 8', '8 shape Arrow 9', '9 shape TextBox 10', '10 shape Connector 11', '11 shape Elbow 12', '12 shape Cloud 13', '13 group Group 14', '14 shape Star 15', '15 shape Hidden 16']);
  const [quarters, cropped, turned, icon] = list;
  assert.equal(quarters.anchor.editAs, 'oneCell');
  assert.equal(quarters.descr, 'Four coloured quarters');
  assert.equal(quarters.item.src.mime, 'image/png');
  assert.deepEqual(cropped.item.crop, { l: 0.25, t: 0.1, r: 0, b: 0.2 });
  assert.equal(turned.anchor.type, 'oneCellAnchor');
  assert.deepEqual([turned.item.xfrm.rot, turned.item.xfrm.flipH, turned.item.xfrm.flipV], [30, true, false]);
  // (An icon: its SVG, and the PNG beside it for what can't draw SVG.)
  assert.equal(icon.item.svg.mime, 'image/svg+xml');
  assert.equal(icon.item.src.mime, 'image/png');
  assert.match(icon.item.svg.url, /^data:image\/svg\+xml;base64,/);
  const rounded = list[5].item;
  assert.equal(rounded.geom.prst, 'roundRect');
  assert.equal(rounded.style.fill.idx, '1');
  assert.deepEqual(rounded.text.paras[0].runs.map((r) => [r.text, r.size, r.bold]), [['Rounded', 14, true]]);
  assert.equal(rounded.text.anchor, 'ctr');
  assert.equal(list[9].item.txBox, true);
  assert.equal(list[10].item.connector, true);
  assert.deepEqual(list[10].item.line.tail, { type: 'triangle', w: 'med', len: 'med' });
  assert.equal(list[13].item.children.length, 2);
  assert.equal(list[14].anchor.type, 'absoluteAnchor');
  assert.equal(list[15].hidden, true);
});

test('where objects sit: a two-cell anchor corner to corner, a one-cell its extent, an absolute its position', () => {
  const list = readObjects(book(), 0);
  // (Quarters, editAs oneCell over hidden row 6: corner to corner, as Excel's PDF has it, 6 rows not 7.)
  assert.deepEqual(anchorRect(list[0].anchor, rows, cols, list[0].item), { x: 64, y: 20, w: 248, h: 120 });
  // (Cropped, over hidden column G: three columns wide.)
  assert.deepEqual(anchorRect(list[1].anchor, rows, cols, list[1].item), { x: 64 * 4 + 120, y: 20, w: 192, h: 120 });
  assert.deepEqual(anchorRect(list[2].anchor, rows, cols, list[2].item), { x: 64 * 8 + 120 + 8, y: 40 + 4, w: 120, h: 80 });
  assert.deepEqual(anchorRect(list[14].anchor, rows, cols, list[14].item), { x: 900, y: 220, w: 80, h: 80 });
});

test('preset outlines: the common shapes by ECMA-376\'s formulas, others their box', () => {
  const rect = presetPaths('rect', 100, 50);
  assert.deepEqual(rect.paths[0], [['M', 0, 0], ['L', 100, 0], ['L', 100, 50], ['L', 0, 50], ['Z']]);
  assert.equal(rect.known, true);
  // (A right arrow 100 by 40: its head 20 long (half the short side), its shaft half the height.)
  assert.deepEqual(presetPaths('rightArrow', 100, 40).paths[0].slice(0, 4), [['M', 0, 10], ['L', 80, 10], ['L', 80, 0], ['L', 100, 20]]);
  assert.deepEqual(presetPaths('triangle', 60, 30).paths[0], [['M', 0, 30], ['L', 30, 0], ['L', 60, 30], ['Z']]);
  assert.equal(presetPaths('line', 10, 10).line, true);
  assert.equal(presetPaths('cloud', 10, 10).known, false);
  // (An ellipse's text box: inside it, 14.6% in from each side.)
  const e = presetPaths('ellipse', 100, 100).text;
  assert.ok(Math.abs(e.x - 14.645) < 0.01 && Math.abs(e.w - 70.71) < 0.01);
  // (Medium arrowheads 8px long and wide on lines up to 2.67px, as Excel's PDF has them.)
  const head = arrowHead('triangle', [100, 0], [0, 0], 1.333);
  assert.deepEqual(head.cmds.slice(0, 3).map((c) => c.slice(1).map((v) => Math.round(v * 100) / 100)), [[100, 0], [92, 4], [92, -4]]);
  // (Custom geometry: a path's own coordinates scaled to the shape.)
  const cust = customPaths([{ w: 10, h: 10, cmds: [{ op: 'moveTo', pts: [[0, 0]] }, { op: 'lnTo', pts: [[10, 10]] }, { op: 'close', pts: [] }] }], 50, 20);
  assert.deepEqual(cust[0].cmds, [['M', 0, 0], ['L', 50, 20], ['Z']]);
});

test('what objects paint: Excel\'s colours, turned frames, a group\'s children, a crop', () => {
  const list = readObjects(book(), 0);
  const paint = (k) => paintObject(list[k], anchorRect(list[k].anchor, rows, cols, list[k].item), THEME);
  // (Accent 2 at 75% luminance in fractional HSL: Excel's C55A11; the default style's line, accent 1 shaded 50%: 2F528F.)
  assert.equal(paint(7)[0].fill.css, '#C55A11');
  assert.equal(paint(7)[0].line, null);
  const rounded = paint(5);
  assert.deepEqual([rounded[0].fill.css, rounded[0].line.css, +rounded[0].line.width.toFixed(3)], ['#4472C4', '#2F528F', 0.667]);
  const text = rounded.find((o) => o.t === 'text');
  assert.equal(text.body.paras[0].runs[0].color.css, '#FFFFFF');
  assert.equal(text.body.anchor, 'ctr');
  // (Turned 45 degrees, the arrow's frame is its anchor's box with width and height swapped about its middle.)
  assert.deepEqual(frameOf({ x: 0, y: 0, w: 200, h: 100 }, 45), { x: 50, y: -50, w: 100, h: 200 });
  assert.deepEqual(frameOf({ x: 0, y: 0, w: 200, h: 100 }, 30), { x: 0, y: 0, w: 200, h: 100 });
  // (The group's children: their space half the group's own, so each is drawn twice its size.)
  const g = paint(13).filter((o) => o.t === 'path');
  const box = (cmds) => { const xs = cmds.filter((c) => c[0] !== 'Z').flatMap((c) => c.slice(1).filter((_, i) => i % 2 === 0)); return [Math.min(...xs), Math.max(...xs)]; };
  const r = anchorRect(list[13].anchor, rows, cols, list[13].item);
  assert.deepEqual(box(g[0].cmds).map(Math.round), [Math.round(r.x), Math.round(r.x + r.w / 2)]);
  // (Cropped 25% at the left: the whole picture 4/3 the frame's width, starting a third of the frame before it.)
  const img = paint(1).find((o) => o.t === 'image');
  assert.deepEqual(img.m.map((v) => Math.round(v * 100) / 100), [256, 0, 0, 171.43, 312, 2.86]);
  // (A connector: a line with its arrowhead filled at its end.)
  const cx = paint(10);
  assert.equal(cx.length, 2);
  assert.equal(cx[0].fill, null);
  assert.equal(cx[1].fill.css, '#4472C4');
  // (A hidden shape paints nothing.)
  assert.deepEqual(paint(15), []);
  assert.deepEqual(resolve({ scheme: 'accent1', mods: [['alpha', 0.5]] }, THEME), { css: '#4472C4', alpha: 0.5 });
});

test('moving, sizing and deleting objects: only their anchors and parts change', () => {
  const wb = book();
  const before = wb.pkg.text('xl/drawings/drawing1.xml');
  setObjectAnchor(wb, 0, 2, { col: 1, colOff: 0, row: 3, rowOff: 9525 }, null, { x: 64 * 9525, y: 41 * 9525, w: 100 * 9525, h: 50 * 9525 });
  let after = wb.pkg.text('xl/drawings/drawing1.xml');
  const turned = /<xdr:oneCellAnchor><xdr:from>[\s\S]*?<\/xdr:oneCellAnchor>/.exec(after)[0];
  assert.match(turned, /<xdr:from><xdr:col>1<\/xdr:col><xdr:colOff>0<\/xdr:colOff><xdr:row>3<\/xdr:row><xdr:rowOff>9525<\/xdr:rowOff><\/xdr:from><xdr:ext cx="952500" cy="476250"\/>/);
  assert.match(turned, /<a:xfrm rot="1800000" flipH="1"><a:off x="609600" y="390525"\/><a:ext cx="952500" cy="476250"\/>/);
  // (Everything but that anchor as it was.)
  assert.equal(after.replace(turned, ''), before.replace(/<xdr:oneCellAnchor><xdr:from>[\s\S]*?<\/xdr:oneCellAnchor>/, ''));
  // (A two-cell anchor takes both corners.)
  setObjectAnchor(wb, 0, 4, { col: 2, colOff: 0, row: 2, rowOff: 0 }, { col: 4, colOff: 9525, row: 5, rowOff: 0 }, { x: 0, y: 0, w: 10, h: 10 });
  after = wb.pkg.text('xl/drawings/drawing1.xml');
  assert.match(after, /<xdr:to><xdr:col>4<\/xdr:col><xdr:colOff>9525<\/xdr:colOff><xdr:row>5<\/xdr:row><xdr:rowOff>0<\/xdr:rowOff><\/xdr:to><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="6" name="Rectangle 5"/);
  // (Deleted: the picture shared by three keeps its part until the last of them goes.)
  deleteObjects(wb, 0, [0, 1]);
  assert.ok(wb.pkg.has('xl/media/image1.png'));
  assert.equal(readObjects(wb, 0)[0].name, 'Turned');
  deleteObjects(wb, 0, [0, 1]);
  assert.ok(!wb.pkg.has('xl/media/image1.png'));
  assert.ok(!wb.pkg.has('xl/media/image2.png') && !wb.pkg.has('xl/media/image3.svg'));
  assert.deepEqual(wb.pkg.readRels('xl/drawings/drawing1.xml').map((r) => r.id), []);
  assert.equal(readObjects(wb, 0).length, 12);
});

test('anchors.xlsx: a chart counted among the objects, and deleted with its part', () => {
  const wb = book('anchors.xlsx');
  const list = readObjects(wb, 0);
  assert.deepEqual(list.map((o) => [o.k, o.kind, o.chart ?? null]), [[0, 'chart', 0], [1, 'shape', null]]);
  deleteObjects(wb, 0, [0]);
  assert.ok(!wb.pkg.has('xl/charts/chart1.xml'));
  assert.deepEqual(readObjects(wb, 0).map((o) => o.name), ['Rectangle 2']);
});
