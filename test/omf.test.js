import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { OMF_VERSION, OMF, OMF_RANGES } from '../src/editor/omf.js';
import { useOmf, omfFamily, omfFace } from '../src/editor/omffonts.js';
import { fontStack } from '../src/editor/fonts.js';

// The map comes from open-metric-fonts' packages (tools/omf-map.py, ported from open-docx-editor): Aptos Narrow's twin
// and Aptos's, regular, italic, bold and bold italic, each in the package's alphabets.
const faceCount = (name) => OMF[name][2].reduce((n, f) => n + f[3].split(',').length, 0);
test('the faces of the twins of Aptos Narrow and Aptos, each alphabet a file and its range', () => {
  assert.match(OMF_VERSION, /^\d+\.\d+\.\d+$/);
  for (const name of ['Aptos Narrow', 'Aptos']) {
    const [pkg, , faces] = OMF[name];
    assert.equal(pkg, 'morrow-sans');
    assert.deepEqual(faces.map((f) => [f[0], f[1]]), [[400, 0], [400, 1], [700, 0], [700, 1]]);
    for (const [, , , alphabets] of faces) {
      assert.ok(alphabets.split(',').includes('latin'));
      for (const a of alphabets.split(',')) assert.match(OMF_RANGES[`${pkg}:${a}`] || OMF_RANGES[a], /^U\+[0-9A-F]+(-[0-9A-F]+)?(,U\+[0-9A-F]+(-[0-9A-F]+)?)*$/);
    }
  }
  assert.ok(OMF_RANGES.latin.startsWith('U+0000-00FF'));
});

test('every family of the release, its faces as Office links them', () => {
  assert.equal(Object.keys(OMF).length, 151);
  for (const [name, [pkg, m, faces]] of Object.entries(OMF)) {
    assert.match(pkg, /^[a-z-]+$/, name);
    assert.equal(m.win.length, 2, name);
    assert.ok(m.win[0] > 0.5 && m.win[0] + m.win[1] < 2.5, name);
    assert.ok(faces.length >= 1 && faces.every(([w, i]) => [400, 700].includes(w) && [0, 1].includes(i)), name);
  }
  // (The Windows metrics of the originals, as engine/metrics.js read them from the installed fonts.)
  assert.deepEqual(OMF.Tahoma[1].win, [1.0005, 0.2065]);
  assert.deepEqual(OMF.Verdana[1].win, [1.0054, 0.21]);
  assert.deepEqual(OMF['Century Gothic'][1].win, [0.9712, 0.2202]);
  assert.deepEqual(OMF.Garamond[1].win, [0.8618, 0.2632]);
});

test('the face drawn for bold and italic, as CSS matching picks it', () => {
  assert.deepEqual(omfFace('verdana', true, true).slice(0, 2), ['glade', 'glade-bolditalic']);
  assert.deepEqual(omfFace('Verdana', false, false).slice(0, 2), ['glade', 'glade']);
  // (Tahoma has no italic: its regular and bold, slanted by the browser.)
  assert.deepEqual(omfFace('Tahoma', true, true).slice(0, 2), ['tamarack', 'tamarack-bold']);
  assert.deepEqual(omfFace('Tahoma', false, true).slice(0, 2), ['tamarack', 'tamarack']);
  assert.equal(omfFace('Calibri', false, false), null);
});

test('the faces of a family added to document.fonts the first time it is named, each alphabet with its range', () => {
  const fonts = [];
  const saved = { document: globalThis.document, FontFace: globalThis.FontFace };
  globalThis.document = { fonts: { add: (f) => fonts.push(f) }, baseURI: 'https://example.test/app/page.html' };
  globalThis.FontFace = class { constructor(family, src, opts) { Object.assign(this, { family, src, ...opts }); } };
  try {
    assert.equal(omfFamily('Tahoma'), null, 'no folder: no twins');
    // (Relative to the page, as a site serving one copy for every viewer points at it: /lib/fonts/omf/<version>/.)
    assert.equal(useOmf(`/lib/fonts/omf/${OMF_VERSION}`), `https://example.test/lib/fonts/omf/${OMF_VERSION}/`);
    assert.equal(fonts.length, 0, 'nothing added before a workbook names a font');
    assert.equal(fontStack('Aptos Narrow'), "'Aptos Narrow', 'oxe-omf Aptos Narrow', 'Morrow Sans Narrow', Calibri, Carlito, sans-serif");
    assert.equal(fonts.length, faceCount('Aptos Narrow'));
    assert.equal(fontStack('tahoma'), "tahoma, 'oxe-omf Tahoma', sans-serif");
    assert.equal(fontStack('Tahoma'), "Tahoma, 'oxe-omf Tahoma', sans-serif");
    assert.equal(fonts.length, faceCount('Aptos Narrow') + faceCount('Tahoma'), 'once per family');
    assert.equal(fontStack('Calibri'), 'Calibri, Carlito, sans-serif');
    const narrow = fonts.find((f) => f.family === 'oxe-omf Aptos Narrow' && f.weight === '400' && f.style === 'normal' && /-latin\.woff2/.test(f.src));
    assert.equal(narrow.src, `url("https://example.test/lib/fonts/omf/${OMF_VERSION}/morrow-sans/files/morrowsans-narrow-latin.woff2") format("woff2")`);
    assert.ok(narrow.unicodeRange.startsWith('U+0000-00FF'));
    assert.equal(narrow.display, 'swap');
    // (A package's own "other" alphabet: its range, not the shared one.)
    const other = fonts.find((f) => f.family === 'oxe-omf Tahoma' && /tamarack-other\.woff2/.test(f.src));
    assert.equal(other.unicodeRange, OMF_RANGES['tamarack:other']);
  } finally {
    useOmf('');
    globalThis.document = saved.document;
    globalThis.FontFace = saved.FontFace;
  }
});
