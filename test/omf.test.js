import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { OMF_VERSION, OMF, OMF_RANGES } from '../src/editor/omf.js';
import { addOmfFaces } from '../src/editor/omffonts.js';

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

test('added to document.fonts from the folder given, once each, each with its range', () => {
  const fonts = [];
  const saved = { document: globalThis.document, FontFace: globalThis.FontFace };
  globalThis.document = { fonts: { add: (f) => fonts.push(f) }, baseURI: 'https://example.test/app/page.html' };
  globalThis.FontFace = class { constructor(family, src, opts) { Object.assign(this, { family, src, ...opts }); } };
  try {
    // (Relative to the page, as a site serving one copy for every viewer points at it: /lib/fonts/omf/<version>/.)
    assert.equal(addOmfFaces(`/lib/fonts/omf/${OMF_VERSION}`), faceCount('Aptos Narrow') + faceCount('Aptos'));
    assert.equal(addOmfFaces(`/lib/fonts/omf/${OMF_VERSION}/`), 0);
    const narrow = fonts.find((f) => f.family === 'Morrow Sans Narrow' && f.weight === '400' && f.style === 'normal' && /-latin\.woff2/.test(f.src));
    assert.equal(narrow.src, `url("https://example.test/lib/fonts/omf/${OMF_VERSION}/morrow-sans/files/morrowsans-narrow-latin.woff2") format("woff2")`);
    assert.ok(narrow.unicodeRange.startsWith('U+0000-00FF'));
    assert.equal(narrow.display, 'swap');
  } finally {
    globalThis.document = saved.document;
    globalThis.FontFace = saved.FontFace;
  }
});
