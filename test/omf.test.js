import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { OMF_VERSION, OMF_FACES } from '../src/editor/omf.js';
import { addOmfFaces } from '../src/editor/omffonts.js';

// The faces come from open-metric-fonts' morrow-sans package (tools/omf-faces.mjs): Aptos Narrow's twin and Aptos's,
// regular, italic, bold and bold italic, each in the package's alphabets.
test('the faces of Morrow Sans Narrow and Morrow Sans, each alphabet a file and its range', () => {
  assert.match(OMF_VERSION, /^\d+\.\d+\.\d+$/);
  for (const family of ['Morrow Sans Narrow', 'Morrow Sans']) {
    for (const [weight, style] of [[400, 'normal'], [400, 'italic'], [700, 'normal'], [700, 'italic']]) {
      const faces = OMF_FACES.filter((f) => f[0] === family && f[1] === weight && f[2] === style);
      assert.ok(faces.some((f) => /-latin\.woff2$/.test(f[3]) && f[4].startsWith('U+0000-00FF')), `${family} ${weight} ${style}: latin`);
      for (const f of faces) assert.match(f[4], /^U\+[0-9A-F]+(-[0-9A-F]+)?(,U\+[0-9A-F]+(-[0-9A-F]+)?)*$/);
    }
  }
});

test('added to document.fonts from the folder given, once each, each with its range', () => {
  const fonts = [];
  const saved = { document: globalThis.document, FontFace: globalThis.FontFace };
  globalThis.document = { fonts: { add: (f) => fonts.push(f) }, baseURI: 'https://example.test/app/page.html' };
  globalThis.FontFace = class { constructor(family, src, opts) { Object.assign(this, { family, src, ...opts }); } };
  try {
    // (Relative to the page, as a site serving one copy for every viewer points at it: /lib/fonts/omf/<version>/.)
    assert.equal(addOmfFaces(`/lib/fonts/omf/${OMF_VERSION}`), OMF_FACES.length);
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
