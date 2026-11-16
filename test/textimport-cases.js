// Made-up text files and Text Import settings, opened by Excel's Text Import Wizard over COM
// (tools/excel-compare/excel-textimport.ps1, giving test/textimport-excel.json) and by ours (test/textimport.test.js).
// Settings in ooxml/textimport.js's terms: kind 'delimited' or 'fixed', seps (the delimiter characters), quote ('"',
// "'" or ''), consecutive, breaks (fixed width: where columns start), startRow (1 for the first), columns (per column:
// 'general', 'text', 'MDY', 'DMY', 'YMD' or 'skip'), decimal and group (the number separators), encoding.
const TAB = String.fromCharCode(9);
const lines = (...rows) => rows.join('\r\n') + '\r\n';

export const TEXT_IMPORT_CASES = [
  {
    name: 'semicolon-de', encoding: 'utf-8', seps: ';', quote: '"', decimal: ',', group: '.', columns: ['general', 'general', 'DMY'],
    text: lines('item;amount;when', 'alpha;1.234,5;05.03.2024', 'beta;-2,25;13.12.2023', 'gamma;"x;y";1.2.24', 'delta;3,5%;31/1/2024', 'eps;1.5;5.3'),
  },
  {
    name: 'tab-us', encoding: 'utf-8', seps: TAB, quote: '"', columns: ['text', 'general', 'skip', 'MDY'],
    text: lines(['code', 'value', 'drop', 'date'].join(TAB), ['007', '1,234.50', 'x', '3/5/2024'].join(TAB), ['0042', '$12', 'y', '12/31/2023'].join(TAB),
      ['A-1', '50%', 'z', '2024-03-05'].join(TAB), ['B 2', '(7)', 'w', 'not a date'].join(TAB)),
  },
  {
    name: 'pipe-single-quote', encoding: 'utf-8', seps: '|', quote: "'", columns: ['general', 'general'],
    text: lines("name|note", "'a|b'|plain", "'it''s'|'quoted'", 'x"y|"q"'),
  },
  {
    name: 'comma-no-quote', encoding: 'utf-8', seps: ',', quote: '', columns: ['general', 'general', 'general'],
    text: lines('a,"b,c",d', '1,2,3'),
  },
  {
    name: 'consecutive-spaces', encoding: 'utf-8', seps: ' ', quote: '"', consecutive: true, columns: ['general', 'general', 'general'],
    text: lines('one   two  three', ' 1  2   3', '"a b" c d'),
  },
  {
    name: 'fixed-width', encoding: 'utf-8', kind: 'fixed', breaks: [6, 14], columns: ['general', 'text', 'general'],
    text: lines('ID    NAME    AMOUNT', '00012 Widget    15.25', '00013 Gear      1,000', '7     Bolt     -3'),
  },
  {
    name: 'start-row-3', encoding: 'utf-8', seps: ',', quote: '"', startRow: 3, columns: ['general', 'general'],
    text: lines('report title', 'made up', 'k,v', 'a,1', 'b,2'),
  },
  {
    name: 'utf16', encoding: 'utf-16le', bom: true, seps: TAB, quote: '"', columns: ['general', 'general'],
    text: lines(['word', 'n'].join(TAB), ['caf' + String.fromCharCode(0xe9), '1'].join(TAB), [String.fromCharCode(0x65e5, 0x672c), '2'].join(TAB)),
  },
  {
    name: 'ansi', encoding: 'windows-1252', seps: ',', quote: '"', columns: ['general', 'general'],
    text: lines('word,n', 'caf' + String.fromCharCode(0xe9) + ',1', String.fromCharCode(0x20ac) + '5,2'),
  },
  {
    name: 'ymd-dates', encoding: 'utf-8', seps: ',', quote: '"', columns: ['YMD', 'DMY', 'MDY'],
    text: lines('2024-03-05,05/03/2024,03/05/2024', '24/3/5,5-3-24,3-5-24', '2024.3.5,5.3.2024,13/5/2024', 'x,31/2/2024,2/30/2024'),
  },
  {
    name: 'general-us', encoding: 'utf-8', seps: ',', quote: '"', columns: ['general', 'general', 'general'],
    text: lines('=1+2,TRUE,#N/A', '1e3,"1,234",10:30', '1/2,  pad  ,-'),
  },
];

/** A case's file as bytes, in its encoding (UTF-16 with its byte order mark when `bom`). */
export function caseBytes(c) {
  if (c.encoding === 'utf-16le') {
    const out = [];
    if (c.bom) out.push(0xff, 0xfe);
    for (let i = 0; i < c.text.length; i++) { const u = c.text.charCodeAt(i); out.push(u & 255, u >> 8); }
    return Uint8Array.from(out);
  }
  if (c.encoding === 'windows-1252') return Uint8Array.from([...c.text].map((ch) => (ch === String.fromCharCode(0x20ac) ? 0x80 : ch.charCodeAt(0))));
  return new TextEncoder().encode(c.text);
}
