// The cases excel-richruns.ps1 has Excel do (written to richruns-cases.json): Format Cells' Font tab on part of a
// cell's text (`part`: a source cell, its characters' fonts set as the box sets them), Copy and Paste of a cell of rich
// text (`paste`: `how` cut, paste (Ctrl+V) or otherBook, else Paste Special's `paste` number, xlPasteAll -4104,
// xlPasteFormats -4122, xlPasteValues -4163, xlPasteFormulas -4123, xlPasteAllExceptBorders 7, with `transpose`), and
// another program's HTML with mixed formatting in one cell (`html`). The answers: test/richruns-excel.json.
// usage: node tools/excel-compare/richruns-cases.mjs
import { writeFileSync } from 'node:fs';

const TEXT = 'Hello brave world';
// ("brave", characters 7 to 11.)
const part = (font) => ({ text: TEXT, parts: [{ start: 7, len: 5, font }] });
const rich = { text: TEXT, parts: [{ start: 7, len: 5, font: { Bold: true, color: 'FF0000' } }] };
const richStart = { text: TEXT, parts: [{ start: 1, len: 5, font: { Italic: true } }, { start: 13, len: 5, font: { Size: 14 } }] };

const cases = {
  part: [
    { name: 'bold italic', src: part({ FontStyle: 'Bold Italic' }) },
    { name: 'superscript', src: part({ Superscript: true }) },
    { name: 'subscript', src: part({ Subscript: true }) },
    { name: 'superscript then subscript', src: { text: TEXT, parts: [{ start: 7, len: 5, font: { Superscript: true } }, { start: 7, len: 5, font: { Subscript: true } }] } },
    { name: 'strikethrough', src: part({ Strikethrough: true }) },
    { name: 'double underline', src: part({ Underline: -4119 }) },
    { name: 'single accounting underline', src: part({ Underline: 4 }) },
    { name: 'double accounting underline', src: part({ Underline: 5 }) },
    { name: 'colour', src: part({ color: '00B050' }) },
    { name: 'size', src: part({ Size: 16 }) },
    { name: 'font name', src: part({ Name: 'Courier New' }) },
    { name: 'all at once', src: part({ Name: 'Georgia', Size: 9, FontStyle: 'Italic', Underline: 2, Strikethrough: true, color: '0070C0', Superscript: true }) },
    { name: 'the whole text', src: { text: TEXT, parts: [{ start: 1, len: 17, font: { Bold: true } }] } },
  ],
  paste: [
    { name: 'all', src: rich, paste: -4104 },
    { name: 'ctrl+v', src: rich, how: 'paste' },
    { name: 'all over text', src: rich, dst: { text: 'Plain text here' }, paste: -4104 },
    { name: 'formats over longer text', src: rich, dst: { text: 'Some other plain text' }, paste: -4122 },
    { name: 'formats over shorter text', src: rich, dst: { text: 'Hi there' }, paste: -4122 },
    { name: 'formats over the same text', src: rich, dst: { text: TEXT }, paste: -4122 },
    { name: 'formats over a number', src: rich, dst: { number: 42 }, paste: -4122 },
    { name: 'formats over empty', src: rich, paste: -4122 },
    { name: 'formats from runs at start and end', src: richStart, dst: { text: 'Some other plain text' }, paste: -4122 },
    { name: 'values', src: rich, paste: -4163 },
    { name: 'formulas', src: rich, paste: -4123 },
    { name: 'all except borders', src: rich, paste: 7 },
    { name: 'values and number formats', src: rich, paste: 12 },
    { name: 'all transposed', src: rich, paste: -4104, transpose: true },
    { name: 'cut', src: rich, how: 'cut' },
    { name: 'to another workbook', src: rich, how: 'otherBook' },
    { name: 'all from runs at start and end', src: richStart, paste: -4104 },
    { name: 'formats from plain over rich', src: { text: 'Plain', parts: [{ start: 1, len: 5, font: { Underline: 2 } }] }, dst: rich, paste: -4122 },
    { name: 'formats from a number over rich', src: { number: 7 }, dst: rich, paste: -4122 },
    { name: 'superscript run', src: part({ Superscript: true, Size: 8 }), paste: -4104 },
  ],
  html: [
    { name: 'sup and sub', html: '<html><body><!--StartFragment--><table><tr><td>H<sub>2</sub>O and x<sup>2</sup></td></tr></table><!--EndFragment--></body></html>', text: 'H2O and x2\r\n' },
    { name: 'u, s, i, b', html: '<html><body><!--StartFragment--><table><tr><td>one <u>two</u> <s>three</s> <i>four</i> <b>five</b></td></tr></table><!--EndFragment--></body></html>', text: 'one two three four five\r\n' },
    { name: 'colour and size spans', html: '<html><body><!--StartFragment--><table><tr><td>red <span style="color:#FF0000">here</span> and <span style="font-size:16pt">big</span></td></tr></table><!--EndFragment--></body></html>', text: 'red here and big\r\n' },
    { name: 'font family span', html: '<html><body><!--StartFragment--><table><tr><td style="font-family:Georgia">mono <span style="font-family:\'Courier New\'">code</span> end</td></tr></table><!--EndFragment--></body></html>', text: 'mono code end\r\n' },
    { name: 'vertical-align css', html: '<html><body><!--StartFragment--><table><tr><td>E=mc<span style="vertical-align:super">2</span> and a<span style="vertical-align:sub">i</span></td></tr></table><!--EndFragment--></body></html>', text: 'E=mc2 and ai\r\n' },
    { name: 'word paragraph', html: "<html xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w=\"urn:schemas-microsoft-com:office:word\"><head><style><!-- p.MsoNormal {margin:0in; font-size:12.0pt; font-family:\"Aptos\",sans-serif;} --></style></head><body><!--StartFragment--><table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0><tr><td width=200 valign=top><p class=MsoNormal>Word <b>bold</b> and <span style='color:#C00000'>red</span> and <u>under</u></p></td></tr></table><!--EndFragment--></body></html>", text: 'Word bold and red and under\r\n' },
    { name: 'double underline css', html: '<html><body><!--StartFragment--><table><tr><td>a <span style="text-decoration:underline;text-underline-style:double">dbl</span> b</td></tr></table><!--EndFragment--></body></html>', text: 'a dbl b\r\n' },
    { name: 'all one format', html: '<html><body><!--StartFragment--><table><tr><td><b>all</b> <b>bold</b></td></tr></table><!--EndFragment--></body></html>', text: 'all bold\r\n' },
  ],
};
writeFileSync(new URL('./richruns-cases.json', import.meta.url), `${JSON.stringify(cases, null, 2)}\n`);
