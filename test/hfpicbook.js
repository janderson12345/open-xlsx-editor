// Workbooks with pictures in their headers and footers, made in code for the tests and for Excel to print: a sheet of
// numbers, its headerFooter, and a VML drawing of pictures (legacyDrawingHF) as Excel writes one, the picture a PNG made
// here (a grey block in a black frame).
import { zlibSync, unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { csvWorkbook } from '../src/ooxml/csv.js';

const CRC = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
const crc32 = (bytes) => { let c = 0xffffffff; for (const b of bytes) c = CRC[(c ^ b) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
const u32 = (n) => [(n >>> 24) & 0xff, (n >>> 16) & 0xff, (n >>> 8) & 0xff, n & 0xff];

/** A w by h PNG (RGB): a grey block in a two-pixel black frame. */
export function blockPng(w = 80, h = 40) {
  const raw = new Uint8Array(h * (1 + 3 * w));
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const edge = x < 2 || y < 2 || x >= w - 2 || y >= h - 2;
      raw.fill(edge ? 0 : 160, y * (1 + 3 * w) + 1 + 3 * x, y * (1 + 3 * w) + 4 + 3 * x);
    }
  }
  const chunk = (type, data) => { const td = [...strToU8(type), ...data]; return [...u32(data.length), ...td, ...u32(crc32(td))]; };
  return new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, ...chunk('IHDR', [...u32(w), ...u32(h), 8, 2, 0, 0, 0]), ...chunk('IDAT', [...zlibSync(raw)]), ...chunk('IEND', [])]);
}

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;');

/**
 * A workbook of `rows` rows of numbers whose sheet has the headerFooter `hf` ({ oddHeader, oddFooter, firstHeader, ...,
 * differentFirst, differentOddEven }) and the pictures `pictures` ([{ id: 'LH' | 'CF' | 'RHFIRST' | ..., w, h }], points).
 */
export function hfPictureBook({ hf = {}, pictures = [], rows = 12 } = {}) {
  const csv = Array.from({ length: rows }, (_, r) => `${r + 1},${(r + 1) * 2},${(r + 1) * 3}`).join('\r\n');
  const files = unzipSync(csvWorkbook(strToU8(csv), 'Pics.csv', new Date(2026, 0, 15)).toBytes());
  const sheetPath = 'xl/worksheets/sheet1.xml';
  const flags = `${hf.differentOddEven ? ' differentOddEven="1"' : ''}${hf.differentFirst ? ' differentFirst="1"' : ''}`;
  const parts = ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter'].filter((k) => hf[k]).map((k) => `<${k}>${esc(hf[k])}</${k}>`).join('');
  let xml = strFromU8(files[sheetPath]).replace(/<headerFooter[\s\S]*?(<\/headerFooter>|\/>)/, '');
  const hfXml = `<headerFooter${flags}>${parts}</headerFooter>`;
  const legacy = pictures.length ? '<legacyDrawingHF r:id="rIdHF"/>' : '';
  // (After pageMargins and pageSetup, before anything else of CT_Worksheet's that follows them here.)
  const at = xml.search(/<(rowBreaks|colBreaks|drawing|legacyDrawing|tableParts|extLst)\b|<\/worksheet>/);
  xml = xml.slice(0, at) + hfXml + legacy + xml.slice(at);
  if (!/xmlns:r=/.test(xml)) xml = xml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
  files[sheetPath] = strToU8(xml);
  if (pictures.length) {
    const relsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
    const rel = '<Relationship Id="rIdHF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>';
    files[relsPath] = strToU8(files[relsPath] ? strFromU8(files[relsPath]).replace('</Relationships>', `${rel}</Relationships>`)
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel}</Relationships>`);
    const shapes = pictures.map((p, i) => `<v:shape id="${p.id}" o:spid="_x0000_s${1025 + i}" type="#_x0000_t75" style='position:absolute;margin-left:0;margin-top:0;width:${p.w}pt;height:${p.h}pt;z-index:${i + 1}'><v:imagedata o:relid="rId1" o:title="block"/><o:lock v:ext="edit" rotation="t"/></v:shape>`).join('');
    files['xl/drawings/vmlDrawing1.vml'] = strToU8(`<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>${shapes}</xml>`);
    files['xl/drawings/_rels/vmlDrawing1.vml.rels'] = strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>');
    files['xl/media/image1.png'] = blockPng();
    let types = strFromU8(files['[Content_Types].xml']);
    for (const [ext, type] of [['vml', 'application/vnd.openxmlformats-officedocument.vmlDrawing'], ['png', 'image/png']]) {
      if (!types.includes(`Extension="${ext}"`)) types = types.replace('<Default ', `<Default Extension="${ext}" ContentType="${type}"/><Default `);
    }
    files['[Content_Types].xml'] = strToU8(types);
  }
  return zipSync(files, { level: 6 });
}
