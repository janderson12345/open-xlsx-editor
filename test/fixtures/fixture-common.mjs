// Shared helpers for the fixture builders: XML escaping, the package parts every workbook needs, and write().
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { zipSync, strToU8 } from 'fflate';

export const here = path.dirname(fileURLToPath(import.meta.url));
export const X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const RT = (t) => `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${t}`;
export const CT = (t) => `application/vnd.openxmlformats-officedocument.spreadsheetml.${t}+xml`;
export const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
export const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
export const col = (c) => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };

export const rels = (list) => `${DECL}<Relationships xmlns="${PR}">${list.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join('')}</Relationships>`;

export function contentTypes(overrides) {
  return `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    overrides.map(([part, type]) => `<Override PartName="/${part}" ContentType="${type}"/>`).join('') + '</Types>';
}

export const STYLES = `${DECL}<styleSheet xmlns="${X}">` +
  '<fonts count="2"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>' +
  '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
  '<dxfs count="1"><dxf><font><b/></font></dxf></dxfs><tableStyles count="0"/></styleSheet>';

export const docProps = () => ({
  'docProps/core.xml': `${DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>open-xlsx-editor fixtures</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created></cp:coreProperties>`,
  'docProps/app.xml': `${DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>open-xlsx-editor fixtures</Application></Properties>`,
});
export const PROPS_CT = [['docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'], ['docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml']];
export const ROOT_RELS = rels([['rId1', RT('officeDocument'), 'xl/workbook.xml'], ['rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml'], ['rId3', RT('extended-properties'), 'docProps/app.xml']]);

export function write(name, parts) {
  const files = {};
  for (const [k, v] of Object.entries(parts)) files[k] = [typeof v === 'string' ? strToU8(v) : v, { mtime: new Date('2026-01-01T00:00:00') }];
  fs.writeFileSync(path.join(here, name), zipSync(files, { level: 6 }));
  console.log(`wrote test/fixtures/${name}`);
}

// Shared strings collected while writing cells.
export function stringTable() {
  const list = [];
  const index = new Map();
  return {
    add(s) { if (!index.has(s)) { index.set(s, list.length); list.push(`<si><t${/^\s|\s$/.test(s) ? ' xml:space="preserve"' : ''}>${esc(s)}</t></si>`); } return index.get(s); },
    raw(xml) { list.push(xml); return list.length - 1; },
    xml: (count) => `${DECL}<sst xmlns="${X}" count="${count}" uniqueCount="${list.length}">${list.join('')}</sst>`,
  };
}

// A theme part of our own: Calibri for body text and headings, plain colours. (A workbook without one is drawn in
// Excel 365's default theme, Aptos Narrow, which browsers don't have.)
export const THEME = `${DECL}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="open-xlsx-editor fixtures"><a:themeElements>` +
  '<a:clrScheme name="fixtures">' + [['dk1', '000000'], ['lt1', 'FFFFFF'], ['dk2', '1F3864'], ['lt2', 'E7E6E6'], ['accent1', '4472C4'], ['accent2', 'ED7D31'],
  ['accent3', 'A5A5A5'], ['accent4', 'FFC000'], ['accent5', '5B9BD5'], ['accent6', '70AD47'], ['hlink', '0563C1'], ['folHlink', '954F72']]
    .map(([n, c]) => `<a:${n}><a:srgbClr val="${c}"/></a:${n}>`).join('') + '</a:clrScheme>' +
  '<a:fontScheme name="fixtures"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
  '<a:fmtScheme name="fixtures"><a:fillStyleLst>' + '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'.repeat(3) + '</a:fillStyleLst>' +
  '<a:lnStyleLst>' + '<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'.repeat(3) + '</a:lnStyleLst>' +
  '<a:effectStyleLst>' + '<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3) + '</a:effectStyleLst>' +
  '<a:bgFillStyleLst>' + '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'.repeat(3) + '</a:bgFillStyleLst></a:fmtScheme>' +
  '</a:themeElements></a:theme>';
