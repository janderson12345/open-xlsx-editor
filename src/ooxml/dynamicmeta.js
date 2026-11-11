// The cell metadata that marks a formula as Excel 365's dynamic array formula (the cell's cm, pointing into
// xl/metadata.xml's XLDAPR entry), as Excel writes it (measured: tools/excel-compare/excel-dynamic.ps1). A workbook
// without the part gets Excel's own; one with it (rich values, other metadata) gets the XLDAPR entry added.
import { kids, kid, attr, parseXml, serializeXml } from './xml.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata';
const CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml';
const XDA = 'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray';
const EXT = '{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}';
const TYPE_ATTRS = 'minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"';
const FRESH = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="${XDA}"><metadataTypes count="1"><metadataType name="XLDAPR" ${TYPE_ATTRS}/></metadataTypes>` +
  `<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="${EXT}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>` +
  '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';

/** The metadata part's path, or null. */
function partOf(wb) {
  return wb.pkg.rels.find((r) => r.type === REL)?.target ?? null;
}

/** The cm index (1-based) of the dynamic array entry, the part made or added to when it has none. */
export function dynamicCm(wb) {
  if (wb._dynamicCm) return wb._dynamicCm;
  let path = partOf(wb);
  if (!path) {
    path = `${wb.pkg.mainDir}metadata.xml`;
    wb.pkg.setText(path, FRESH);
    wb.pkg.addRel(REL, path.slice(wb.pkg.mainDir.length));
    wb.pkg.addOverride(path, CT);
    wb._dynamicCm = 1;
    return 1;
  }
  wb._dynamicCm = addEntry(wb, path);
  return wb._dynamicCm;
}

/** The XLDAPR entry found or added in an existing part: its cell metadata block's 1-based index. */
function addEntry(wb, path) {
  const doc = parseXml(wb.pkg.text(path));
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const child = (parent, name, before = []) => {
    let el = kid(parent, name, ns);
    if (el) return el;
    el = doc.createElementNS(ns, name);
    const next = [...parent.childNodes].find((n) => n.nodeType === 1 && before.includes(n.localName));
    parent.insertBefore(el, next || null);
    return el;
  };
  const types = child(root, 'metadataTypes', ['metadataStrings', 'mdxMetadata', 'futureMetadata', 'cellMetadata', 'valueMetadata', 'extLst']);
  let t = kids(types, 'metadataType', ns).findIndex((x) => attr(x, 'name') === 'XLDAPR') + 1;
  if (!t) {
    const frag = parseXml(`<x xmlns="${ns}"><metadataType name="XLDAPR" ${TYPE_ATTRS}/></x>`).documentElement.firstChild;
    types.appendChild(doc.importNode(frag, true));
    t = kids(types, 'metadataType', ns).length;
    types.setAttribute('count', String(t));
  }
  let future = kids(root, 'futureMetadata', ns).find((x) => attr(x, 'name') === 'XLDAPR');
  if (!future) {
    const frag = parseXml(`<x xmlns="${ns}" xmlns:xda="${XDA}"><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="${EXT}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata></x>`).documentElement.firstChild;
    const next = [...root.childNodes].find((n) => n.nodeType === 1 && ['cellMetadata', 'valueMetadata', 'extLst'].includes(n.localName));
    future = root.insertBefore(doc.importNode(frag, true), next || null);
  }
  const cells = child(root, 'cellMetadata', ['valueMetadata', 'extLst']);
  const blocks = kids(cells, 'bk', ns);
  const found = blocks.findIndex((bk) => kids(bk, 'rc', ns).some((rc) => Number(attr(rc, 't')) === t && attr(rc, 'v') === '0'));
  if (found >= 0) return found + 1;
  const bk = doc.createElementNS(ns, 'bk');
  const rc = doc.createElementNS(ns, 'rc');
  rc.setAttribute('t', String(t));
  rc.setAttribute('v', '0');
  bk.appendChild(rc);
  cells.appendChild(bk);
  cells.setAttribute('count', String(blocks.length + 1));
  wb.pkg.setText(path, serializeXml(doc));
  return blocks.length + 1;
}
