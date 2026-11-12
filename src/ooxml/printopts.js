// Page Layout > Sheet Options' Print boxes as Excel writes them (measured over COM, PageSetup.PrintGridlines and
// PrintHeadings on a new workbook, and the files Excel saved: docs/VIEW.md): <printOptions headings="1" gridLines="1"/>
// before pageMargins, each attribute only while its box is ticked, the element gone when neither is.
import { kid } from './xml.js';
import { AFTER_MERGES } from './sheetedit.js';

// (printOptions comes before these parts of a worksheet.)
const AFTER = AFTER_MERGES.slice(AFTER_MERGES.indexOf('printOptions') + 1);

/** Sheet si's Print Gridlines (`gridLines`) and Print Headings (`headings`) set, either or both. */
export function setPrintOptions(wb, si, { gridLines, headings }) {
  const part = wb._part(si);
  const { doc, model } = part;
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  let el = kid(root, 'printOptions', ns);
  if (!el) {
    el = doc.createElementNS(ns, root.prefix ? `${root.prefix}:printOptions` : 'printOptions');
    root.insertBefore(el, [...root.childNodes].find((n) => n.nodeType === 1 && AFTER.includes(n.localName)) || null);
  }
  const set = (name, on) => { if (on) el.setAttribute(name, '1'); else el.removeAttribute(name); };
  // (Excel writes headings first.)
  if (headings != null) { set('headings', headings); model.page.headings = !!headings; }
  if (gridLines != null) { set('gridLines', gridLines); model.page.gridLines = !!gridLines; }
  if (!el.attributes.length) root.removeChild(el);
  wb.dirty.add(wb.sheets[si].path);
  wb.edited = true;
}
