// Style sheets the editor writes itself (a workbook's cell formats, the paper size), made through the CSSOM: a
// constructed CSSStyleSheet adopted by the document. A Content-Security-Policy of style-src 'self' lets a page's
// script do this, as it doesn't let it add a <style> element or a style="" attribute (measured in Edge, README:
// Self-hosting), so the editor runs under such a policy with nothing inline allowed.

const sheets = new WeakMap();

/** Where `owner`'s sheet goes: its shadow root when it is in one, else its document. */
const holderOf = (owner) => {
  const top = owner.getRootNode?.();
  return top && 'adoptedStyleSheets' in top ? top : document;
};

/** Sets the CSS of `owner`'s own style sheet (an element, or the document), made and adopted the first time. */
export function setSheet(owner, cssText) {
  let entry = sheets.get(owner);
  if (!entry) {
    const sheet = new CSSStyleSheet();
    const holder = holderOf(owner);
    holder.adoptedStyleSheets = [...holder.adoptedStyleSheets, sheet];
    entry = { sheet, holder };
    sheets.set(owner, entry);
  }
  entry.sheet.replaceSync(cssText);
  return entry.sheet;
}

/** Removes `owner`'s style sheet, if it has one. */
export function dropSheet(owner) {
  const entry = sheets.get(owner);
  if (!entry) return;
  entry.holder.adoptedStyleSheets = entry.holder.adoptedStyleSheets.filter((s) => s !== entry.sheet);
  sheets.delete(owner);
}
