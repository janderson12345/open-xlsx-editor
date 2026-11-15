// What kind of file a workbook is, and saving it as another: .xlsx, .xltx (template), .xlsm (macro-enabled), .xltm
// (macro-enabled template). Measured with Excel's SaveAs of one workbook in all four (formats 51, 54, 52, 53): the
// files differ only in the content type of the workbook part in [Content_Types].xml (and the documentId Excel makes
// new on every save). A new workbook made from a template, as opening one in Excel makes it ("book1", no path), saved
// as .xlsx is the same file as the workbook saved as .xlsx. So a template opens and edits as any workbook here, and
// saving as another kind changes that one content type, the rest of the file as it was.

export const FILE_TYPES = {
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  xltx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
  xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
  xltm: 'application/vnd.ms-excel.template.macroEnabled.main+xml',
};

const MACRO = new Set(['xlsm', 'xltm']);
const CT = '[Content_Types].xml';

/** The workbook part's Override element in [Content_Types].xml text, and its ContentType's value. */
function mainOverride(pkg) {
  const text = pkg.text(CT);
  const part = `/${pkg.mainPath}`;
  const re = /<Override\b[^>]*>/g;
  for (let m = re.exec(text); m; m = re.exec(text)) {
    const name = /\bPartName="([^"]*)"/.exec(m[0]);
    if (name && name[1].toLowerCase() === part.toLowerCase()) return { text, el: m[0], at: m.index, type: /\bContentType="([^"]*)"/.exec(m[0])?.[1] ?? null };
  }
  return { text, el: null, type: null };
}

/** The package's kind: 'xlsx', 'xltx', 'xlsm' or 'xltm' (null for another, as a .csv made into a workbook isn't). */
export function fileType(pkg) {
  const { type } = mainOverride(pkg);
  return Object.keys(FILE_TYPES).find((k) => FILE_TYPES[k] === type) ?? null;
}

/** Whether the package holds a VBA project (xl/vbaProject.bin, or wherever the workbook's relationship points). */
export const hasMacros = (pkg) => Object.keys(pkg.files).some((p) => /vbaProject\.bin$/i.test(p));

/** The package made a `type` file ('xlsx', 'xltx', 'xlsm', 'xltm'), only the workbook part's content type changed.
 *  A workbook with macros isn't made a macro-free kind: Excel would drop the macros (not measured here), so this
 *  refuses, as it does a type it doesn't know. */
export function setFileType(pkg, type) {
  if (!FILE_TYPES[type]) throw new Error(`Not a workbook type: ${type}`);
  if (!MACRO.has(type) && hasMacros(pkg)) throw new Error(`This workbook has macros, which a .${type} file can't keep. Save it as .${type === 'xltx' ? 'xltm' : 'xlsm'}.`);
  const { text, el, at } = mainOverride(pkg);
  if (!el) throw new Error('The workbook part has no content type to change');
  const next = el.replace(/\bContentType="[^"]*"/, `ContentType="${FILE_TYPES[type]}"`);
  if (next !== el) pkg.setText(CT, text.slice(0, at) + next + text.slice(at + el.length));
}
