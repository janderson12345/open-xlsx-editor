// Probes for the print model: print area, titles and breaks; scale; fit to a page wide; headers and footers.
import { write, contentTypes, rels, RT, CT, DECL, X, R, ROOT_RELS, docProps, PROPS_CT, STYLES, THEME, col } from './fixture-common.mjs';

// Probes for Excel's print model, one workbook each (Excel prints the sheet a workbook opens on): a grid of labels
// "r1c1", "r1c2"... with our theme, and one feature of page setup each.
// - print-area.xlsx: a print area (B3:H70), rows 1:2 repeated at the top of each page, a manual break before row 30
//   and before column F.
// - print-scale.xlsx: printed at 60%.
// - print-fit.xlsx: fit to one page wide (as many tall as it takes), 24 columns.
// - print-headers.xlsx: a header (left text, "Page &P of &N" centred, the sheet's name right) and a footer (the file's
//   name centred), over two pages.
export function printProbes() {
  const grid = (rows, cols) => Array.from({ length: rows }, (_, r) => `<row r="${r + 1}">${Array.from({ length: cols }, (__, c) => `<c r="${col(c)}${r + 1}" t="inlineStr"><is><t>r${r + 1}c${c + 1}</t></is></c>`).join('')}</row>`).join('');
  const book = (name, { rows, cols, before = '', after = '', names = '' }) => write(name, {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')], ['xl/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'], ...PROPS_CT]),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Probe" sheetId="1" r:id="rId1"/></sheets>${names ? `<definedNames>${names}</definedNames>` : ''}</workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('styles'), 'styles.xml'], ['rId3', RT('theme'), 'theme/theme1.xml']]),
    'xl/worksheets/sheet1.xml': `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}">${before}<sheetData>${grid(rows, cols)}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${after}</worksheet>`,
    'xl/styles.xml': STYLES,
    'xl/theme/theme1.xml': THEME,
  });
  book('print-area.xlsx', {
    rows: 80, cols: 10,
    names: '<definedName name="_xlnm.Print_Area" localSheetId="0">Probe!$B$3:$H$70</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">Probe!$1:$2</definedName>',
    after: '<pageSetup orientation="portrait"/><rowBreaks count="1" manualBreakCount="1"><brk id="29" max="16383" man="1"/></rowBreaks><colBreaks count="1" manualBreakCount="1"><brk id="5" max="1048575" man="1"/></colBreaks>',
  });
  book('print-scale.xlsx', { rows: 90, cols: 14, after: '<pageSetup scale="60" orientation="portrait"/>' });
  book('print-fit.xlsx', { rows: 70, cols: 24, before: '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>', after: '<pageSetup fitToWidth="1" fitToHeight="0" orientation="portrait"/>' });
  book('print-headers.xlsx', {
    rows: 70, cols: 6,
    after: '<pageSetup orientation="portrait"/><headerFooter><oddHeader>&amp;LQuarterly stock&amp;CPage &amp;P of &amp;N&amp;R&amp;A</oddHeader><oddFooter>&amp;C&amp;F</oddFooter></headerFooter>',
  });
}
