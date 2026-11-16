// Which ribbon group each command belongs to, for mountXlsxEditor's `features` (index.js).
/** The ribbon group a formatting command belongs to, for `features`: one of a group turned off doesn't run, from the
 *  ribbon or the keyboard. */
const FEATURE_OF = { toggleFont: 'font', stepFont: 'font', borders: 'font', drawBorders: 'font', moreColor: 'font', toggleWrap: 'alignment', stepIndent: 'alignment', align: 'alignment', merge: 'alignment',
  resizeLines: 'cells', columnWidth: 'cells', rowHeight: 'cells', fitLines: 'cells', hideLines: 'cells', rowHeightBox: 'cells', columnWidthBox: 'cells',
  freeze: 'window', zoom: 'zoom', zoomBox: 'zoom', zoomSelection: 'zoom', zoomToSelection: 'zoom',
  lines: 'cells', cells: 'cells', insertCells: 'cells', deleteCells: 'cells', insertRows: 'cells', deleteRows: 'cells', insertColumns: 'cells', deleteColumns: 'cells',
  clipCopy: 'clipboard', clipCut: 'clipboard', clipPaste: 'clipboard', copy: 'clipboard', cut: 'clipboard', paste: 'clipboard',
  pasteSpecial: 'clipboard', pasteSpecialBox: 'clipboard', pasteAs: 'clipboard', formatPainter: 'clipboard', paintFormats: 'clipboard',
  formatCells: 'font', formatCellsBox: 'font',
  stepDecimal: 'number', clearFormats: 'editing', clearAs: 'editing', fill: 'editing', autoFill: 'editing',
  sort: 'editing', sortAsc: 'editing', sortDesc: 'editing', sortBox: 'editing', autoSum: 'editing',
  filterToggle: 'editing', filterBy: 'editing', filterClear: 'editing', filterReapply: 'editing', filterAll: 'editing', filterSort: 'editing',
  formatAsTable: 'styles', newCellStyle: 'styles', modifyCellStyle: 'styles', newCellStyleBox: 'styles', modifyCellStyleBox: 'styles', tableInsert: 'tables' };
export const groupOf = (name, change) => (name !== 'format' ? FEATURE_OF[name] : change.cellStyle ? 'styles' : change.numFmt != null || change.style ? 'number' : change.align ? 'alignment' : 'font');
