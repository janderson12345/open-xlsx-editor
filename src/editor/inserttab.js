// The Insert tab, Excel's groups as far as they are here: Tables (PivotTable: its boxes, pivotbox.js, and the
// PivotTable made, pivotnew.js; Table: the Create Table box, acts-tables.js), Charts (a chart of the selection, chartnew.js, by Excel's menus of each kind) and
// Sparklines (Line, Column, Win/Loss: the Create Sparklines box, acts-sparklines.js) and Links (Insert Hyperlink,
// linkbox.js).

/** Excel's Charts group: each menu's kinds, by the names Excel gives them (chartxml.js CHART_KINDS). */
export const CHART_MENUS = [
  ['▥', 'Insert Column or Bar Chart', 'Column', [['columnClustered', 'Clustered Column'], ['columnStacked', 'Stacked Column'], ['columnStacked100', '100% Stacked Column'],
    ['barClustered', 'Clustered Bar'], ['barStacked', 'Stacked Bar'], ['barStacked100', '100% Stacked Bar']]],
  ['📈', 'Insert Line or Area Chart', 'Line', [['line', 'Line'], ['lineMarkers', 'Line with Markers'], ['area', 'Area'], ['areaStacked', 'Stacked Area']]],
  ['◔', 'Insert Pie or Doughnut Chart', 'Pie', [['pie', 'Pie'], ['doughnut', 'Doughnut']]],
  ['⁘', 'Insert Scatter (X, Y) Chart', 'Scatter', [['scatter', 'Scatter'], ['scatterLines', 'Scatter with Straight Lines and Markers'], ['scatterSmooth', 'Scatter with Smooth Lines and Markers']]],
];

/** The Insert tab. `act(name, ...args)` runs a command. */
export function insertTab(kit, { act }) {
  const { grp, btn, menu, item } = kit;
  const groups = [
    grp('Tables', btn('▦', 'PivotTable', () => act('pivotBox'), { label: 'PivotTable', big: true }),
      btn('▤', 'Table (Ctrl+T)', () => act('tableInsert'), { label: 'Table', big: true })),
    grp('Charts', ...CHART_MENUS.map(([icon, title, label, kinds]) => menu(icon, title, () => kinds.map(([kind, name]) => item(name, () => act('insertChart', kind))), { label, big: true }))),
    grp('Sparklines', btn('📈', 'Insert Line Sparkline', () => act('sparklineInsert', 'line'), { label: 'Line', big: true }),
      btn('▥', 'Insert Column Sparkline', () => act('sparklineInsert', 'column'), { label: 'Column', big: true }),
      btn('◫', 'Insert Win/Loss Sparkline', () => act('sparklineInsert', 'stacked'), { label: 'Win/Loss', big: true })),
    grp('Links', btn('🔗', 'Link (Ctrl+K)', () => act('linkBox'), { label: 'Link', big: true })),
  ];
  return { groups, update() {} };
}
