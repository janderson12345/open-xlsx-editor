// Other people's cursors in a co-edited workbook (collab-peers.js gives them): each person's ranges lightly shaded and
// their active cell outlined in their colour, with their name on a small flag at the cell's top-right corner, as Excel
// for the web shows its co-authors. In an overlay layer of its own under this editor's selection, so that one's own
// cursor stays on top. View state only.
import { div } from './dom.js';
import { rangeBox } from './selection.js';

/** Draw st.peers ([{ name, color, ink, ar, ac, areas }]) into st.peerLayer. */
export function drawPeers(st) {
  if (!st.peerLayer) return;
  const kids = [];
  for (const p of st.peers || []) {
    const paint = (el) => { el.style.setProperty('--oxe-peer', p.color); el.style.setProperty('--oxe-peer-ink', p.ink); return el; };
    const single = p.areas.length === 1 && p.areas[0].r0 === p.areas[0].r1 && p.areas[0].c0 === p.areas[0].c1;
    if (!single) for (const a of p.areas) kids.push(paint(rangeBox(st, a.r0, a.c0, a.r1, a.c1, 'oxe-peer-range')));
    const m = st.sheet.merges.find((x) => p.ar >= x.r0 && p.ar <= x.r1 && p.ac >= x.c0 && p.ac <= x.c1);
    const cell = m || { r0: p.ar, c0: p.ac, r1: p.ar, c1: p.ac };
    const box = paint(rangeBox(st, cell.r0, cell.c0, cell.r1, cell.c1, 'oxe-peer-active'));
    box.dataset.peer = p.name;
    // (On the first row the flag hangs below the cell: above it would be under the column headings.)
    const flag = paint(div(cell.r0 === 0 ? 'oxe-peer-flag oxe-peer-flag-below' : 'oxe-peer-flag'));
    flag.textContent = p.name;
    box.append(flag);
    kids.push(box);
  }
  st.peerLayer.replaceChildren(...kids);
}
