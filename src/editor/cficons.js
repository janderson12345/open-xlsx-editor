// Icon sets' icons, drawn as small SVGs in the look of Excel's (not its own pictures): arrows, traffic lights, signs,
// symbols, flags, ratings, quarters, stars, triangles and boxes; index 0 is a set's lowest.
const RED = '#C00000';
const YELLOW = '#F4B400';
const GREEN = '#00A651';
const GREY = '#808080';
const BLACK = '#262626';
const GOLD = '#F4B400';

const arrow = (angle, fill) => `<path d="M6 1 L11 6 L8 6 L8 11 L4 11 L4 6 L1 6 Z" fill="${fill}" transform="rotate(${angle} 6 6)"/>`;
const dot = (fill) => `<circle cx="6" cy="6" r="5" fill="${fill}"/>`;
const ring = (fill, mark) => `<circle cx="6" cy="6" r="5" fill="${fill}"/>${mark}`;
const tick = '<path d="M3 6.2 L5.2 8.4 L9 4" stroke="#fff" stroke-width="1.6" fill="none"/>';
const cross = '<path d="M3.8 3.8 L8.2 8.2 M8.2 3.8 L3.8 8.2" stroke="#fff" stroke-width="1.6"/>';
const bang = '<path d="M6 2.6 V7" stroke="#fff" stroke-width="1.6"/><circle cx="6" cy="9" r="0.9" fill="#fff"/>';
const flag = (fill) => `<path d="M2.5 1 V11" stroke="${GREY}" stroke-width="1"/><path d="M3 1.5 H10 L8 4 L10 6.5 H3 Z" fill="${fill}"/>`;
const bars = (n, of) => Array.from({ length: of }, (_, i) => `<rect x="${1 + i * (10 / of)}" y="${10 - (i + 1) * (9 / of)}" width="${10 / of - 1}" height="${(i + 1) * (9 / of)}" fill="${i < n ? '#2F5597' : '#D9D9D9'}"/>`).join('');
const quarter = (k) => {
  if (k === 0) return '<circle cx="6" cy="6" r="5" fill="#fff" stroke="#404040"/>';
  if (k === 4) return '<circle cx="6" cy="6" r="5" fill="#404040"/>';
  const a = (k / 4) * Math.PI * 2;
  return `<circle cx="6" cy="6" r="5" fill="#fff" stroke="#404040"/><path d="M6 6 L6 1 A5 5 0 ${k > 2 ? 1 : 0} 1 ${6 + 5 * Math.sin(a)} ${6 - 5 * Math.cos(a)} Z" fill="#404040"/>`;
};
const star = (k) => {
  const pts = Array.from({ length: 10 }, (_, i) => { const r = i % 2 ? 2.3 : 5.2; const a = (Math.PI / 5) * i - Math.PI / 2; return `${6 + r * Math.cos(a)},${6 + r * Math.sin(a)}`; }).join(' ');
  const clip = k === 1 ? '<clipPath id="h"><rect x="0" y="0" width="6" height="12"/></clipPath>' : '';
  return `${clip}<polygon points="${pts}" fill="#fff" stroke="${GOLD}"/>${k ? `<polygon points="${pts}" fill="${GOLD}"${k === 1 ? ' clip-path="url(#h)"' : ''}/>` : ''}`;
};
const boxes = (k) => [0, 1, 2, 3].map((i) => `<rect x="${i % 2 ? 6.5 : 1}" y="${i < 2 ? 1 : 6.5}" width="4.5" height="4.5" fill="${i < k ? '#2F5597' : '#D9D9D9'}"/>`).join('');

/** Each set's icons, lowest first. */
const SETS = {
  '3Arrows': [arrow(180, RED), arrow(90, YELLOW), arrow(0, GREEN)],
  '3ArrowsGray': [arrow(180, GREY), arrow(90, GREY), arrow(0, GREY)],
  '4Arrows': [arrow(180, RED), arrow(135, YELLOW), arrow(45, YELLOW), arrow(0, GREEN)],
  '4ArrowsGray': [arrow(180, GREY), arrow(135, GREY), arrow(45, GREY), arrow(0, GREY)],
  '5Arrows': [arrow(180, RED), arrow(135, YELLOW), arrow(90, YELLOW), arrow(45, YELLOW), arrow(0, GREEN)],
  '5ArrowsGray': [arrow(180, GREY), arrow(135, GREY), arrow(90, GREY), arrow(45, GREY), arrow(0, GREY)],
  '3TrafficLights1': [dot(RED), dot(YELLOW), dot(GREEN)],
  '3TrafficLights2': [ring(RED, '<circle cx="6" cy="6" r="5" fill="none" stroke="#404040"/>'), ring(YELLOW, '<circle cx="6" cy="6" r="5" fill="none" stroke="#404040"/>'), ring(GREEN, '<circle cx="6" cy="6" r="5" fill="none" stroke="#404040"/>')],
  '4TrafficLights': [dot(BLACK), dot(RED), dot(YELLOW), dot(GREEN)],
  '4RedToBlack': [dot(BLACK), dot(GREY), dot('#F4A6A6'), dot(RED)],
  '3Signs': ['<path d="M6 1 L11 6 L6 11 L1 6 Z" fill="' + RED + '"/>', `<path d="M6 1.5 L11 10.5 H1 Z" fill="${YELLOW}"/>`, dot(GREEN)],
  '3Symbols': [ring(RED, cross), ring(YELLOW, bang), ring(GREEN, tick)],
  '3Symbols2': [`<path d="M2.5 2.5 L9.5 9.5 M9.5 2.5 L2.5 9.5" stroke="${RED}" stroke-width="2.2"/>`, `<path d="M6 1.5 V7.5" stroke="${YELLOW}" stroke-width="2.2"/><circle cx="6" cy="10" r="1.2" fill="${YELLOW}"/>`, `<path d="M2 6.5 L5 9.5 L10.5 3" stroke="${GREEN}" stroke-width="2.2" fill="none"/>`],
  '3Flags': [flag(RED), flag(YELLOW), flag(GREEN)],
  '4Rating': [bars(1, 4), bars(2, 4), bars(3, 4), bars(4, 4)],
  '5Rating': [bars(0, 4), bars(1, 4), bars(2, 4), bars(3, 4), bars(4, 4)],
  '5Quarters': [quarter(0), quarter(1), quarter(2), quarter(3), quarter(4)],
  '3Stars': [star(0), star(1), star(2)],
  '3Triangles': [`<path d="M6 10 L11 3 H1 Z" fill="${RED}"/>`, `<rect x="2" y="5" width="8" height="2.2" fill="${YELLOW}"/>`, `<path d="M6 2 L11 9 H1 Z" fill="${GREEN}"/>`],
  '5Boxes': [boxes(0), boxes(1), boxes(2), boxes(3), boxes(4)],
};
/** Names a screen reader can give an icon, lowest first, by kind of set. */
export const ICON_WORDS = (set, index, count) => `${set.replace(/^\d/, '').replace(/([a-z])([A-Z0-9])/g, '$1 $2').toLowerCase()} icon ${index + 1} of ${count}`;

/** An icon as an <svg> element (12 by 12), or null for a set not known. */
export function iconElement(set, index) {
  const icons = SETS[set];
  if (!icons || !icons[index]) return null;
  const span = document.createElement('span');
  span.className = 'oxe-cf-icon';
  span.setAttribute('aria-hidden', 'true');
  span.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 12 12" width="12" height="12">${icons[index]}</svg>`;
  return span;
}
