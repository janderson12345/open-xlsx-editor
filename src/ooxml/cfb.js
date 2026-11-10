// Compound files (MS-CFB, the "OLE" container): what an encrypted workbook is saved in, its streams EncryptionInfo and
// EncryptedPackage beside a \x06DataSpaces storage. Read: every stream by its path ("EncryptionInfo",
// "\x06DataSpaces/Version"). Written: version 3 (512-byte sectors), streams under 4096 bytes in the mini stream, as
// the format requires.

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const FREE = 0xffffffff;
const END = 0xfffffffe;
const FAT_SECT = 0xfffffffd;
const DIFAT_SECT = 0xfffffffc;
const NONE = 0xffffffff;
const CUTOFF = 4096;

/** Whether `bytes` start as a compound file does. */
export function isCompoundFile(bytes) {
  return bytes.length >= 512 && SIGNATURE.every((b, i) => bytes[i] === b);
}

/** Every stream of a compound file: Map(path => bytes), storages joined with "/". Throws on a damaged file. */
export function readCompoundFile(bytes) {
  if (!isCompoundFile(bytes)) throw new Error('not a compound file');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = (o) => view.getUint16(o, true);
  const u32 = (o) => view.getUint32(o, true);
  const shift = u16(0x1e);
  const size = 1 << shift;
  const miniSize = 1 << u16(0x20);
  const sector = (n) => {
    const o = (n + 1) * size;
    if (o + size > bytes.length + size) throw new Error('damaged compound file: a sector past its end');
    return bytes.subarray(o, Math.min(o + size, bytes.length));
  };
  // (The FAT's sectors: the header's 109, then the DIFAT sectors', 127 each and the next one's number last.)
  const fatSectors = [];
  for (let i = 0; i < 109; i++) { const s = u32(0x4c + i * 4); if (s !== FREE) fatSectors.push(s); }
  for (let d = u32(0x44), n = u32(0x48); d !== END && d !== FREE && n > 0; n--) {
    const s = sector(d);
    const dv = new DataView(s.buffer, s.byteOffset, s.byteLength);
    for (let i = 0; i < size / 4 - 1; i++) { const x = dv.getUint32(i * 4, true); if (x !== FREE) fatSectors.push(x); }
    d = dv.getUint32(size - 4, true);
  }
  const fat = [];
  for (const s of fatSectors) {
    const b = sector(s);
    const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
    for (let i = 0; i < b.length / 4; i++) fat.push(dv.getUint32(i * 4, true));
  }
  const chain = (start, table) => {
    const out = [];
    for (let s = start, guard = 0; s !== END && s !== FREE; s = table[s]) {
      if (s >= table.length || ++guard > table.length) throw new Error('damaged compound file: a broken sector chain');
      out.push(s);
    }
    return out;
  };
  const join = (parts) => { const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0)); let o = 0; for (const p of parts) { out.set(p, o); o += p.length; } return out; };
  const stream = (start) => join(chain(start, fat).map(sector));
  const dir = stream(u32(0x30));
  const dv = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  const entries = [];
  for (let o = 0; o + 128 <= dir.length; o += 128) {
    const len = dv.getUint16(o + 0x40, true);
    let name = '';
    for (let i = 0; i + 2 < len; i += 2) name += String.fromCharCode(dv.getUint16(o + i, true));
    entries.push({ name, type: dir[o + 0x42], left: dv.getUint32(o + 0x44, true), right: dv.getUint32(o + 0x48, true), child: dv.getUint32(o + 0x4c, true), start: dv.getUint32(o + 0x74, true), size: dv.getUint32(o + 0x78, true) });
  }
  const root = entries[0];
  if (!root || root.type !== 5) throw new Error('damaged compound file: no root entry');
  const miniFat = [];
  if (u32(0x40) > 0) {
    const b = stream(u32(0x3c));
    const m = new DataView(b.buffer, b.byteOffset, b.byteLength);
    for (let i = 0; i < b.length / 4; i++) miniFat.push(m.getUint32(i * 4, true));
  }
  const mini = root.start === END ? new Uint8Array(0) : stream(root.start);
  const read = (e) => {
    if (e.size < u32(0x38)) return join(chain(e.start, miniFat).map((s) => mini.subarray(s * miniSize, (s + 1) * miniSize))).subarray(0, e.size);
    return stream(e.start).subarray(0, e.size);
  };
  const out = new Map();
  const seen = new Set();
  const walk = (i, prefix) => {
    if (i === NONE || i >= entries.length || seen.has(i)) return;
    seen.add(i);
    const e = entries[i];
    walk(e.left, prefix);
    if (e.type === 2) out.set(prefix + e.name, read(e));
    else if (e.type === 1) walk(e.child, `${prefix}${e.name}/`);
    walk(e.right, prefix);
  };
  walk(root.child, '');
  return out;
}

/** The order siblings are kept in: shorter names first, then by their upper-case letters. */
const compareNames = (a, b) => a.length - b.length || (a.toUpperCase() < b.toUpperCase() ? -1 : a.toUpperCase() > b.toUpperCase() ? 1 : 0);

/** A compound file holding `streams` (Map or entries of path => bytes, storages joined with "/"). */
export function writeCompoundFile(streams) {
  // (The tree of storages and streams; every storage's children as a balanced binary tree, all black.)
  const root = { name: 'Root Entry', type: 5, kids: new Map() };
  for (const [path, data] of streams instanceof Map ? streams : new Map(streams)) {
    const parts = path.split('/');
    let at = root;
    for (const p of parts.slice(0, -1)) {
      if (!at.kids.has(p)) at.kids.set(p, { name: p, type: 1, kids: new Map() });
      at = at.kids.get(p);
    }
    at.kids.set(parts[parts.length - 1], { name: parts[parts.length - 1], type: 2, data });
  }
  const list = [];
  const add = (node) => { node.id = list.length; list.push(node); return node; };
  const place = (node) => {
    const kids = [...(node.kids?.values() ?? [])].sort((a, b) => compareNames(a.name, b.name));
    for (const k of kids) add(k);
    const tree = (lo, hi) => {
      if (lo > hi) return NONE;
      const mid = (lo + hi) >> 1;
      kids[mid].left = tree(lo, mid - 1);
      kids[mid].right = tree(mid + 1, hi);
      return kids[mid].id;
    };
    node.child = tree(0, kids.length - 1);
    for (const k of kids) if (k.type === 1) place(k);
  };
  add(root);
  place(root);
  const streamsList = list.filter((n) => n.type === 2);
  const small = streamsList.filter((n) => n.data.length < CUTOFF);
  const big = streamsList.filter((n) => n.data.length >= CUTOFF);
  // (The mini stream: small streams in 64-byte mini sectors, chained in the mini FAT.)
  const miniFat = [];
  let miniLen = 0;
  for (const n of small) {
    const count = Math.ceil(n.data.length / 64);
    n.start = count ? miniLen / 64 : END;
    for (let k = 0; k < count; k++) miniFat.push(k === count - 1 ? END : miniLen / 64 + k + 1);
    n.miniAt = miniLen;
    miniLen += count * 64;
  }
  const sectors = (len) => Math.ceil(len / 512);
  const dirSectors = sectors(list.length * 128);
  const miniFatSectors = sectors(miniFat.length * 4);
  const miniSectors = sectors(miniLen);
  const bigSectors = big.reduce((n, s) => n + sectors(s.data.length), 0);
  const others = dirSectors + miniFatSectors + miniSectors + bigSectors;
  let fatCount = 1;
  let difatCount = 0;
  for (;;) {
    difatCount = fatCount > 109 ? Math.ceil((fatCount - 109) / 127) : 0;
    if (fatCount * 128 >= others + fatCount + difatCount) break;
    fatCount++;
  }
  const total = fatCount + difatCount + others;
  const fat = new Array(fatCount * 128).fill(FREE);
  let next = 0;
  const run = (count, mark = null) => {
    const first = next;
    for (let k = 0; k < count; k++) fat[next + k] = mark ?? (k === count - 1 ? END : next + k + 1);
    next += count;
    return count ? first : END;
  };
  const fatAt = run(fatCount, FAT_SECT);
  const difatAt = run(difatCount, DIFAT_SECT);
  const dirAt = run(dirSectors);
  const miniFatAt = run(miniFatSectors);
  root.start = run(miniSectors);
  root.size = miniLen;
  for (const n of big) n.start = run(sectors(n.data.length));
  const out = new Uint8Array((total + 1) * 512);
  const dv = new DataView(out.buffer);
  const at = (s) => (s + 1) * 512;
  // Header.
  out.set(SIGNATURE, 0);
  dv.setUint16(0x18, 0x3e, true);
  dv.setUint16(0x1a, 3, true);
  dv.setUint16(0x1c, 0xfffe, true);
  dv.setUint16(0x1e, 9, true);
  dv.setUint16(0x20, 6, true);
  dv.setUint32(0x2c, fatCount, true);
  dv.setUint32(0x30, dirAt, true);
  dv.setUint32(0x38, CUTOFF, true);
  dv.setUint32(0x3c, miniFatSectors ? miniFatAt : END, true);
  dv.setUint32(0x40, miniFatSectors, true);
  dv.setUint32(0x44, difatCount ? difatAt : END, true);
  dv.setUint32(0x48, difatCount, true);
  for (let i = 0; i < 109; i++) dv.setUint32(0x4c + i * 4, i < fatCount ? fatAt + i : FREE, true);
  // The DIFAT sectors: the rest of the FAT's sector numbers, 127 to a sector, the next one's last.
  for (let d = 0; d < difatCount; d++) {
    const o = at(difatAt + d);
    for (let i = 0; i < 127; i++) { const k = 109 + d * 127 + i; dv.setUint32(o + i * 4, k < fatCount ? fatAt + k : FREE, true); }
    dv.setUint32(o + 508, d < difatCount - 1 ? difatAt + d + 1 : END, true);
  }
  fat.forEach((x, i) => dv.setUint32(at(fatAt) + i * 4, x, true));
  // The directory.
  const dirBase = at(dirAt);
  for (let i = 0; i < dirSectors * 4; i++) {
    const o = dirBase + i * 128;
    const n = list[i];
    if (!n) { dv.setUint32(o + 0x44, NONE, true); dv.setUint32(o + 0x48, NONE, true); dv.setUint32(o + 0x4c, NONE, true); continue; }
    for (let k = 0; k < n.name.length; k++) dv.setUint16(o + k * 2, n.name.charCodeAt(k), true);
    dv.setUint16(o + 0x40, (n.name.length + 1) * 2, true);
    out[o + 0x42] = n.type;
    out[o + 0x43] = 1; // (black)
    dv.setUint32(o + 0x44, n.left ?? NONE, true);
    dv.setUint32(o + 0x48, n.right ?? NONE, true);
    dv.setUint32(o + 0x4c, n.type === 2 ? NONE : n.child, true);
    dv.setUint32(o + 0x74, n.type === 1 ? 0 : n.start, true);
    dv.setUint32(o + 0x78, n.type === 1 ? 0 : n.type === 5 ? n.size : n.data.length, true);
  }
  miniFat.forEach((x, i) => dv.setUint32(at(miniFatAt) + i * 4, x, true));
  for (let i = miniFat.length; i < miniFatSectors * 128; i++) dv.setUint32(at(miniFatAt) + i * 4, FREE, true);
  for (const n of small) out.set(n.data, at(root.start) + n.miniAt);
  for (const n of big) out.set(n.data, at(n.start));
  return out;
}
