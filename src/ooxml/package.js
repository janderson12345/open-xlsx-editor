// The OPC package (the zip) behind an .xlsx (or any Office Open XML file). Every part is kept byte-for-byte;
// only the parts we change are rewritten on save, so anything we don't understand (drawings, pivot caches,
// comments, custom XML, macros...) survives.
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { parseXml, serializeXml, kids, NS, REL_TYPES, esc } from './xml.js';

export class OpcPackage {
  constructor(files) {
    /** @type {Record<string, Uint8Array>} */
    this.files = files;
    this.mainPath = this._findMainPart();
    this.mainDir = this.mainPath.slice(0, this.mainPath.lastIndexOf('/') + 1);
    this.rels = this.readRels(this.mainPath);
  }

  static fromBytes(bytes) {
    const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    let files;
    try {
      files = unzipSync(u8);
    } catch (e) {
      throw new Error('Not an Office Open XML file (could not unzip): ' + e.message);
    }
    if (!files['[Content_Types].xml']) throw new Error('Not an Office Open XML file: [Content_Types].xml missing');
    return new OpcPackage(files);
  }

  has(path) {
    return path in this.files;
  }
  text(path) {
    return this.files[path] ? strFromU8(this.files[path]) : null;
  }
  xml(path) {
    const t = this.text(path);
    return t == null ? null : parseXml(t);
  }
  setText(path, s) {
    this.files[path] = strToU8(s);
  }

  _findMainPart() {
    const rels = this.readRels('');
    const main = rels.find((r) => r.type === REL_TYPES.officeDocument || r.type.endsWith('/officeDocument'));
    if (!main) throw new Error('Not an Office Open XML file: no main part (officeDocument relationship)');
    return main.target.replace(/^\//, '');
  }

  /** Relationships of a part ('' = package root). Targets are resolved to package paths. */
  readRels(partPath) {
    const dir = partPath ? partPath.slice(0, partPath.lastIndexOf('/') + 1) : '';
    const name = partPath ? partPath.slice(partPath.lastIndexOf('/') + 1) : '';
    const relsPath = `${dir}_rels/${name}.rels`;
    const doc = this.xml(relsPath);
    if (!doc) return [];
    return kids(doc.documentElement, 'Relationship', NS.rel).map((r) => {
      const target = r.getAttribute('Target');
      const external = r.getAttribute('TargetMode') === 'External';
      return {
        id: r.getAttribute('Id'),
        type: r.getAttribute('Type'),
        target: external ? target : resolvePath(dir, target),
        rawTarget: target,
        external,
      };
    });
  }

  relById(id) {
    return this.rels.find((r) => r.id === id) || null;
  }
  partByType(type) {
    const r = this.rels.find((x) => x.type === type);
    return r ? r.target : null;
  }

  /** Add a relationship to a part (the main part, xl/workbook.xml, by default); returns the new rId (`want`, when it's
   *  free: the id another editor of a shared workbook gave the same part). */
  addRel(type, target, external = false, partPath = this.mainPath, want = null) {
    const dir = partPath.slice(0, partPath.lastIndexOf('/') + 1);
    const relsPath = `${dir}_rels/${partPath.slice(dir.length)}.rels`;
    const doc = this.xml(relsPath) || parseXml(`<Relationships xmlns="${NS.rel}"/>`);
    const used = new Set(kids(doc.documentElement, 'Relationship', NS.rel).map((r) => r.getAttribute('Id')));
    // (Co-editing gives each editor a prefix of its own, idPrefix: two editors adding parts at once can't both take
    // the next rIdN, and a part another editor shares never meets an id this one made for something else.)
    const prefix = this.idPrefix || 'rId';
    let n = this.idPrefix ? 1 : used.size + 1;
    while (used.has(prefix + n)) n++;
    const id = want && !used.has(want) ? want : prefix + n;
    const el = doc.createElementNS(NS.rel, 'Relationship');
    el.setAttribute('Id', id);
    el.setAttribute('Type', type);
    el.setAttribute('Target', target);
    if (external) el.setAttribute('TargetMode', 'External');
    doc.documentElement.appendChild(el);
    this.setText(relsPath, serializeXml(doc));
    if (partPath === this.mainPath) this.rels = this.readRels(this.mainPath);
    return id;
  }

  /** Register a content-type override for a new part. */
  addOverride(partName, contentType) {
    const doc = this.xml('[Content_Types].xml');
    const exists = kids(doc.documentElement, 'Override', NS.ct).some(
      (o) => o.getAttribute('PartName') === '/' + partName,
    );
    if (exists) return;
    const el = doc.createElementNS(NS.ct, 'Override');
    el.setAttribute('PartName', '/' + partName);
    el.setAttribute('ContentType', contentType);
    doc.documentElement.appendChild(el);
    this.setText('[Content_Types].xml', serializeXml(doc));
  }

  /** Take a part out of the package: the part, its own relationships, the main part's relationships to it and its
   *  content-type override. */
  removePart(path) {
    delete this.files[path];
    const name = path.slice(path.lastIndexOf('/') + 1);
    delete this.files[`${path.slice(0, path.length - name.length)}_rels/${name}.rels`];
    const dir = this.mainDir;
    const relsPath = `${dir}_rels/${this.mainPath.slice(dir.length)}.rels`;
    const rels = this.xml(relsPath);
    if (rels) {
      for (const r of kids(rels.documentElement, 'Relationship', NS.rel)) {
        if (r.getAttribute('TargetMode') !== 'External' && resolvePath(dir, r.getAttribute('Target')) === path) r.parentNode.removeChild(r);
      }
      this.setText(relsPath, serializeXml(rels));
      this.rels = this.readRels(this.mainPath);
    }
    const ct = this.xml('[Content_Types].xml');
    for (const o of kids(ct.documentElement, 'Override', NS.ct)) if (o.getAttribute('PartName') === '/' + path) o.parentNode.removeChild(o);
    this.setText('[Content_Types].xml', serializeXml(ct));
  }

  /** Take a part out with the parts only it leads to (a deleted sheet's drawings, comments, tables, printer
   *  settings): whatever another part still reaches through its relationships stays. */
  removeTree(path) {
    const walk = (from, skip) => {
      const seen = new Set();
      const queue = [...from];
      while (queue.length) {
        const p = queue.pop();
        if (seen.has(p) || p === skip) continue;
        seen.add(p);
        for (const r of this.readRels(p)) if (!r.external && this.has(r.target)) queue.push(r.target);
      }
      return seen;
    };
    const kept = walk(this.readRels('').filter((r) => !r.external).map((r) => r.target), path);
    for (const p of walk([path], null)) if (!kept.has(p)) this.removePart(p);
  }

  /** Resolve an image relationship to a data blob for display. */
  media(rId) {
    const rel = this.relById(rId);
    if (!rel || rel.external) return null;
    const bytes = this.files[rel.target];
    if (!bytes) return null;
    const ext = rel.target.split('.').pop().toLowerCase();
    const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' }[ext] || 'application/octet-stream';
    return { bytes, mime };
  }

  toBytes() {
    // Store [Content_Types].xml first, as Office itself does.
    const ordered = { '[Content_Types].xml': this.files['[Content_Types].xml'] };
    for (const [k, v] of Object.entries(this.files)) if (k !== '[Content_Types].xml') ordered[k] = v;
    return zipSync(ordered, { level: 6 });
  }
}

export function resolvePath(baseDir, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (baseDir + target).split('/');
  const out = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.' && p !== '') out.push(p);
  }
  return out.join('/');
}

export { esc };
