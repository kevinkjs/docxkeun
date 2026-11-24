/**
 * Open Packaging Conventions: parts, content types and relationships on top of the zip layer.
 * Parts that nobody touched are written back from their original compressed bytes.
 */
import { readZip, writeZip, type ZipRawData, type ZipWriteEntry } from '../zip/zip.js';
import { decodeXmlBytes, utf8Encode } from '../util/text.js';
import { cloneElement, el, elements, parseXml, serializeXml, type XmlDocument } from '../xml/index.js';
import { CONTENT_TYPE, IMAGE_CONTENT_TYPES, isRelType, type RelKind, REL } from './constants.js';
import { extension, normalizePartName, partKey, relsPartFor, relativeTarget, resolveTarget } from './paths.js';

export interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

interface Part {
  name: string;
  raw?: ZipRawData;
  bytes?: Uint8Array;
  xml?: XmlDocument;
  dirty: boolean;
  /** Changes whenever the part may have changed; lets checkpoints reuse unchanged snapshots. */
  rev: number;
}

/** An opaque snapshot of every part, content type and relationship list (see `OpcPackage.checkpoint`). */
export interface PackageCheckpoint {
  readonly parts: Array<[string, Part]>;
  readonly order: string[];
  readonly defaults: Array<[string, string]>;
  readonly overrides: Array<[string, string]>;
  readonly contentTypesDirty: boolean;
  readonly rels: Array<[string, Relationship[]]>;
}

export class OpcError extends Error {
  override name = 'OpcError';
}

const CT_PART = '[Content_Types].xml';

export class OpcPackage {
  private readonly parts = new Map<string, Part>();
  private readonly order: string[] = [];
  private readonly defaults = new Map<string, string>();
  private readonly overrides = new Map<string, string>();
  private contentTypesDirty = false;
  private revSeq = 0;
  /** Last frozen copy per part, with the revision it was taken at. Never handed to live state. */
  private readonly snapshots = new Map<string, { rev: number; part: Part }>();
  private readonly relsCache = new Map<string, Relationship[]>();
  private readonly readers = new Map<string, () => Promise<Uint8Array>>();
  private contentTypesRaw: ZipRawData | undefined;

  /** Open a package from zip bytes. */
  static async open(bytes: Uint8Array): Promise<OpcPackage> {
    const pkg = new OpcPackage();
    const entries = readZip(bytes);
    let ctEntry: (typeof entries)[number] | undefined;
    for (const entry of entries) {
      if (entry.name.endsWith('/')) continue; // directory entries
      if (entry.name === CT_PART) {
        ctEntry = entry;
        continue;
      }
      const name = normalizePartName(entry.name);
      // A zip can repeat a name (some tools append parts instead of replacing them). Keep the
      // last copy, as unzip tools do, in the first copy's place, so the part is saved once.
      if (!pkg.parts.has(partKey(name))) pkg.order.push(partKey(name));
      pkg.parts.set(partKey(name), { name, raw: entry.raw, dirty: false, rev: 0 });
      // Keep a reader so bytes are inflated lazily.
      pkg.readers.set(partKey(name), entry.read);
    }
    if (!ctEntry) throw new OpcError('Missing [Content_Types].xml: not an OOXML package');
    pkg.contentTypesRaw = ctEntry.raw;
    const ct = parseXml(decodeXmlBytes(await ctEntry.read()));
    for (const node of elements(ct.root)) {
      const local = node.name.replace(/^.*:/, '');
      if (local === 'Default' && node.attrs.Extension && node.attrs.ContentType) {
        pkg.defaults.set(node.attrs.Extension.toLowerCase(), node.attrs.ContentType);
      } else if (local === 'Override' && node.attrs.PartName && node.attrs.ContentType) {
        pkg.overrides.set(partKey(node.attrs.PartName), node.attrs.ContentType);
      }
    }
    return pkg;
  }

  /** An empty package with the standard defaults. */
  static create(): OpcPackage {
    const pkg = new OpcPackage();
    pkg.defaults.set('rels', CONTENT_TYPE.relationships);
    pkg.defaults.set('xml', CONTENT_TYPE.xml);
    pkg.contentTypesDirty = true;
    return pkg;
  }

  /**
   * Snapshot the whole package so `restore` can bring it back. Parsed XML is deep-cloned; bytes are shared
   * because they are replaced, never mutated in place.
   */
  checkpoint(): PackageCheckpoint {
    return {
      parts: [...this.parts].map(([key, part]) => [key, this.frozenCopy(key, part)]),
      order: [...this.order],
      defaults: [...this.defaults],
      overrides: [...this.overrides],
      contentTypesDirty: this.contentTypesDirty,
      rels: [...this.relsCache].map(([key, list]) => [key, list.map((r) => ({ ...r }))]),
    };
  }

  private frozenCopy(key: string, part: Part): Part {
    const cached = this.snapshots.get(key);
    if (cached && cached.rev === part.rev) return cached.part;
    const copy = copyPart(part);
    this.snapshots.set(key, { rev: part.rev, part: copy });
    return copy;
  }

  /** Put the package back as it was at `checkpoint()`. The same checkpoint can be restored again. */
  restore(cp: PackageCheckpoint): void {
    this.parts.clear();
    for (const [key, part] of cp.parts) this.parts.set(key, { ...copyPart(part), rev: ++this.revSeq });
    this.snapshots.clear();
    this.order.splice(0, this.order.length, ...cp.order);
    this.defaults.clear();
    for (const [k, v] of cp.defaults) this.defaults.set(k, v);
    this.overrides.clear();
    for (const [k, v] of cp.overrides) this.overrides.set(k, v);
    this.contentTypesDirty = cp.contentTypesDirty;
    this.relsCache.clear();
    for (const [key, list] of cp.rels) this.relsCache.set(key, list.map((r) => ({ ...r })));
  }

  partNames(): string[] {
    return this.order.map((k) => this.parts.get(k)!.name);
  }

  has(name: string): boolean {
    return this.parts.has(partKey(name));
  }

  contentType(name: string): string | undefined {
    const key = partKey(name);
    return this.overrides.get(key) ?? this.defaults.get(extension(name));
  }

  /** Every part with a given content type. */
  partsOfType(contentType: string): string[] {
    return this.partNames().filter((n) => this.contentType(n) === contentType);
  }

  async getBytes(name: string): Promise<Uint8Array> {
    const part = this.requirePart(name);
    if (part.xml && part.dirty) return utf8Encode(serializeXml(part.xml));
    if (part.bytes) return part.bytes;
    const reader = this.readers.get(partKey(name));
    if (!reader) throw new OpcError(`Part ${name} has no data`);
    part.bytes = await reader();
    return part.bytes;
  }

  /** Parsed XML of a part (cached). Mutate it, then call `markDirty`. */
  async getXml(name: string): Promise<XmlDocument> {
    const part = this.requirePart(name);
    if (!part.xml) part.xml = parseXml(decodeXmlBytes(await this.getBytes(name)));
    return part.xml;
  }

  /** Parsed XML of a part that was already loaded with `getXml` (no I/O). */
  cachedXml(name: string): XmlDocument | undefined {
    return this.parts.get(partKey(name))?.xml;
  }

  async tryGetXml(name: string | undefined): Promise<XmlDocument | undefined> {
    if (!name || !this.has(name)) return undefined;
    return this.getXml(name);
  }

  markDirty(name: string): void {
    const part = this.requirePart(name);
    part.dirty = true;
    part.rev = ++this.revSeq;
  }

  isDirty(name: string): boolean {
    return this.parts.get(partKey(name))?.dirty ?? false;
  }

  setXml(name: string, doc: XmlDocument, contentType?: string): void {
    const part = this.ensurePart(name, contentType);
    part.xml = doc;
    part.bytes = undefined;
    part.dirty = true;
    part.rev = ++this.revSeq;
  }

  setBytes(name: string, bytes: Uint8Array, contentType?: string): void {
    const part = this.ensurePart(name, contentType);
    part.bytes = bytes;
    part.xml = undefined;
    part.dirty = true;
    part.rev = ++this.revSeq;
  }

  removePart(name: string): void {
    const key = partKey(name);
    if (!this.parts.delete(key)) return;
    this.snapshots.delete(key);
    this.order.splice(this.order.indexOf(key), 1);
    if (this.overrides.delete(key)) this.contentTypesDirty = true;
    this.relsCache.delete(key);
    const relsName = relsPartFor(normalizePartName(name));
    if (this.has(relsName)) this.removePart(relsName);
  }

  private ensurePart(name: string, contentType?: string): Part {
    const clean = normalizePartName(name);
    const key = partKey(clean);
    let part = this.parts.get(key);
    if (!part) {
      part = { name: clean, dirty: true, rev: 0 };
      this.parts.set(key, part);
      this.order.push(key);
    }
    if (contentType) this.setContentType(clean, contentType);
    else if (!this.contentType(clean)) {
      const ext = extension(clean);
      const image = IMAGE_CONTENT_TYPES[ext];
      if (image) this.setDefaultContentType(ext, image);
    }
    return part;
  }

  setContentType(name: string, contentType: string): void {
    const key = partKey(name);
    if (this.overrides.get(key) === contentType) return;
    if (this.defaults.get(extension(name)) === contentType && !this.overrides.has(key)) return;
    this.overrides.set(key, contentType);
    this.contentTypesDirty = true;
  }

  setDefaultContentType(ext: string, contentType: string): void {
    if (this.defaults.get(ext.toLowerCase()) === contentType) return;
    this.defaults.set(ext.toLowerCase(), contentType);
    this.contentTypesDirty = true;
  }

  private requirePart(name: string): Part {
    const part = this.parts.get(partKey(name));
    if (!part) throw new OpcError(`No part named ${name}`);
    return part;
  }

  // ---- relationships -------------------------------------------------------

  /** Relationships of a source part ('' = package level). */
  async relationships(source: string): Promise<Relationship[]> {
    const key = partKey(source);
    const cached = this.relsCache.get(key);
    if (cached) return cached;
    const relsName = relsPartFor(normalizePartName(source));
    const rels: Relationship[] = [];
    if (this.has(relsName)) {
      const doc = await this.getXml(relsName);
      for (const node of elements(doc.root)) {
        if (node.name.replace(/^.*:/, '') !== 'Relationship') continue;
        const { Id, Type, Target, TargetMode } = node.attrs;
        if (!Id || !Type || Target === undefined) continue;
        rels.push({ id: Id, type: Type, target: Target, external: TargetMode === 'External' });
      }
    }
    this.relsCache.set(key, rels);
    return rels;
  }

  async relationshipById(source: string, id: string): Promise<Relationship | undefined> {
    return (await this.relationships(source)).find((r) => r.id === id);
  }

  async relationshipsOfKind(source: string, kind: RelKind): Promise<Relationship[]> {
    return (await this.relationships(source)).filter((r) => isRelType(r.type, kind));
  }

  /** Part name targeted by the first relationship of a kind, if it exists. */
  async targetPart(source: string, kind: RelKind): Promise<string | undefined> {
    const rel = (await this.relationshipsOfKind(source, kind)).find((r) => !r.external);
    return rel ? resolveTarget(source, rel.target) : undefined;
  }

  resolve(source: string, rel: Relationship): string {
    return rel.external ? rel.target : resolveTarget(source, rel.target);
  }

  /** Add a relationship and return its id. Internal targets are given as part names. */
  async addRelationship(source: string, type: string, target: string, external = false): Promise<string> {
    const rels = await this.relationships(source);
    const used = new Set(rels.map((r) => r.id));
    let n = rels.length + 1;
    while (used.has(`rId${n}`)) n++;
    const id = `rId${n}`;
    const relTarget = external ? target : relativeTarget(normalizePartName(source), target);
    rels.push({ id, type, target: relTarget, external });
    this.writeRels(source, rels);
    return id;
  }

  /** Reuse an existing relationship to the same target, or add one. */
  async ensureRelationship(source: string, type: string, target: string, external = false): Promise<string> {
    const rels = await this.relationships(source);
    const wanted = external ? target : relativeTarget(normalizePartName(source), target);
    const found = rels.find((r) => r.type === type && r.external === external && (r.target === wanted || (!external && resolveTarget(source, r.target) === normalizePartName(target))));
    return found ? found.id : this.addRelationship(source, type, target, external);
  }

  async removeRelationship(source: string, id: string): Promise<void> {
    const rels = await this.relationships(source);
    const i = rels.findIndex((r) => r.id === id);
    if (i < 0) return;
    rels.splice(i, 1);
    this.writeRels(source, rels);
  }

  private writeRels(source: string, rels: Relationship[]): void {
    const root = el(
      'Relationships',
      { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' },
      rels.map((r) => el('Relationship', { Id: r.id, Type: r.type, Target: r.target, TargetMode: r.external ? 'External' : undefined })),
    );
    this.setXml(relsPartFor(normalizePartName(source)), { declaration: null, prolog: [], root, epilog: [] });
    this.relsCache.set(partKey(source), rels);
  }

  // ---- saving ---------------------------------------------------------------

  private contentTypesXml(): XmlDocument {
    const children = [
      ...[...this.defaults].map(([ext, type]) => el('Default', { Extension: ext, ContentType: type })),
      ...[...this.overrides]
        .filter(([key]) => this.parts.has(key))
        .map(([key, type]) => el('Override', { PartName: '/' + this.parts.get(key)!.name, ContentType: type })),
    ];
    return {
      declaration: null,
      prolog: [],
      root: el('Types', { xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types' }, children),
      epilog: [],
    };
  }

  /** Serialise to zip bytes. Unchanged parts keep their original compressed data. */
  async save(options: { modified?: Date } = {}): Promise<Uint8Array> {
    const entries: ZipWriteEntry[] = [];
    if (!this.contentTypesDirty && this.contentTypesRaw) entries.push({ name: CT_PART, raw: this.contentTypesRaw });
    else entries.push({ name: CT_PART, data: utf8Encode(serializeXml(this.contentTypesXml())) });
    for (const key of this.order) {
      const part = this.parts.get(key)!;
      if (!part.dirty && part.raw) entries.push({ name: part.name, raw: part.raw });
      else entries.push({ name: part.name, data: await this.getBytes(part.name) });
    }
    return writeZip(entries, options);
  }

  /** True when content types changed since opening (exposed for tests). */
  get contentTypesChanged(): boolean {
    return this.contentTypesDirty;
  }
}

export { REL };

function copyPart(part: Part): Part {
  const copy: Part = { ...part };
  if (part.xml) copy.xml = { ...part.xml, root: cloneElement(part.xml.root) };
  return copy;
}
