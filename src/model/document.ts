/** A loaded .docx: its package, parsed parts and readers for every story. */
import { CONTENT_TYPE, OpcPackage, REL, isRelType, relTypeName } from '../opc/index.js';
import { attr, elements, firstChild, parseXml, textContent, type XmlDocument } from '../xml/index.js';
import { StoryReader, type Block, type Story } from './content.js';
import { NumberingDefs } from './numbering.js';
import { StyleSheet } from './styles.js';
import { Theme } from './theme.js';
import { parseTwipsMeasure } from './units.js';

export interface CoreProperties {
  title?: string;
  subject?: string;
  creator?: string;
  keywords?: string;
  description?: string;
  lastModifiedBy?: string;
  created?: string;
  modified?: string;
}

export interface WordParts {
  styles?: string;
  numbering?: string;
  settings?: string;
  theme?: string;
  fontTable?: string;
  footnotes?: string;
  endnotes?: string;
  comments?: string;
  commentsExtended?: string;
  commentsIds?: string;
  commentsExtensible?: string;
  people?: string;
}

const PART_KINDS: Array<keyof WordParts> = [
  'styles',
  'numbering',
  'settings',
  'theme',
  'fontTable',
  'footnotes',
  'endnotes',
  'comments',
  'commentsExtended',
  'commentsIds',
  'commentsExtensible',
  'people',
];

export type RelMap = Map<string, { target: string; external: boolean }>;

export interface OpenOptions {
  /** Password for encrypted (password-to-open) documents. */
  password?: string;
  /** Show tracked deletions instead of hiding them. */
  showDeletions?: boolean;
}

/** Hook used by the security module to decrypt password-protected files (set when it is imported). */
export const decryptionHooks: { decrypt?: (bytes: Uint8Array, password: string | undefined) => Promise<Uint8Array> } = {};

export class DocxError extends Error {
  override name = 'DocxError';
}

export function isCompoundFile(bytes: Uint8Array): boolean {
  return bytes.length >= 8 && bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0 && bytes[4] === 0xa1 && bytes[5] === 0xb1 && bytes[6] === 0x1a && bytes[7] === 0xe1;
}

export async function toBytes(input: Uint8Array | ArrayBuffer | Blob): Promise<Uint8Array> {
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  return new Uint8Array(await input.arrayBuffer());
}

export class DocxDocument {
  styles!: StyleSheet;
  numbering!: NumberingDefs;
  theme!: Theme;
  readonly parts: WordParts = {};
  documentXml!: XmlDocument;
  private readonly relMaps = new Map<string, RelMap>();
  private bodyCache: Story | undefined;
  private readonly storyCache = new Map<string, Story>();
  private readonly listeners = new Set<(part: string) => void>();
  /** Set when the source file was encrypted, so `save()` can re-encrypt by default. */
  sourcePassword: string | undefined;

  private constructor(readonly pkg: OpcPackage, readonly mainPart: string, private readonly options: OpenOptions) {}

  static async open(input: Uint8Array | ArrayBuffer | Blob, options: OpenOptions = {}): Promise<DocxDocument> {
    let bytes = await toBytes(input);
    let password: string | undefined;
    if (isCompoundFile(bytes)) {
      if (!decryptionHooks.decrypt) {
        throw new DocxError('This file is encrypted or is a legacy .doc. Import "docxkit/security" to open password-protected files.');
      }
      bytes = await decryptionHooks.decrypt(bytes, options.password);
      password = options.password;
    }
    const pkg = await OpcPackage.open(bytes);
    const doc = await DocxDocument.fromPackage(pkg, options);
    doc.sourcePassword = password;
    return doc;
  }

  static async fromPackage(pkg: OpcPackage, options: OpenOptions = {}): Promise<DocxDocument> {
    const main = await pkg.targetPart('', 'officeDocument');
    if (!main || !pkg.has(main)) throw new DocxError('Package has no main document part');
    const doc = new DocxDocument(pkg, main, options);
    await doc.load();
    return doc;
  }

  /** A new, empty document with Word's default styles. */
  static async create(): Promise<DocxDocument> {
    const pkg = OpcPackage.create();
    pkg.setXml('word/document.xml', parseXml(BLANK_DOCUMENT), CONTENT_TYPE.document);
    pkg.setXml('word/styles.xml', parseXml(BLANK_STYLES), CONTENT_TYPE.styles);
    pkg.setXml('word/settings.xml', parseXml(BLANK_SETTINGS), CONTENT_TYPE.settings);
    await pkg.addRelationship('', REL.officeDocument, 'word/document.xml');
    await pkg.addRelationship('word/document.xml', REL.styles, 'word/styles.xml');
    await pkg.addRelationship('word/document.xml', REL.settings, 'word/settings.xml');
    return DocxDocument.fromPackage(pkg);
  }

  private async load(): Promise<void> {
    const { pkg, mainPart } = this;
    this.documentXml = await pkg.getXml(mainPart);
    const rels = await pkg.relationships(mainPart);
    for (const kind of PART_KINDS) {
      const rel = rels.find((r) => !r.external && isRelType(r.type, kind));
      if (rel) {
        const name = pkg.resolve(mainPart, rel);
        if (pkg.has(name)) this.parts[kind] = name;
      }
    }
    await this.loadRels(mainPart);
    // Headers, footers and notes have their own relationships (images, hyperlinks).
    for (const rel of rels) {
      if (rel.external) continue;
      const kind = relTypeName(rel.type);
      if (kind === 'header' || kind === 'footer') {
        const name = pkg.resolve(mainPart, rel);
        if (pkg.has(name)) {
          await pkg.getXml(name);
          await this.loadRels(name);
        }
      }
    }
    for (const kind of ['footnotes', 'endnotes', 'comments'] as const) {
      const name = this.parts[kind];
      if (name) {
        await pkg.getXml(name);
        await this.loadRels(name);
      }
    }
    for (const kind of PART_KINDS) {
      const name = this.parts[kind];
      if (name) await pkg.getXml(name);
    }
    this.reparseDefinitions();
  }

  private async loadRels(part: string): Promise<RelMap> {
    const map: RelMap = new Map();
    for (const r of await this.pkg.relationships(part)) map.set(r.id, { target: this.pkg.resolve(part, r), external: r.external });
    this.relMaps.set(part, map);
    return map;
  }

  /** Re-read styles, numbering and theme (after they were edited). */
  reparseDefinitions(): void {
    this.styles = new StyleSheet(this.xmlSync(this.parts.styles));
    this.numbering = new NumberingDefs(this.xmlSync(this.parts.numbering));
    this.theme = Theme.parse(this.xmlSync(this.parts.theme));
    this.invalidate();
  }

  /** XML of a part that `load()` already parsed. */
  xmlSync(part: string | undefined): XmlDocument | undefined {
    if (!part) return undefined;
    return this.pkg.cachedXml(part);
  }

  rels(part: string): RelMap {
    return this.relMaps.get(part) ?? new Map();
  }

  /** Refresh a part's relationship map after relationships were added. */
  async refreshRels(part: string): Promise<void> {
    await this.loadRels(part);
  }

  private reader(part: string): StoryReader {
    return new StoryReader({ styles: this.styles, numbering: this.numbering, rels: this.rels(part), showDeletions: this.options.showDeletions });
  }

  /** The main body, flattened for layout. Cached until `invalidate()`. */
  body(): Story {
    if (!this.bodyCache) {
      const body = firstChild(this.documentXml.root, 'w:body');
      if (!body) throw new DocxError('document.xml has no w:body');
      this.bodyCache = this.reader(this.mainPart).readStory(body);
    }
    return this.bodyCache;
  }

  /** Part name of a header/footer referenced by a relationship id in the main document. */
  headerFooterPart(rId: string): string | undefined {
    const rel = this.rels(this.mainPart).get(rId);
    return rel && !rel.external ? rel.target : undefined;
  }

  /** Header or footer content. */
  headerFooter(part: string): Story {
    let story = this.storyCache.get(part);
    if (!story) {
      const xml = this.xmlSync(part);
      story = xml ? this.reader(part).readStory(xml.root) : { blocks: [] };
      this.storyCache.set(part, story);
    }
    return story;
  }

  /** Blocks of a footnote or endnote by id. */
  note(type: 'footnote' | 'endnote', id: string): Block[] {
    const part = type === 'footnote' ? this.parts.footnotes : this.parts.endnotes;
    const xml = this.xmlSync(part);
    if (!part || !xml) return [];
    const key = `${part}#${id}`;
    let story = this.storyCache.get(key);
    if (!story) {
      const node = elements(xml.root).find((n) => attr(n, 'w:id') === id);
      story = node ? this.reader(part).readStory(node) : { blocks: [] };
      this.storyCache.set(key, story);
    }
    return story.blocks;
  }

  get defaultTabStop(): number {
    const settings = this.xmlSync(this.parts.settings);
    return parseTwipsMeasure(attr(firstChild(settings?.root, 'w:defaultTabStop'), 'w:val')) ?? 36;
  }

  get evenAndOddHeaders(): boolean {
    const settings = this.xmlSync(this.parts.settings);
    const node = firstChild(settings?.root, 'w:evenAndOddHeaders');
    return !!node && attr(node, 'w:val') !== '0' && attr(node, 'w:val') !== 'false';
  }

  async coreProperties(): Promise<CoreProperties> {
    const part = await this.pkg.targetPart('', 'coreProperties');
    const xml = await this.pkg.tryGetXml(part);
    if (!xml) return {};
    const get = (name: string) => {
      const n = firstChild(xml.root, name);
      return n ? textContent(n) || undefined : undefined;
    };
    return {
      title: get('dc:title'),
      subject: get('dc:subject'),
      creator: get('dc:creator'),
      keywords: get('cp:keywords'),
      description: get('dc:description'),
      lastModifiedBy: get('cp:lastModifiedBy'),
      created: get('dcterms:created'),
      modified: get('dcterms:modified'),
    };
  }

  /** Bytes of an image or other media part. */
  async media(part: string): Promise<Uint8Array> {
    return this.pkg.getBytes(part);
  }

  /** Mark a part as edited: it will be re-serialised on save and caches are refreshed. */
  touch(part: string): void {
    this.pkg.markDirty(part);
    if (part === this.parts.styles || part === this.parts.numbering || part === this.parts.theme) this.reparseDefinitions();
    else this.invalidate();
    for (const l of this.listeners) l(part);
  }

  invalidate(): void {
    this.bodyCache = undefined;
    this.storyCache.clear();
  }

  /** Subscribe to edits; returns an unsubscribe function. */
  onChange(listener: (part: string) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  /** Serialise to .docx bytes. */
  async save(options: { modified?: Date } = {}): Promise<Uint8Array> {
    return this.pkg.save(options);
  }
}

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const BLANK_DOCUMENT = `${DECL}<w:document ${W}><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/></w:sectPr></w:body></w:document>`;

const BLANK_STYLES = `${DECL}<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:color w:val="2F5496"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:color w:val="2F5496"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style></w:styles>`;

const BLANK_SETTINGS = `${DECL}<w:settings ${W}><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;
