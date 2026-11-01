/**
 * Document → PDF, drawn from the same layout as the screen view: selectable text in embedded
 * font subsets, vector borders, images, clickable links, named destinations for bookmarks,
 * an outline from headings, and comments as highlight annotations with threaded replies.
 */
import { readComments, type CommentInfo } from '../comments/read.js';
import { randomBytes } from '../crypto/index.js';
import type { FontRegistry } from '../fonts/registry.js';
import { layoutDocument } from '../layout/engine.js';
import type { DisplayItem, LayoutResult, TextItem } from '../layout/types.js';
import type { DocxDocument } from '../model/document.js';
import { addImage } from './images.js';
import { PdfFont } from './fonts.js';
import { name, num, pdfDate, str, type PdfDict, type PdfRef, type PdfValue } from './objects.js';
import { PdfWriter, type PdfEncryption } from './writer.js';

export interface PdfExportOptions {
  /** The fonts used for layout; PDF embeds subsets of them. */
  fonts: FontRegistry;
  /** Reuse a layout you already have (it must have been made with the same fonts). */
  layout?: LayoutResult;
  /** Include comments as annotations (default true). */
  comments?: boolean;
  /** Build a bookmark outline from headings (default true). */
  outline?: boolean;
  /** URL schemes allowed in link annotations (default http, https, mailto, tel). */
  allowedProtocols?: string[];
  metadata?: { title?: string; author?: string; subject?: string; keywords?: string };
  /** Applied while writing (see docxkeun/security's pdfSecurity). */
  encryption?: PdfEncryption;
}

function rgb(hex: string): string {
  const c = /^[0-9a-f]{6}$/i.test(hex) ? hex : '000000';
  return [0, 2, 4].map((i) => num(Number.parseInt(c.slice(i, i + 2), 16) / 255)).join(' ');
}

function safeLink(url: string | undefined, allowed: string[]): string | undefined {
  if (!url) return undefined;
  try {
    const u = new URL(url);
    return allowed.includes(u.protocol) ? u.href : undefined;
  } catch {
    return undefined;
  }
}

class PageContent {
  readonly ops: string[] = [];
  readonly fonts = new Map<string, PdfRef>();
  readonly images = new Map<string, PdfRef>();
}

export async function exportPdf(doc: DocxDocument, options: PdfExportOptions): Promise<Uint8Array> {
  const layout = options.layout ?? layoutDocument(doc, { fonts: options.fonts });
  const writer = new PdfWriter();
  const catalogRef = writer.alloc();
  const pagesRef = writer.alloc();
  const pageRefs = layout.pages.map(() => writer.alloc());
  const allowed = options.allowedProtocols ?? ['http:', 'https:', 'mailto:', 'tel:'];

  const fonts = new Map<string, { font: PdfFont; ref: PdfRef }>();
  const fontFor = (item: TextItem) => {
    const face = item.font.face;
    if (!face) throw new Error('PDF export needs registered font files (FontRegistry), not system font metrics');
    let entry = fonts.get(face.id);
    if (!entry) {
      entry = { font: new PdfFont(face, `F${fonts.size + 1}`), ref: writer.alloc() };
      fonts.set(face.id, entry);
    }
    return entry;
  };

  const images = new Map<string, { ref: PdfRef; resource: string } | null>();
  const imageFor = async (target: string | undefined, external: boolean) => {
    if (!target || external) return null;
    if (!images.has(target)) {
      try {
        const { ref } = await addImage(writer, await doc.media(target));
        images.set(target, { ref, resource: `Im${images.size + 1}` });
      } catch {
        images.set(target, null); // unsupported format (EMF, WMF, GIF...): drawn as a placeholder
      }
    }
    return images.get(target) ?? null;
  };

  const pageAnnots: PdfRef[][] = layout.pages.map(() => []);
  if (options.comments ?? true) addComments(writer, readComments(doc), layout, pageAnnots);

  for (const page of layout.pages) {
    const content = new PageContent();
    const H = page.height;
    for (const item of page.items) await drawItem(item, content, H, fontFor, imageFor);
    const stream = await writer.stream({}, new TextEncoder().encode(content.ops.join('\n')));
    const fontRes: PdfDict = {};
    for (const [res, ref] of content.fonts) fontRes[res] = ref;
    const xobjects: PdfDict = {};
    for (const [res, ref] of content.images) xobjects[res] = ref;

    // Link annotations.
    for (const link of page.links) {
      const rect = [link.x, H - link.y - link.h, link.x + link.w, H - link.y];
      let action: PdfDict | undefined;
      let dest: PdfValue | undefined;
      const url = safeLink(link.target.url, allowed);
      if (url) action = { S: name('URI'), URI: str(url) };
      else if (link.target.anchor) {
        const anchor = layout.anchors.get(link.target.anchor);
        if (anchor) dest = [pageRefs[anchor.page]!, name('XYZ'), anchor.x, layout.pages[anchor.page]!.height - anchor.y, null];
      }
      if (!action && !dest) continue;
      pageAnnots[page.index]!.push(
        writer.add({ Type: name('Annot'), Subtype: name('Link'), Rect: rect, Border: [0, 0, 0], F: 4, A: action, Dest: dest, Contents: link.target.tooltip ? str(link.target.tooltip) : undefined }),
      );
    }

    writer.set(pageRefs[page.index]!, {
      Type: name('Page'),
      Parent: pagesRef,
      MediaBox: [0, 0, page.width, page.height],
      Resources: { Font: fontRes, XObject: Object.keys(xobjects).length ? xobjects : undefined, ProcSet: [name('PDF'), name('Text'), name('ImageB'), name('ImageC')] },
      Contents: stream,
      Annots: pageAnnots[page.index]!.length ? pageAnnots[page.index] : undefined,
    });
  }

  for (const { font, ref } of fonts.values()) await font.write(writer, ref);

  writer.set(pagesRef, { Type: name('Pages'), Kids: pageRefs, Count: pageRefs.length });

  const catalog: PdfDict = { Type: name('Catalog'), Pages: pagesRef, ViewerPreferences: { DisplayDocTitle: true } };
  const outline = (options.outline ?? true) ? addOutline(writer, layout, pageRefs) : undefined;
  if (outline) {
    catalog.Outlines = outline;
    catalog.PageMode = name('UseOutlines');
  }
  const dests = [...layout.anchors]
    .filter(([n]) => n && n !== '_GoBack')
    .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
    .flatMap(([n, a]) => [str(n), [pageRefs[a.page]!, name('XYZ'), a.x, layout.pages[a.page]!.height - a.y, null]] as PdfValue[]);
  if (dests.length) catalog.Names = { Dests: { Names: dests } };
  if (options.encryption?.extensions) catalog.Extensions = options.encryption.extensions;
  writer.set(catalogRef, catalog);

  const core = await doc.coreProperties();
  const meta = options.metadata ?? {};
  const now = pdfDate(new Date());
  const text = (v: string | undefined) => (v ? str(v) : undefined);
  const info = writer.add({
    Title: text(meta.title ?? core.title),
    Author: text(meta.author ?? core.creator),
    Subject: text(meta.subject ?? core.subject),
    Keywords: text(meta.keywords ?? core.keywords),
    Creator: str('docxkeun'),
    Producer: str('docxkeun'),
    CreationDate: str(now),
    ModDate: str(now),
  });
  return writer.build(catalogRef, info, randomBytes(16), options.encryption);
}

async function drawItem(
  item: DisplayItem,
  content: PageContent,
  H: number,
  fontFor: (item: TextItem) => { font: PdfFont; ref: PdfRef },
  imageFor: (target: string | undefined, external: boolean) => Promise<{ ref: PdfRef; resource: string } | null>,
): Promise<void> {
  const ops = content.ops;
  switch (item.type) {
    case 'rect':
      ops.push(`${rgb(item.fill)} rg ${num(item.x)} ${num(H - item.y - item.h)} ${num(item.w)} ${num(item.h)} re f`);
      break;
    case 'line': {
      const dash = item.style === 'dashed' ? `[${num(item.width * 3)} ${num(item.width * 2)}] 0 d` : item.style === 'dotted' ? `[${num(item.width)} ${num(item.width)}] 0 d` : '[] 0 d';
      const seg = (x1: number, y1: number, x2: number, y2: number, w: number) =>
        ops.push(`${num(w)} w ${dash} ${rgb(item.color)} RG ${num(x1)} ${num(H - y1)} m ${num(x2)} ${num(H - y2)} l S`);
      if (item.style === 'double') {
        const horizontal = Math.abs(item.y2 - item.y1) < Math.abs(item.x2 - item.x1);
        const w = Math.max(item.width / 3, 0.25);
        const d = (item.width / 2 - w / 2 + w) / 2;
        if (horizontal) {
          seg(item.x1, item.y1 - d, item.x2, item.y2 - d, w);
          seg(item.x1, item.y1 + d, item.x2, item.y2 + d, w);
        } else {
          seg(item.x1 - d, item.y1, item.x2 - d, item.y2, w);
          seg(item.x1 + d, item.y1, item.x2 + d, item.y2, w);
        }
      } else seg(item.x1, item.y1, item.x2, item.y2, item.width);
      break;
    }
    case 'image': {
      const img = await imageFor(item.target, item.external);
      if (!img) {
        ops.push(`0.945 0.953 0.957 rg ${num(item.x)} ${num(H - item.y - item.h)} ${num(item.w)} ${num(item.h)} re f`);
        break;
      }
      content.images.set(img.resource, img.ref);
      ops.push(`q ${num(item.w)} 0 0 ${num(item.h)} ${num(item.x)} ${num(H - item.y - item.h)} cm /${img.resource} Do Q`);
      break;
    }
    case 'text': {
      if (!item.chars.length) break;
      const { font, ref } = fontFor(item);
      content.fonts.set(font.resource, ref);
      const scale = item.scale || 1;
      const parts: string[] = [];
      let hex = '';
      let x = item.xs[0]!;
      item.chars.forEach((ch, i) => {
        const gid = item.glyphs[i] ?? 0;
        font.use(gid, ch);
        hex += gid.toString(16).padStart(4, '0');
        const natural = (font.width(gid) * item.size * scale) / 1000;
        x += natural;
        const next = item.xs[i + 1];
        if (next !== undefined) {
          const diff = next - x;
          if (Math.abs(diff) > 0.001) {
            parts.push(`<${hex}>`, num((-diff * 1000) / (item.size * scale)));
            hex = '';
            x = next;
          }
        }
      });
      if (hex) parts.push(`<${hex}>`);
      const skew = item.font.syntheticItalic ? 0.21 : 0;
      const bold = item.font.syntheticBold ? ` 2 Tr ${num(item.size * 0.03)} w ${rgb(item.color)} RG` : '';
      const tz = scale !== 1 ? ` ${num(scale * 100)} Tz` : '';
      ops.push(`BT /${font.resource} ${num(item.size)} Tf${tz}${bold} ${rgb(item.color)} rg 1 0 ${num(skew)} 1 ${num(item.xs[0]!)} ${num(H - item.y)} Tm [${parts.join(' ')}] TJ ET`);
      break;
    }
  }
}

function addComments(writer: PdfWriter, comments: CommentInfo[], layout: LayoutResult, pageAnnots: PdfRef[][]): void {
  const byId = new Map(comments.map((c) => [c.id, c]));
  const date = (c: CommentInfo) => {
    const d = new Date(c.dateUtc ?? c.date ?? '');
    return Number.isNaN(d.getTime()) ? undefined : str(pdfDate(d));
  };
  const parents = new Map<string, PdfRef>();
  for (const page of layout.pages) {
    const groups = new Map<string, Array<{ x: number; y: number; w: number; h: number }>>();
    for (const area of page.comments) {
      if (area.w <= 0 || !byId.has(area.id) || byId.get(area.id)!.parentId) continue;
      const list = groups.get(area.id) ?? [];
      list.push(area);
      groups.set(area.id, list);
    }
    const H = page.height;
    for (const [id, rects] of groups) {
      const c = byId.get(id)!;
      const quads = rects.flatMap((r) => [r.x, H - r.y, r.x + r.w, H - r.y, r.x, H - r.y - r.h, r.x + r.w, H - r.y - r.h]);
      const x1 = Math.min(...rects.map((r) => r.x));
      const x2 = Math.max(...rects.map((r) => r.x + r.w));
      const y1 = Math.min(...rects.map((r) => H - r.y - r.h));
      const y2 = Math.max(...rects.map((r) => H - r.y));
      const annot = writer.alloc();
      const popup = writer.add({ Type: name('Annot'), Subtype: name('Popup'), Parent: annot, Rect: [page.width - 200, y2 - 120, page.width - 10, y2], Open: false });
      writer.set(annot, {
        Type: name('Annot'),
        Subtype: name('Highlight'),
        Rect: [x1, y1, x2, y2],
        QuadPoints: quads,
        C: [1, 0.82, 0],
        F: 4,
        Contents: str(c.text),
        T: str(c.author),
        M: date(c),
        CreationDate: date(c),
        NM: str(`docxkeun-comment-${c.id}-p${page.index}`),
        Subj: str('Comment'),
        Popup: popup,
      });
      pageAnnots[page.index]!.push(annot, popup);
      if (!parents.has(id)) parents.set(id, annot);
      if (c.resolved) {
        pageAnnots[page.index]!.push(
          writer.add({ Type: name('Annot'), Subtype: name('Text'), Rect: [x1, y1, x2, y2], IRT: annot, State: str('Completed'), StateModel: str('Review'), T: str(c.author), F: 30, Contents: str('Completed') }),
        );
      }
    }
  }
  // Replies are text annotations "in reply to" the root of their thread.
  const rootOf = (c: CommentInfo): CommentInfo => (c.parentId && byId.has(c.parentId) ? rootOf(byId.get(c.parentId)!) : c);
  for (const c of comments) {
    if (!c.parentId) continue;
    const parent = parents.get(rootOf(c).id);
    const anchor = layout.commentAnchors.get(rootOf(c).id);
    if (!parent || !anchor) continue;
    const H = layout.pages[anchor.page]!.height;
    const r = anchor.rect;
    pageAnnots[anchor.page]!.push(
      writer.add({
        Type: name('Annot'),
        Subtype: name('Text'),
        // Zero-size: the reply lives in its thread, not as an icon on the page.
        Rect: [r.x, H - r.y, r.x, H - r.y],
        IRT: parent,
        RT: name('R'),
        F: 28,
        Contents: str(c.text),
        T: str(c.author),
        M: date(c),
        CreationDate: date(c),
        NM: str(`docxkeun-comment-${c.id}`),
        Name: name('Comment'),
        Open: false,
      }),
    );
  }
}

function addOutline(writer: PdfWriter, layout: LayoutResult, pageRefs: PdfRef[]): PdfRef | undefined {
  interface Node {
    title: string;
    level: number;
    dest: PdfValue;
    ref: PdfRef;
    children: Node[];
    parent?: Node;
  }
  const root: Node = { title: '', level: -1, dest: null, ref: writer.alloc(), children: [] };
  let last = root;
  for (const page of layout.pages) {
    for (const h of page.headings) {
      const node: Node = { title: h.title, level: h.level, dest: [pageRefs[page.index]!, name('XYZ'), h.x, page.height - h.y, null], ref: writer.alloc(), children: [] };
      let parent = last;
      while (parent !== root && parent.level >= node.level) parent = parent.parent!;
      node.parent = parent;
      parent.children.push(node);
      last = node;
    }
  }
  if (!root.children.length) return undefined;
  const count = (n: Node): number => n.children.reduce((s, c) => s + 1 + count(c), 0);
  const write = (n: Node) => {
    n.children.forEach((c, i) => {
      writer.set(c.ref, {
        Title: str(c.title),
        Parent: n.ref,
        Prev: n.children[i - 1]?.ref,
        Next: n.children[i + 1]?.ref,
        First: c.children[0]?.ref,
        Last: c.children[c.children.length - 1]?.ref,
        Count: c.children.length ? count(c) : undefined,
        Dest: c.dest,
      });
      write(c);
    });
  };
  write(root);
  writer.set(root.ref, { Type: name('Outlines'), First: root.children[0]!.ref, Last: root.children[root.children.length - 1]!.ref, Count: count(root) });
  return root.ref;
}
