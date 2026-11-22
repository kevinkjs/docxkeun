/**
 * Re-using paragraph layout between runs. Measuring characters and breaking lines is most of
 * layout's time; a paragraph whose content, width and surroundings are unchanged since the last
 * run gets its previous atoms and lines back, so typing re-measures only the paragraph typed in.
 */
import type { DocParagraph } from '../model/content.js';
import type { XmlElement } from '../xml/index.js';
import type { Atom, Line } from './lines.js';

/** @internal */
export interface CachedParagraph {
  atoms: Atom[];
  lines: Line[];
  /** The range tracker's state after the paragraph. */
  trackerAfter: unknown;
}

/** Keep one per document and pass it to every `layoutDocument` call (the viewer does). */
export class LayoutCache {
  private paragraphs = new WeakMap<XmlElement, { key: string; value: CachedParagraph }>();
  private owners: unknown[] = [];
  hits = 0;
  misses = 0;

  /** @internal Start a run: anything resolved against other styles, theme or fonts is dropped. */
  begin(owners: unknown[]): void {
    if (owners.length !== this.owners.length || owners.some((o, i) => o !== this.owners[i])) {
      this.paragraphs = new WeakMap();
      this.owners = owners;
    }
    this.hits = 0;
    this.misses = 0;
  }

  /** @internal */
  get(el: XmlElement, key: string): CachedParagraph | undefined {
    const e = this.paragraphs.get(el);
    if (e && e.key === key) {
      this.hits++;
      return e.value;
    }
    this.misses++;
    return undefined;
  }

  /** @internal */
  set(el: XmlElement, key: string, value: CachedParagraph): void {
    this.paragraphs.set(el, { key, value });
  }

  clear(): void {
    this.paragraphs = new WeakMap();
  }
}

const isXml = (v: unknown): boolean => !!v && typeof v === 'object' && (v as { type?: unknown }).type === 'element' && Array.isArray((v as { children?: unknown }).children);

/** Everything about a paragraph that its atoms and lines depend on (XML nodes left out: their content is in the model). */
export function paragraphKey(para: DocParagraph, parts: unknown[]): string {
  return JSON.stringify(parts) + JSON.stringify(para, (_k, v: unknown) => (isXml(v) ? undefined : v));
}
