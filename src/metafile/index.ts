/**
 * Windows metafiles (EMF and WMF), the vector pictures old Word documents embed for equations,
 * Visio drawings and pasted charts: format detection, natural size, and rendering into the
 * layout display list so the SVG viewer and the PDF exporter draw them like any other shape.
 */
import type { DisplayItem, FontProvider } from '../layout/types.js';
import { emfSize, isEmf, renderEmf } from './emf.js';
import { readWmfHeader, renderWmf, wmfSize } from './wmf.js';

export type MetafileKind = 'emf' | 'wmf';

export function metafileKind(bytes: Uint8Array): MetafileKind | undefined {
  if (isEmf(bytes)) return 'emf';
  if (readWmfHeader(bytes)) return 'wmf';
  return undefined;
}

/**
 * Natural size in points. EMF: the header's frame rectangle (0.01 mm). WMF: the placeable
 * header's bounding box at its units per inch, else the window extent at 1440 units per inch.
 */
export function metafileSize(bytes: Uint8Array): { width: number; height: number } | undefined {
  const kind = metafileKind(bytes);
  if (kind === 'emf') return emfSize(bytes);
  if (kind === 'wmf') return wmfSize(bytes);
  return undefined;
}

/**
 * Draw a metafile scaled to fill `frame` (page points, y down). Returns no items for data that is
 * not a metafile (or an EMF+-only picture, whose drawing lives entirely in EMF+ records);
 * damaged records are skipped rather than failing the whole picture.
 */
export function renderMetafile(bytes: Uint8Array, frame: { x: number; y: number; w: number; h: number }, fonts: FontProvider): DisplayItem[] {
  const kind = metafileKind(bytes);
  if (kind === 'emf') return renderEmf(bytes, frame, fonts);
  if (kind === 'wmf') return renderWmf(bytes, frame, fonts);
  return [];
}
