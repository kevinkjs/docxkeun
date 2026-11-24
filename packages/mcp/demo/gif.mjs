/**
 * A small animated GIF writer for screen recordings (runs in the browser or Node).
 *
 * One palette for the whole recording (median cut over sample frames, no dithering, so still parts
 * stay still), and each frame stores only the rectangle that changed, with unchanged pixels
 * transparent. Frames that change nothing lengthen the previous frame instead.
 */

const TRANSPARENT = 255; // palette index left free for "unchanged"
const key15 = (r, g, b) => ((r >> 3) << 10) | ((g >> 3) << 5) | (b >> 3);

/** Count colours (5 bits per channel) in RGBA pixels, into `counts` (32768 entries). */
export function countColors(rgba, counts = new Uint32Array(32768)) {
  for (let i = 0; i < rgba.length; i += 4) counts[key15(rgba[i], rgba[i + 1], rgba[i + 2])]++;
  return counts;
}

/**
 * A palette of up to 255 colours by median cut over the counted colours, and a table from each
 * 15-bit colour to its nearest palette entry.
 */
export function buildPalette(counts) {
  const colors = [];
  for (let k = 0; k < 32768; k++) if (counts[k]) colors.push(k);
  const channel = (k, c) => (c === 0 ? k >> 10 : c === 1 ? (k >> 5) & 31 : k & 31);
  const describe = (list) => {
    let total = 0;
    const lo = [31, 31, 31];
    const hi = [0, 0, 0];
    for (const k of list) {
      total += counts[k];
      for (let c = 0; c < 3; c++) {
        const v = channel(k, c);
        if (v < lo[c]) lo[c] = v;
        if (v > hi[c]) hi[c] = v;
      }
    }
    const ranges = hi.map((h, c) => h - lo[c]);
    const wide = ranges.indexOf(Math.max(...ranges));
    return { list, total, wide, range: ranges[wide] };
  };
  const boxes = [describe(colors)];
  while (boxes.length < 255) {
    // Split the box with the most pixels times the widest spread.
    let pick = -1;
    let best = 0;
    boxes.forEach((b, i) => {
      const score = b.total * b.range;
      if (b.list.length > 1 && score > best) {
        best = score;
        pick = i;
      }
    });
    if (pick < 0) break;
    const box = boxes[pick];
    const sorted = [...box.list].sort((a, b) => channel(a, box.wide) - channel(b, box.wide));
    let seen = 0;
    let cut = 1;
    for (; cut < sorted.length; cut++) {
      seen += counts[sorted[cut - 1]];
      if (seen >= box.total / 2) break;
    }
    // Both halves keep at least one colour (a heavy colour at either end stands alone).
    cut = Math.min(Math.max(cut, 1), sorted.length - 1);
    boxes.splice(pick, 1, describe(sorted.slice(0, cut)), describe(sorted.slice(cut)));
  }
  const palette = new Uint8Array(256 * 3);
  boxes.forEach((b, i) => {
    const sum = [0, 0, 0];
    for (const k of b.list) for (let c = 0; c < 3; c++) sum[c] += channel(k, c) * counts[k];
    for (let c = 0; c < 3; c++) palette[i * 3 + c] = Math.round(((sum[c] / b.total) * 255) / 31);
  });
  const lookup = new Uint8Array(32768);
  for (let k = 0; k < 32768; k++) {
    const r = (channel(k, 0) * 255) / 31;
    const g = (channel(k, 1) * 255) / 31;
    const b = (channel(k, 2) * 255) / 31;
    let nearest = 0;
    let distance = Infinity;
    for (let i = 0; i < boxes.length; i++) {
      const dr = r - palette[i * 3];
      const dg = g - palette[i * 3 + 1];
      const db = b - palette[i * 3 + 2];
      const d = 2 * dr * dr + 4 * dg * dg + 3 * db * db;
      if (d < distance) {
        distance = d;
        nearest = i;
      }
    }
    lookup[k] = nearest;
  }
  return { palette, lookup };
}

/** GIF's variable-length LZW, 8-bit indices, as length-prefixed sub-blocks. */
function lzw(indices, out) {
  const min = 8;
  const clear = 1 << min;
  const eoi = clear + 1;
  // Code table keyed by prefix code × 256 + index; a generation number empties it at each clear.
  const codes = lzw.codes ?? (lzw.codes = new Int16Array(4096 * 256));
  const stamps = lzw.stamps ?? (lzw.stamps = new Int32Array(4096 * 256));
  let generation = (lzw.generation ?? 0) + 1;
  let size = min + 1;
  let next = eoi + 1;
  let bits = 0;
  let held = 0;
  const block = [];
  const emit = (code) => {
    held |= code << bits;
    bits += size;
    while (bits >= 8) {
      block.push(held & 255);
      held >>>= 8;
      bits -= 8;
      if (block.length === 255) {
        out.push(255, ...block);
        block.length = 0;
      }
    }
  };
  out.push(min);
  emit(clear);
  let prefix = indices[0];
  for (let i = 1; i < indices.length; i++) {
    const k = indices[i];
    const key = (prefix << 8) | k;
    if (stamps[key] === generation) {
      prefix = codes[key];
      continue;
    }
    emit(prefix);
    if (next === 4096) {
      emit(clear);
      generation++;
      next = eoi + 1;
      size = min + 1;
    } else {
      if (next >= 1 << size) size++;
      codes[key] = next++;
      stamps[key] = generation;
    }
    prefix = k;
  }
  emit(prefix);
  emit(eoi);
  if (bits > 0) block.push(held & 255);
  if (block.length) out.push(block.length, ...block);
  out.push(0);
  lzw.generation = generation;
}

export class GifWriter {
  /** @param {{ width: number, height: number, palette: Uint8Array, lookup: Uint8Array }} options */
  constructor({ width, height, palette, lookup }) {
    this.width = width;
    this.height = height;
    this.lookup = lookup;
    this.chunks = [];
    this.previous = undefined; // indices on screen
    this.pending = undefined; // the last frame, waiting for its length
    const head = [...'GIF89a'].map((c) => c.charCodeAt(0));
    head.push(width & 255, width >> 8, height & 255, height >> 8, 0xf7, 0, 0);
    this.chunks.push(Uint8Array.from(head), palette);
    // Loop forever.
    this.chunks.push(Uint8Array.from([0x21, 0xff, 11, ...[...'NETSCAPE2.0'].map((c) => c.charCodeAt(0)), 3, 1, 0, 0, 0]));
  }

  /** Add a frame (RGBA, width × height) shown for `centiseconds`. */
  add(rgba, centiseconds) {
    const { width, height, lookup } = this;
    const indices = new Uint8Array(width * height);
    for (let p = 0, i = 0; p < indices.length; p++, i += 4) indices[p] = lookup[key15(rgba[i], rgba[i + 1], rgba[i + 2])];
    const before = this.previous;
    let x0 = 0;
    let y0 = 0;
    let x1 = width - 1;
    let y1 = height - 1;
    if (before) {
      x0 = width;
      y0 = height;
      x1 = -1;
      y1 = -1;
      for (let y = 0; y < height; y++) {
        const row = y * width;
        for (let x = 0; x < width; x++) {
          if (indices[row + x] !== before[row + x]) {
            if (x < x0) x0 = x;
            if (x > x1) x1 = x;
            if (y < y0) y0 = y;
            y1 = y;
          }
        }
      }
      if (x1 < 0) {
        // Nothing changed: the frame on screen stays longer.
        this.pending.centiseconds += centiseconds;
        return;
      }
    }
    this.flush();
    const w = x1 - x0 + 1;
    const h = y1 - y0 + 1;
    const rect = new Uint8Array(w * h);
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const p = (y + y0) * width + x + x0;
        rect[y * w + x] = before && indices[p] === before[p] ? TRANSPARENT : indices[p];
      }
    }
    this.pending = { x: x0, y: y0, w, h, rect, centiseconds, transparent: !!before };
    this.previous = indices;
  }

  flush() {
    const f = this.pending;
    if (!f) return;
    const out = [];
    const delay = Math.max(2, Math.round(f.centiseconds));
    // Graphic control: leave the frame in place (disposal 1), transparency for unchanged pixels.
    out.push(0x21, 0xf9, 4, (1 << 2) | (f.transparent ? 1 : 0), delay & 255, delay >> 8, TRANSPARENT, 0);
    out.push(0x2c, f.x & 255, f.x >> 8, f.y & 255, f.y >> 8, f.w & 255, f.w >> 8, f.h & 255, f.h >> 8, 0);
    lzw(f.rect, out);
    this.chunks.push(Uint8Array.from(out));
    this.pending = undefined;
  }

  /** The finished GIF. */
  finish() {
    this.flush();
    this.chunks.push(Uint8Array.of(0x3b));
    const out = new Uint8Array(this.chunks.reduce((n, c) => n + c.length, 0));
    let at = 0;
    for (const c of this.chunks) {
      out.set(c, at);
      at += c.length;
    }
    return out;
  }
}
