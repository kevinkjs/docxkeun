/**
 * A minimal MP4 writer for one H.264 track: samples in decode order, no B-frames,
 * moov before mdat so players can start before the whole file arrives.
 */

const u8 = (...bytes) => Uint8Array.from(bytes);
const u16 = (n) => u8((n >>> 8) & 255, n & 255);
const u32 = (n) => u8((n >>> 24) & 255, (n >>> 16) & 255, (n >>> 8) & 255, n & 255);
const ascii = (s) => Uint8Array.from(s, (c) => c.charCodeAt(0));

function concat(parts) {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}

function box(type, ...parts) {
  const body = concat(parts);
  return concat([u32(body.length + 8), ascii(type), body]);
}

const fullBox = (type, version, flags, ...parts) => box(type, u8(version, (flags >>> 16) & 255, (flags >>> 8) & 255, flags & 255), ...parts);

const MATRIX = concat([u32(0x00010000), u32(0), u32(0), u32(0), u32(0x00010000), u32(0), u32(0), u32(0), u32(0x40000000)]);

/**
 * @param {{ width: number, height: number, timescale: number, description: Uint8Array,
 *           samples: Array<{ size: number, duration: number, key: boolean }>, data: Uint8Array }} track
 */
export function writeMp4({ width, height, timescale, description, samples, data }) {
  const duration = samples.reduce((n, s) => n + s.duration, 0);
  const movieDuration = Math.round((duration / timescale) * 1000);

  // stts: runs of equal durations.
  const runs = [];
  for (const s of samples) {
    const last = runs.at(-1);
    if (last && last.duration === s.duration) last.count++;
    else runs.push({ count: 1, duration: s.duration });
  }
  const keys = samples.flatMap((s, i) => (s.key ? [i + 1] : []));

  const moov = (chunkOffset) => {
    const avc1 = box('avc1',
      u8(0, 0, 0, 0, 0, 0), u16(1),                 // reserved, data_reference_index
      u16(0), u16(0), u32(0), u32(0), u32(0),        // pre_defined, reserved, pre_defined
      u16(width), u16(height),
      u32(0x00480000), u32(0x00480000), u32(0),      // 72 dpi, reserved
      u16(1), new Uint8Array(32), u16(0x0018), u16(0xffff),
      box('avcC', description));
    const stbl = box('stbl',
      fullBox('stsd', 0, 0, u32(1), avc1),
      fullBox('stts', 0, 0, u32(runs.length), ...runs.flatMap((r) => [u32(r.count), u32(r.duration)])),
      fullBox('stss', 0, 0, u32(keys.length), ...keys.map(u32)),
      fullBox('stsc', 0, 0, u32(1), u32(1), u32(samples.length), u32(1)),
      fullBox('stsz', 0, 0, u32(0), u32(samples.length), ...samples.map((s) => u32(s.size))),
      fullBox('stco', 0, 0, u32(1), u32(chunkOffset)));
    const minf = box('minf',
      fullBox('vmhd', 0, 1, u16(0), u16(0), u16(0), u16(0)),
      box('dinf', fullBox('dref', 0, 0, u32(1), fullBox('url ', 0, 1))),
      stbl);
    const mdia = box('mdia',
      fullBox('mdhd', 0, 0, u32(0), u32(0), u32(timescale), u32(duration), u16(0x55c4), u16(0)),
      fullBox('hdlr', 0, 0, u32(0), ascii('vide'), u32(0), u32(0), u32(0), ascii('VideoHandler\0')),
      minf);
    const tkhd = fullBox('tkhd', 0, 3,
      u32(0), u32(0), u32(1), u32(0), u32(movieDuration), u32(0), u32(0),
      u16(0), u16(0), u16(0), u16(0), MATRIX, u32(width << 16), u32(height << 16));
    const mvhd = fullBox('mvhd', 0, 0,
      u32(0), u32(0), u32(1000), u32(movieDuration), u32(0x00010000), u16(0x0100), u16(0), u32(0), u32(0),
      MATRIX, new Uint8Array(24), u32(2));
    return box('moov', mvhd, box('trak', tkhd, mdia));
  };

  const ftyp = box('ftyp', ascii('isom'), u32(512), ascii('isom'), ascii('iso2'), ascii('avc1'), ascii('mp41'));
  const size = moov(0).length;
  const header = concat([ftyp, moov(ftyp.length + size + 8)]);
  return concat([header, u32(data.length + 8), ascii('mdat'), data]);
}
