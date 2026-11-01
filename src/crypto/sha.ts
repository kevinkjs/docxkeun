/**
 * SHA-1 and SHA-2 (FIPS 180-4). Synchronous so key derivations with 100,000 rounds stay fast.
 * The SHA-2 round constants and initial values are derived from the primes, as the standard
 * defines them, instead of being typed in by hand.
 */

export type HashName = 'SHA-1' | 'SHA-256' | 'SHA-384' | 'SHA-512';

function primes(count: number): bigint[] {
  const out: bigint[] = [];
  for (let n = 2; out.length < count; n++) {
    let prime = true;
    for (let d = 2; d * d <= n; d++) {
      if (n % d === 0) {
        prime = false;
        break;
      }
    }
    if (prime) out.push(BigInt(n));
  }
  return out;
}

function iroot(value: bigint, k: bigint): bigint {
  // Integer k-th root by Newton's method.
  if (value < 2n) return value;
  let x = 1n << (BigInt(value.toString(2).length) / k + 1n);
  for (;;) {
    const y = ((k - 1n) * x + value / x ** (k - 1n)) / k;
    if (y >= x) return x;
    x = y;
  }
}

/** First `bits` bits of the fractional part of the k-th root of each prime. */
function fractionalRoots(ps: bigint[], k: bigint, bits: bigint): bigint[] {
  const mask = (1n << bits) - 1n;
  return ps.map((p) => iroot(p << (bits * k), k) & mask);
}

const PRIMES = primes(80);
const K256 = Uint32Array.from(fractionalRoots(PRIMES.slice(0, 64), 3n, 32n).map(Number));
const H256 = Uint32Array.from(fractionalRoots(PRIMES.slice(0, 8), 2n, 32n).map(Number));
const K512_BIG = fractionalRoots(PRIMES, 3n, 64n);
const H512_BIG = fractionalRoots(PRIMES.slice(0, 8), 2n, 64n);
const H384_BIG = fractionalRoots(PRIMES.slice(8, 16), 2n, 64n);

function split64(values: bigint[]): Int32Array {
  const out = new Int32Array(values.length * 2);
  values.forEach((v, i) => {
    out[i * 2] = Number((v >> 32n) & 0xffffffffn) | 0;
    out[i * 2 + 1] = Number(v & 0xffffffffn) | 0;
  });
  return out;
}

const K512 = split64(K512_BIG);
const H512 = split64(H512_BIG);
const H384 = split64(H384_BIG);

/** Message padding shared by all four hashes (big-endian bit length). */
function pad(data: Uint8Array, blockSize: 64 | 128): Uint8Array {
  const lengthBytes = blockSize === 64 ? 8 : 16;
  const total = Math.ceil((data.length + 1 + lengthBytes) / blockSize) * blockSize;
  const out = new Uint8Array(total);
  out.set(data);
  out[data.length] = 0x80;
  const bitLength = data.length * 8;
  const view = new DataView(out.buffer);
  view.setUint32(total - 4, bitLength >>> 0, false);
  view.setUint32(total - 8, Math.floor(bitLength / 0x1_0000_0000), false);
  return out;
}

function sha1(data: Uint8Array): Uint8Array {
  const m = pad(data, 64);
  const view = new DataView(m.buffer);
  let h0 = 0x67452301,
    h1 = 0xefcdab89,
    h2 = 0x98badcfe,
    h3 = 0x10325476,
    h4 = 0xc3d2e1f0;
  const w = new Uint32Array(80);
  for (let off = 0; off < m.length; off += 64) {
    for (let i = 0; i < 16; i++) w[i] = view.getUint32(off + i * 4, false);
    for (let i = 16; i < 80; i++) {
      const x = w[i - 3]! ^ w[i - 8]! ^ w[i - 14]! ^ w[i - 16]!;
      w[i] = (x << 1) | (x >>> 31);
    }
    let a = h0,
      b = h1,
      c = h2,
      d = h3,
      e = h4;
    for (let i = 0; i < 80; i++) {
      let f: number, k: number;
      if (i < 20) {
        f = (b & c) | (~b & d);
        k = 0x5a827999;
      } else if (i < 40) {
        f = b ^ c ^ d;
        k = 0x6ed9eba1;
      } else if (i < 60) {
        f = (b & c) | (b & d) | (c & d);
        k = 0x8f1bbcdc;
      } else {
        f = b ^ c ^ d;
        k = 0xca62c1d6;
      }
      const t = (((a << 5) | (a >>> 27)) + f + e + k + w[i]!) | 0;
      e = d;
      d = c;
      c = (b << 30) | (b >>> 2);
      b = a;
      a = t;
    }
    h0 = (h0 + a) | 0;
    h1 = (h1 + b) | 0;
    h2 = (h2 + c) | 0;
    h3 = (h3 + d) | 0;
    h4 = (h4 + e) | 0;
  }
  const out = new Uint8Array(20);
  const ov = new DataView(out.buffer);
  [h0, h1, h2, h3, h4].forEach((h, i) => ov.setUint32(i * 4, h >>> 0, false));
  return out;
}

function sha256(data: Uint8Array): Uint8Array {
  const m = pad(data, 64);
  const view = new DataView(m.buffer);
  const h = Uint32Array.from(H256);
  const w = new Uint32Array(64);
  for (let off = 0; off < m.length; off += 64) {
    for (let i = 0; i < 16; i++) w[i] = view.getUint32(off + i * 4, false);
    for (let i = 16; i < 64; i++) {
      const x = w[i - 15]!;
      const y = w[i - 2]!;
      const s0 = ((x >>> 7) | (x << 25)) ^ ((x >>> 18) | (x << 14)) ^ (x >>> 3);
      const s1 = ((y >>> 17) | (y << 15)) ^ ((y >>> 19) | (y << 13)) ^ (y >>> 10);
      w[i] = (w[i - 16]! + s0 + w[i - 7]! + s1) | 0;
    }
    let a = h[0]!,
      b = h[1]!,
      c = h[2]!,
      d = h[3]!,
      e = h[4]!,
      f = h[5]!,
      g = h[6]!,
      hh = h[7]!;
    for (let i = 0; i < 64; i++) {
      const S1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
      const ch = (e & f) ^ (~e & g);
      const t1 = (hh + S1 + ch + K256[i]! + w[i]!) | 0;
      const S0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
      const maj = (a & b) ^ (a & c) ^ (b & c);
      const t2 = (S0 + maj) | 0;
      hh = g;
      g = f;
      f = e;
      e = (d + t1) | 0;
      d = c;
      c = b;
      b = a;
      a = (t1 + t2) | 0;
    }
    h[0] = (h[0]! + a) | 0;
    h[1] = (h[1]! + b) | 0;
    h[2] = (h[2]! + c) | 0;
    h[3] = (h[3]! + d) | 0;
    h[4] = (h[4]! + e) | 0;
    h[5] = (h[5]! + f) | 0;
    h[6] = (h[6]! + g) | 0;
    h[7] = (h[7]! + hh) | 0;
  }
  const out = new Uint8Array(32);
  const ov = new DataView(out.buffer);
  for (let i = 0; i < 8; i++) ov.setUint32(i * 4, h[i]!, false);
  return out;
}

/** SHA-512 core over 64-bit words stored as (hi, lo) int32 pairs. */
function sha512Core(data: Uint8Array, init: Int32Array, outBytes: number): Uint8Array {
  const m = pad(data, 128);
  const view = new DataView(m.buffer);
  const h = Int32Array.from(init);
  const w = new Int32Array(160);
  for (let off = 0; off < m.length; off += 128) {
    for (let i = 0; i < 32; i++) w[i] = view.getInt32(off + i * 4, false);
    for (let i = 16; i < 80; i++) {
      // s0 = rotr1 ^ rotr8 ^ shr7 of w[i-15]
      let xh = w[(i - 15) * 2]!,
        xl = w[(i - 15) * 2 + 1]!;
      const s0h = ((xh >>> 1) | (xl << 31)) ^ ((xh >>> 8) | (xl << 24)) ^ (xh >>> 7);
      const s0l = ((xl >>> 1) | (xh << 31)) ^ ((xl >>> 8) | (xh << 24)) ^ ((xl >>> 7) | (xh << 25));
      // s1 = rotr19 ^ rotr61 ^ shr6 of w[i-2]
      xh = w[(i - 2) * 2]!;
      xl = w[(i - 2) * 2 + 1]!;
      const s1h = ((xh >>> 19) | (xl << 13)) ^ ((xl >>> 29) | (xh << 3)) ^ (xh >>> 6);
      const s1l = ((xl >>> 19) | (xh << 13)) ^ ((xh >>> 29) | (xl << 3)) ^ ((xl >>> 6) | (xh << 26));
      // w[i] = w[i-16] + s0 + w[i-7] + s1
      let lo = (w[(i - 16) * 2 + 1]! >>> 0) + (s0l >>> 0) + (w[(i - 7) * 2 + 1]! >>> 0) + (s1l >>> 0);
      let hi = w[(i - 16) * 2]! + s0h + w[(i - 7) * 2]! + s1h + Math.floor(lo / 0x1_0000_0000);
      w[i * 2] = hi | 0;
      w[i * 2 + 1] = lo | 0;
    }
    let ah = h[0]!, al = h[1]!, bh = h[2]!, bl = h[3]!, ch = h[4]!, cl = h[5]!, dh = h[6]!, dl = h[7]!;
    let eh = h[8]!, el = h[9]!, fh = h[10]!, fl = h[11]!, gh = h[12]!, gl = h[13]!, hh = h[14]!, hl = h[15]!;
    for (let i = 0; i < 80; i++) {
      // Σ1(e) = rotr14 ^ rotr18 ^ rotr41
      const S1h = ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23));
      const S1l = ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23));
      const chh = (eh & fh) ^ (~eh & gh);
      const chl = (el & fl) ^ (~el & gl);
      // t1 = h + Σ1 + ch + K[i] + w[i]
      let lo = (hl >>> 0) + (S1l >>> 0) + (chl >>> 0) + (K512[i * 2 + 1]! >>> 0) + (w[i * 2 + 1]! >>> 0);
      const t1h = (hh + S1h + chh + K512[i * 2]! + w[i * 2]! + Math.floor(lo / 0x1_0000_0000)) | 0;
      const t1l = lo | 0;
      // Σ0(a) = rotr28 ^ rotr34 ^ rotr39
      const S0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25));
      const S0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25));
      const majh = (ah & bh) ^ (ah & ch) ^ (bh & ch);
      const majl = (al & bl) ^ (al & cl) ^ (bl & cl);
      lo = (S0l >>> 0) + (majl >>> 0);
      const t2h = (S0h + majh + Math.floor(lo / 0x1_0000_0000)) | 0;
      const t2l = lo | 0;
      hh = gh; hl = gl;
      gh = fh; gl = fl;
      fh = eh; fl = el;
      lo = (dl >>> 0) + (t1l >>> 0);
      eh = (dh + t1h + Math.floor(lo / 0x1_0000_0000)) | 0;
      el = lo | 0;
      dh = ch; dl = cl;
      ch = bh; cl = bl;
      bh = ah; bl = al;
      lo = (t1l >>> 0) + (t2l >>> 0);
      ah = (t1h + t2h + Math.floor(lo / 0x1_0000_0000)) | 0;
      al = lo | 0;
    }
    const add = (idx: number, xh: number, xl: number) => {
      const lo = (h[idx + 1]! >>> 0) + (xl >>> 0);
      h[idx] = (h[idx]! + xh + Math.floor(lo / 0x1_0000_0000)) | 0;
      h[idx + 1] = lo | 0;
    };
    add(0, ah, al); add(2, bh, bl); add(4, ch, cl); add(6, dh, dl);
    add(8, eh, el); add(10, fh, fl); add(12, gh, gl); add(14, hh, hl);
  }
  const out = new Uint8Array(64);
  const ov = new DataView(out.buffer);
  for (let i = 0; i < 16; i++) ov.setInt32(i * 4, h[i]!, false);
  return out.slice(0, outBytes);
}

export function hash(name: HashName, data: Uint8Array): Uint8Array {
  switch (name) {
    case 'SHA-1':
      return sha1(data);
    case 'SHA-256':
      return sha256(data);
    case 'SHA-384':
      return sha512Core(data, H384, 48);
    case 'SHA-512':
      return sha512Core(data, H512, 64);
  }
}

export function hashLength(name: HashName): number {
  return { 'SHA-1': 20, 'SHA-256': 32, 'SHA-384': 48, 'SHA-512': 64 }[name];
}

export function hashBlockSize(name: HashName): number {
  return name === 'SHA-1' || name === 'SHA-256' ? 64 : 128;
}

/** Hash of several byte arrays concatenated. */
export function hashConcat(name: HashName, ...parts: Uint8Array[]): Uint8Array {
  let length = 0;
  for (const p of parts) length += p.length;
  const buf = new Uint8Array(length);
  let o = 0;
  for (const p of parts) {
    buf.set(p, o);
    o += p.length;
  }
  return hash(name, buf);
}

/** HMAC (RFC 2104). */
export function hmac(name: HashName, key: Uint8Array, data: Uint8Array): Uint8Array {
  const block = hashBlockSize(name);
  let k = key.length > block ? hash(name, key) : key;
  const padded = new Uint8Array(block);
  padded.set(k);
  const ipad = new Uint8Array(block);
  const opad = new Uint8Array(block);
  for (let i = 0; i < block; i++) {
    ipad[i] = padded[i]! ^ 0x36;
    opad[i] = padded[i]! ^ 0x5c;
  }
  k = hashConcat(name, ipad, data);
  return hashConcat(name, opad, k);
}

export function hashNameFromString(s: string): HashName | undefined {
  const n = s.toUpperCase().replace(/^SHA(\d)/, 'SHA-$1');
  return n === 'SHA-1' || n === 'SHA-256' || n === 'SHA-384' || n === 'SHA-512' ? n : undefined;
}
