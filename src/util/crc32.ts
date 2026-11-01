/** CRC-32 (IEEE 802.3, reflected polynomial 0xEDB88320) as used by ZIP and PNG. */

const TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

export function crc32(data: Uint8Array, previous = 0): number {
  let crc = ~previous >>> 0;
  for (let i = 0; i < data.length; i++) crc = TABLE[(crc ^ data[i]!) & 0xff]! ^ (crc >>> 8);
  return ~crc >>> 0;
}
