/**
 * Embedded font obfuscation (ECMA-376 Part 2 / §17.8.1): the first 32 bytes of the font are
 * XOR-ed with the 16-byte GUID from `w:fontKey`, read in reverse byte order.
 */

function keyBytes(fontKey: string): Uint8Array {
  const hex = fontKey.replace(/[{}-]/g, '');
  if (!/^[0-9a-fA-F]{32}$/.test(hex)) throw new Error(`Invalid font key ${fontKey}`);
  const key = new Uint8Array(16);
  for (let i = 0; i < 16; i++) key[i] = Number.parseInt(hex.substr(30 - i * 2, 2), 16);
  return key;
}

/** XOR is its own inverse, so this both obfuscates and deobfuscates. */
export function deobfuscateFont(data: Uint8Array, fontKey: string): Uint8Array {
  const key = keyBytes(fontKey);
  const out = data.slice();
  for (let i = 0; i < 32 && i < out.length; i++) out[i]! ^= key[i % 16]!;
  return out;
}

export const obfuscateFont = deobfuscateFont;
