/**
 * Compression through the platform's Compression Streams API.
 * 'deflate-raw' is what ZIP stores; 'deflate' (zlib-wrapped) is what PDF FlateDecode and PNG use.
 */

export type CompressionFormatName = 'deflate-raw' | 'deflate';

async function pump(data: Uint8Array, transform: TransformStream<Uint8Array, Uint8Array>): Promise<Uint8Array> {
  const writer = transform.writable.getWriter();
  const reader = transform.readable.getReader();
  // Write and read concurrently, otherwise back-pressure can stall on large inputs.
  const writing = writer.write(data).then(() => writer.close());
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    total += value.length;
  }
  await writing;
  if (chunks.length === 1) return chunks[0]!;
  const out = new Uint8Array(total);
  let offset = 0;
  for (const c of chunks) {
    out.set(c, offset);
    offset += c.length;
  }
  return out;
}

function requireStreams(): void {
  if (typeof DecompressionStream === 'undefined' || typeof CompressionStream === 'undefined') {
    throw new Error('docxkit needs the Compression Streams API (modern browsers, Node >= 20.12).');
  }
}

export async function inflate(data: Uint8Array, format: CompressionFormatName = 'deflate-raw'): Promise<Uint8Array> {
  requireStreams();
  return pump(data, new DecompressionStream(format) as unknown as TransformStream<Uint8Array, Uint8Array>);
}

export async function deflate(data: Uint8Array, format: CompressionFormatName = 'deflate-raw'): Promise<Uint8Array> {
  requireStreams();
  return pump(data, new CompressionStream(format) as unknown as TransformStream<Uint8Array, Uint8Array>);
}
