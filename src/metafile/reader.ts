/**
 * Little-endian field access for one metafile record. Reads outside the record return 0 (and
 * byte slices come back short) instead of throwing, because metafiles in the wild are often
 * truncated or carry wrong counts, and a damaged record should not lose the rest of the picture.
 */
export class RecordReader {
  /**
   * `start` and `end` are absolute offsets in `bytes`; field offsets are relative to `start`.
   * Pass `view` (over the same bytes) to share one DataView across a file's many records.
   */
  constructor(
    readonly bytes: Uint8Array,
    readonly start: number,
    readonly end: number,
    private readonly view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength),
  ) {}

  get size(): number {
    return this.end - this.start;
  }

  private ok(offset: number, length: number): boolean {
    return offset >= 0 && this.start + offset + length <= this.end;
  }

  u8(offset: number): number {
    return this.ok(offset, 1) ? this.view.getUint8(this.start + offset) : 0;
  }

  u16(offset: number): number {
    return this.ok(offset, 2) ? this.view.getUint16(this.start + offset, true) : 0;
  }

  i16(offset: number): number {
    return this.ok(offset, 2) ? this.view.getInt16(this.start + offset, true) : 0;
  }

  u32(offset: number): number {
    return this.ok(offset, 4) ? this.view.getUint32(this.start + offset, true) : 0;
  }

  i32(offset: number): number {
    return this.ok(offset, 4) ? this.view.getInt32(this.start + offset, true) : 0;
  }

  f32(offset: number): number {
    if (!this.ok(offset, 4)) return 0;
    const v = this.view.getFloat32(this.start + offset, true);
    return Number.isFinite(v) ? v : 0;
  }

  /** Up to `length` bytes from `offset`, clipped to the record. */
  slice(offset: number, length: number): Uint8Array {
    const from = Math.max(this.start, Math.min(this.end, this.start + offset));
    const to = Math.max(from, Math.min(this.end, from + Math.max(0, length)));
    return this.bytes.subarray(from, to);
  }

  /** How many items of `itemSize` bytes fit between `offset` and the end of the record (caps bogus counts). */
  fit(offset: number, count: number, itemSize: number): number {
    const room = Math.floor((this.size - offset) / itemSize);
    return Math.max(0, Math.min(count, room));
  }
}
