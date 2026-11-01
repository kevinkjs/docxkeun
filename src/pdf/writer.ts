/** Collects indirect objects and writes a complete PDF file with a cross-reference table. */
import { ByteWriter } from '../util/bytes.js';
import { deflate } from '../zip/deflate.js';
import { name, PdfRef, PdfStream, str, writeValue, type ObjectCrypt, type PdfDict, type PdfValue } from './objects.js';

/** Applies a security handler while writing (see pdf/security.ts). */
export interface PdfEncryption {
  /** The /Encrypt dictionary (written unencrypted). */
  dict: PdfDict;
  /** Per-object crypt for strings and streams. */
  forObject(num: number, gen: number): ObjectCrypt;
  /** Minimum PDF version / extension the handler needs. */
  extensions?: PdfDict;
}

export class PdfWriter {
  private readonly objects: Array<PdfValue | undefined> = [undefined];

  alloc(): PdfRef {
    this.objects.push(undefined);
    return new PdfRef(this.objects.length - 1);
  }

  set(ref: PdfRef, value: PdfValue): PdfRef {
    this.objects[ref.num] = value;
    return ref;
  }

  add(value: PdfValue): PdfRef {
    return this.set(this.alloc(), value);
  }

  /** A Flate-compressed stream object. */
  async stream(dict: PdfDict, data: Uint8Array, compress = true): Promise<PdfRef> {
    if (!compress) return this.add(new PdfStream(dict, data));
    const packed = await deflate(data, 'deflate');
    return this.add(new PdfStream({ ...dict, Filter: name('FlateDecode') }, packed));
  }

  build(root: PdfRef, info: PdfRef | undefined, fileId: Uint8Array, encryption?: PdfEncryption): Uint8Array {
    const w = new ByteWriter(1 << 20);
    w.ascii('%PDF-1.7\n%\xe2\xe3\xcf\xd3\n');
    const offsets: number[] = [0];
    let encryptRef: PdfRef | undefined;
    if (encryption) encryptRef = this.add(encryption.dict);
    for (let i = 1; i < this.objects.length; i++) {
      const value = this.objects[i];
      offsets[i] = w.length;
      w.ascii(`${i} 0 obj\n`);
      const crypt = encryption && i !== encryptRef?.num ? encryption.forObject(i, 0) : undefined;
      writeValue(w, value === undefined ? null : value, crypt);
      w.ascii('\nendobj\n');
    }
    const xref = w.length;
    w.ascii(`xref\n0 ${this.objects.length}\n`);
    w.ascii('0000000000 65535 f\r\n');
    for (let i = 1; i < this.objects.length; i++) w.ascii(`${String(offsets[i]).padStart(10, '0')} 00000 n\r\n`);
    // The trailer (with the file identifier) is never encrypted.
    const trailer: PdfDict = {
      Size: this.objects.length,
      Root: root,
      Info: info,
      ID: [str(fileId, true), str(fileId, true)],
      Encrypt: encryptRef,
    };
    w.ascii('trailer\n');
    writeValue(w, trailer);
    w.ascii(`\nstartxref\n${xref}\n%%EOF\n`);
    return w.toBytes();
  }
}
