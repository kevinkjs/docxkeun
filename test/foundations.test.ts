import { describe, expect, it } from 'vitest';
import { crc32 } from '../src/util/crc32.js';
import { base64Decode, base64Encode, utf8Encode, utf16leDecode, utf16leEncode } from '../src/util/text.js';
import { readZip, writeZip } from '../src/zip/zip.js';
import { parseXml, serializeXml, textContent, firstChild, XmlParseError } from '../src/xml/index.js';
import { OpcPackage, REL, resolveTarget, relativeTarget, relsPartFor } from '../src/opc/index.js';
import { buildDocx, p } from './helpers/docx.js';

describe('util', () => {
  it('crc32 matches the standard check value', () => {
    expect(crc32(utf8Encode('123456789'))).toBe(0xcbf43926);
    expect(crc32(new Uint8Array())).toBe(0);
  });

  it('base64 round-trips every length', () => {
    for (let len = 0; len < 20; len++) {
      const bytes = Uint8Array.from({ length: len }, (_, i) => (i * 37 + 11) & 0xff);
      const encoded = base64Encode(bytes);
      expect(encoded).toBe(Buffer.from(bytes).toString('base64'));
      expect([...base64Decode(encoded)]).toEqual([...bytes]);
    }
  });

  it('utf16le round-trips', () => {
    expect(utf16leDecode(utf16leEncode('Kata sandi ✓'))).toBe('Kata sandi ✓');
  });
});

describe('zip', () => {
  it('writes and reads stored and deflated entries', async () => {
    const big = utf8Encode('lorem ipsum '.repeat(500));
    const zip = await writeZip([
      { name: 'a.txt', data: utf8Encode('hello') },
      { name: 'dir/big.txt', data: big },
      { name: 'ünïcode.txt', data: utf8Encode('x') },
    ]);
    const entries = readZip(zip);
    expect(entries.map((e) => e.name)).toEqual(['a.txt', 'dir/big.txt', 'ünïcode.txt']);
    expect(entries[1]!.raw.method).toBe(8);
    expect(entries[1]!.raw.compressed.length).toBeLessThan(big.length);
    expect(new TextDecoder().decode(await entries[0]!.read())).toBe('hello');
    expect((await entries[1]!.read()).length).toBe(big.length);
  });

  it('re-writes raw entries byte-identically', async () => {
    const zip = await writeZip([{ name: 'x.xml', data: utf8Encode('<a>'.repeat(100)) }], { modified: new Date(2026, 0, 1) });
    const again = await writeZip(readZip(zip).map((e) => ({ name: e.name, raw: e.raw })), { modified: new Date(2026, 0, 1) });
    expect(Buffer.from(again).equals(Buffer.from(zip))).toBe(true);
  });

  it('detects corrupted data', async () => {
    const zip = await writeZip([{ name: 'x.txt', data: utf8Encode('abcdef'), compress: false }]);
    const corrupted = zip.slice();
    corrupted[30 + 5] = 0x5a; // first data byte after the 30-byte header and 5-byte name
    await expect(readZip(corrupted)[0]!.read()).rejects.toThrow(/CRC/);
  });

  it('rejects non-zip data', () => {
    expect(() => readZip(utf8Encode('not a zip file at all, definitely not'))).toThrow(/Not a zip/);
  });
});

describe('xml', () => {
  it('parses elements, attributes, entities, CDATA and comments', () => {
    const doc = parseXml('<?xml version="1.0"?><!-- c --><root a="1 &amp; 2" b=\'x\'><t>&lt;hi&gt; &#x41;&#66;</t><![CDATA[<raw>]]><e/></root>');
    expect(doc.declaration).toBe('version="1.0"');
    expect(doc.root.attrs).toEqual({ a: '1 & 2', b: 'x' });
    expect(textContent(firstChild(doc.root, 't')!)).toBe('<hi> AB');
    expect(doc.root.children[1]).toEqual({ type: 'cdata', value: '<raw>' });
  });

  it('round-trips through the serializer', () => {
    const src = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t xml:space="preserve"> a &amp; b &lt; c </w:t></w:r><w:r/></w:p>';
    const out = serializeXml(parseXml(src));
    expect(out).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + src);
  });

  it('reports mismatched tags', () => {
    expect(() => parseXml('<a><b></a>')).toThrow(XmlParseError);
  });

  it('does not expand doctype entities', () => {
    const doc = parseXml('<!DOCTYPE x [<!ENTITY e "boom">]><x>&e;</x>');
    expect(textContent(doc.root)).toBe('&e;');
  });

  it('canonicalises unusual prefixes and default namespaces', () => {
    const doc = parseXml(
      '<ns0:document xmlns:ns0="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14"><ns0:body><ns0:p x14:paraId="1" ns0:rsidR="00"/></ns0:body></ns0:document>',
    );
    expect(doc.root.name).toBe('w:document');
    expect(doc.root.attrs['xmlns:w']).toBeDefined();
    expect(doc.root.attrs['mc:Ignorable']).toBe('w14');
    const para = firstChild(firstChild(doc.root, 'w:body'), 'w:p')!;
    expect(para.attrs).toEqual({ 'w14:paraId': '1', 'w:rsidR': '00' });

    const def = parseXml('<document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><body/></document>');
    expect(def.root.name).toBe('w:document');
    expect(firstChild(def.root, 'w:body')).toBeDefined();
  });
});

describe('opc', () => {
  it('resolves relationship targets', () => {
    expect(resolveTarget('word/document.xml', 'media/image1.png')).toBe('word/media/image1.png');
    expect(resolveTarget('word/document.xml', '../customXml/item1.xml')).toBe('customXml/item1.xml');
    expect(resolveTarget('word/document.xml', '/word/styles.xml')).toBe('word/styles.xml');
    expect(resolveTarget('', 'word/document.xml')).toBe('word/document.xml');
    expect(relativeTarget('word/document.xml', 'word/media/a.png')).toBe('media/a.png');
    expect(relativeTarget('word/document.xml', 'customXml/item1.xml')).toBe('../customXml/item1.xml');
    expect(relsPartFor('word/document.xml')).toBe('word/_rels/document.xml.rels');
  });

  it('opens a package, follows relationships and saves untouched parts unchanged', async () => {
    const bytes = await buildDocx({ body: p('Hello'), styles: '<w:docDefaults/>' });
    const pkg = await OpcPackage.open(bytes);
    const main = await pkg.targetPart('', 'officeDocument');
    expect(main).toBe('word/document.xml');
    expect(await pkg.targetPart(main!, 'styles')).toBe('word/styles.xml');
    expect(pkg.contentType('word/document.xml')).toContain('document.main+xml');

    const saved = await pkg.save({ modified: new Date(2026, 0, 1) });
    const before = new Map(readZip(bytes).map((e) => [e.name, e.raw.compressed]));
    for (const e of readZip(saved)) {
      expect(Buffer.from(e.raw.compressed).equals(Buffer.from(before.get(e.name)!))).toBe(true);
    }
  });

  it('adds parts and relationships', async () => {
    const pkg = await OpcPackage.open(await buildDocx({ body: p('x') }));
    const id = await pkg.addRelationship('word/document.xml', REL.hyperlink, 'https://example.com/a?b=1&c=2', true);
    pkg.setBytes('word/media/pic.png', new Uint8Array([1, 2, 3]));
    const reopened = await OpcPackage.open(await pkg.save());
    const rel = await reopened.relationshipById('word/document.xml', id);
    expect(rel).toMatchObject({ target: 'https://example.com/a?b=1&c=2', external: true });
    expect(reopened.contentType('word/media/pic.png')).toBe('image/png');
    expect([...(await reopened.getBytes('word/media/pic.png'))]).toEqual([1, 2, 3]);
  });
});
