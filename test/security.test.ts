import { describe, expect, it } from 'vitest';
import { DocxEditor } from '../src/api/index.js';
import { DocxDocument } from '../src/model/index.js';
import {
  AccessPolicy,
  checkProtectionPassword,
  legacyHash,
  legacyVerifier,
  PermissionError,
  readProtection,
  removeProtection,
  setProtection,
  setWriteProtection,
  checkWriteProtectionPassword,
} from '../src/security/index.js';
import { attr, findAll } from '../src/xml/index.js';
import { buildDocx, p } from './helpers/docx.js';

describe('legacy Word/Excel password hash', () => {
  it('matches the known Excel verifier for "test"', () => {
    expect(legacyVerifier('test').toString(16).toUpperCase()).toBe('CBEB');
  });

  it('agrees with the flat-matrix formulation of MS-OFFCRYPTO 2.3.7.2', () => {
    // Re-derive the key the way the spec writes it: walk the password backwards over a flat matrix.
    const seeds = [0xaefc, 0x7b61, 0x4563, 0x0375, 0xd849, 0x6f45, 0xeb23, 0x47d3, 0xb861, 0x45a0, 0xaa51, 0x76b4, 0x3730, 0x3331, 0x1021];
    const flat: number[] = [];
    for (const s of seeds) {
      let v = s;
      for (let j = 0; j < 7; j++) {
        flat.push(v);
        v <<= 1;
        if (v & 0x10000) v = (v & 0xffff) ^ 0x1021;
      }
    }
    const initial = [0xe1f0, 0x1d0f, 0xcc9c, 0x84c0, 0x110c, 0x0e10, 0xf1ce, 0x313e, 0x1872, 0xe139, 0xd40f, 0x84f9, 0x280c, 0xa96a, 0x4ec3];
    const key = (pw: string) => {
      const bytes = [...pw].map((c) => c.charCodeAt(0) & 0xff);
      let k = initial[bytes.length - 1]!;
      let element = 0x68;
      for (let i = bytes.length - 1; i >= 0; i--) {
        let ch = bytes[i]!;
        for (let b = 0; b < 7; b++) {
          if (ch & 0x40) k ^= flat[element]!;
          ch <<= 1;
          element--;
        }
      }
      return k;
    };
    expect(flat[104]).toBe(0x48c4);
    for (const pw of ['a', 'test', 'Password1', 'abcdefghijklmno']) expect(legacyHash(pw) >>> 16).toBe(key(pw));
  });
});

describe('document protection', () => {
  it('sets, verifies and removes password protection', async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: p('x'), settings: '<w:defaultTabStop w:val="720"/>' }));
    await setProtection(doc, 'readOnly', { password: 'secret', spinCount: 1000 });
    const reopened = await DocxDocument.open(await doc.save());
    expect(readProtection(reopened)).toEqual({ mode: 'readOnly', enforced: true, formatting: false, hasPassword: true });
    const node = findAll(reopened.xmlSync(reopened.parts.settings)!.root, 'w:documentProtection')[0]!;
    expect(attr(node, 'w:cryptAlgorithmSid')).toBe('14');
    expect(attr(node, 'w:salt')).toMatch(/^[A-Za-z0-9+/]{22}==$/);
    // It sits before w:defaultTabStop, as the schema requires.
    const names = reopened.xmlSync(reopened.parts.settings)!.root.children.filter((c) => c.type === 'element').map((c) => (c as { name: string }).name);
    expect(names.indexOf('w:documentProtection')).toBeLessThan(names.indexOf('w:defaultTabStop'));
    expect(checkProtectionPassword(reopened, 'secret')).toBe(true);
    expect(checkProtectionPassword(reopened, 'Secret')).toBe(false);
    await expect(removeProtection(reopened, 'wrong')).rejects.toThrow(/Incorrect/);
    await removeProtection(reopened, 'secret');
    expect(readProtection(reopened).mode).toBe('none');
  });

  it('verifies the ISO attribute form', async () => {
    const { hash } = await import('../src/crypto/index.js');
    const { utf16leEncode, base64Encode } = await import('../src/util/text.js');
    // spinCount 0: hash = SHA-256(salt || UTF-16LE(password))
    const salt = new Uint8Array(16).fill(7);
    const pw = utf16leEncode('pässword');
    const buf = new Uint8Array(salt.length + pw.length);
    buf.set(salt);
    buf.set(pw, salt.length);
    const settings = `<w:documentProtection w:edit="comments" w:enforcement="1" w:algorithmName="SHA-256" w:hashValue="${base64Encode(hash('SHA-256', buf))}" w:saltValue="${base64Encode(salt)}" w:spinCount="0"/>`;
    const doc = await DocxDocument.open(await buildDocx({ body: p('x'), settings }));
    expect(readProtection(doc).mode).toBe('comments');
    expect(checkProtectionPassword(doc, 'pässword')).toBe(true);
    expect(checkProtectionPassword(doc, 'password')).toBe(false);
  });

  it('supports password to modify', async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: p('x'), settings: '' }));
    await setWriteProtection(doc, { password: 'modify', recommended: true });
    expect(checkWriteProtectionPassword(doc, 'modify')).toBe(true);
    expect(checkWriteProtectionPassword(doc, 'nope')).toBe(false);
  });
});

describe('access policy', () => {
  const body = p('Fixed terms of the agreement.') + '<w:p><w:r><w:t xml:space="preserve">Notes: </w:t></w:r><w:permStart w:id="1" w:ed="budi@example.com"/><w:r><w:t>budi writes here</w:t></w:r><w:permEnd w:id="1"/><w:permStart w:id="2" w:edGrp="everyone"/><w:r><w:t xml:space="preserve"> anyone writes here</w:t></w:r><w:permEnd w:id="2"/></w:p>';

  async function setup(role: 'viewer' | 'commenter' | 'editor' | 'owner', user = { id: 'budi@example.com', name: 'Budi' }, mode: 'readOnly' | 'comments' | 'forms' | 'none' = 'readOnly') {
    const doc = await DocxDocument.open(await buildDocx({ body, settings: '' }));
    if (mode !== 'none') await setProtection(doc, mode, { spinCount: 10 });
    const policy = new AccessPolicy(doc, { user, role });
    return { doc, policy, editor: new DocxEditor(doc, policy) };
  }

  it('lets viewers do nothing but read and export', async () => {
    const { editor, policy } = await setup('viewer');
    expect(policy.can('export')).toBe(true);
    expect(() => editor.insertText(editor.find('Fixed')[0]!.start, 'x')).toThrow(PermissionError);
    await expect(editor.addComment(editor.find('Fixed')[0]!, { author: 'Budi', text: 'hi' })).rejects.toThrow(/Viewers/);
  });

  it('enforces read-only protection with per-user and group editable ranges', async () => {
    const { editor } = await setup('editor');
    expect(() => editor.replaceRange(editor.find('Fixed')[0]!, 'Changed')).toThrow(/read-only/);
    editor.replaceRange(editor.find('budi writes')[0]!, 'Budi wrote');
    editor.replaceRange(editor.find('anyone writes')[0]!, 'everyone wrote');
    const other = await setup('editor', { id: 'rina@example.com', name: 'Rina' });
    expect(() => other.editor.replaceRange(other.editor.find('budi writes')[0]!, 'nope')).toThrow(/read-only/);
    other.editor.replaceRange(other.editor.find('anyone writes')[0]!, 'ok');
    // A range that starts inside an editable region but leaves it is refused.
    const m = other.editor.find('ok')[0]!;
    expect(() => other.editor.deleteRange({ start: other.editor.find('Fixed')[0]!.start, end: m.end })).toThrow(PermissionError);
  });

  it('allows comments in comments-only mode, but not edits', async () => {
    const { editor } = await setup('commenter', { id: 'rina@example.com', name: 'Rina' }, 'comments');
    const id = await editor.addComment(editor.find('Fixed terms')[0]!, { author: 'Rina', text: 'Why fixed?' });
    expect(() => editor.formatRange(editor.find('Fixed')[0]!, { bold: true })).toThrow(/Commenters/);
    await editor.editComment(id, 'Why are these fixed?');
    const budi = await setup('editor', { id: 'budi@example.com', name: 'Budi' }, 'comments');
    const id2 = await budi.editor.addComment(budi.editor.find('Fixed')[0]!, { author: 'Rina', text: 'by Rina' });
    void id2;
    const rina = new DocxEditor(budi.doc, new AccessPolicy(budi.doc, { user: { id: 'x', name: 'Someone else' }, role: 'commenter' }));
    await expect(rina.editComment(id2, 'hijack')).rejects.toThrow(/own comments/);
  });

  it('lets owners bypass protection and manage editable ranges', async () => {
    const { editor, doc } = await setup('owner');
    editor.replaceRange(editor.find('Fixed')[0]!, 'Owner-edited');
    const id = editor.addEditableRange(editor.find('terms')[0]!, { group: 'editors' });
    expect(editor.editableRanges().map((r) => r.id)).toContain(id);
    const editorUser = new DocxEditor(doc, new AccessPolicy(doc, { user: { id: 'z' }, role: 'editor' }));
    editorUser.replaceRange(editorUser.find('terms')[0]!, 'clauses');
    expect(() => editorUser.addEditableRange(editorUser.find('agreement')[0]!, { group: 'everyone' })).toThrow(/owners/);
  });

  it('only allows form filling under forms protection', async () => {
    const doc = await DocxDocument.open(
      await buildDocx({ body: '<w:sdt><w:sdtPr><w:tag w:val="name"/></w:sdtPr><w:sdtContent><w:r><w:t>Name</w:t></w:r></w:sdtContent></w:sdt><w:p><w:r><w:t>Static</w:t></w:r></w:p>', settings: '' }),
    );
    await setProtection(doc, 'forms', { spinCount: 10 });
    const editor = new DocxEditor(doc, new AccessPolicy(doc, { user: { id: 'a' }, role: 'editor' }));
    expect(editor.fillContentControls({ name: 'Rina' })).toEqual(['name']);
    expect(() => editor.replaceRange(editor.find('Static')[0]!, 'x')).toThrow(/forms/);
  });
});
