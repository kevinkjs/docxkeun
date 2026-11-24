import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { MAX_INSTRUCTIONS, instructionsSection } from '../src/instructions.js';
import { readInstructionsFile } from '../src/instructions-file.js';

describe('instructionsSection', () => {
  it('wraps the trimmed text after a line saying whose it is', () => {
    const s = instructionsSection('  You are a legal assistant.\n');
    expect(s.startsWith("The user's own instructions for this work follow.")).toBe(true);
    expect(s.endsWith('<user_instructions>\nYou are a legal assistant.\n</user_instructions>')).toBe(true);
  });
  it('is empty for no text or only whitespace', () => {
    for (const t of [undefined, '', '   ', '\n\t\n']) expect(instructionsSection(t)).toBe('');
  });
  it('allows 4000 characters', () => expect(MAX_INSTRUCTIONS).toBe(4000));
});

describe('readInstructionsFile', () => {
  let dir: string;
  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'docxkeun-instructions-'));
  });
  afterEach(() => rm(dir, { recursive: true, force: true }));

  it('strips a BOM and surrounding whitespace', async () => {
    const file = join(dir, 'a.md');
    await writeFile(file, '\uFEFF\n  Act as a proofreader.\r\n\n');
    expect(await readInstructionsFile(file)).toBe('Act as a proofreader.');
  });
  it('resolves a relative path against the working directory', async () => {
    await writeFile(join(dir, 'notes.md'), 'Be brief.');
    expect(await readInstructionsFile('notes.md', dir)).toBe('Be brief.');
  });
  it('returns an empty string for an empty file', async () => {
    await writeFile(join(dir, 'empty.md'), '');
    expect(await readInstructionsFile(join(dir, 'empty.md'))).toBe('');
  });
  it('accepts 4000 characters and refuses 4001', async () => {
    await writeFile(join(dir, 'ok.md'), 'x'.repeat(MAX_INSTRUCTIONS));
    await writeFile(join(dir, 'long.md'), 'x'.repeat(MAX_INSTRUCTIONS + 1));
    expect(await readInstructionsFile(join(dir, 'ok.md'))).toHaveLength(MAX_INSTRUCTIONS);
    await expect(readInstructionsFile(join(dir, 'long.md'))).rejects.toThrow(/has 4001 characters; the limit is 4000\./);
  });
  it('names the path when the file is missing or is a folder', async () => {
    await expect(readInstructionsFile(join(dir, 'missing.md'))).rejects.toThrow(/Cannot read the instructions file .*missing\.md/);
    await expect(readInstructionsFile(dir)).rejects.toThrow(/Cannot read the instructions file/);
  });
});
