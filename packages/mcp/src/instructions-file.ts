import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { MAX_INSTRUCTIONS } from './instructions.js';
import { expandHome } from './workspace.js';

/** Reads the user's own instructions from a text or Markdown file; the text may be empty. */
export async function readInstructionsFile(path: string, cwd: string = process.cwd()): Promise<string> {
  const abs = resolve(cwd, expandHome(path));
  let raw: string;
  try {
    raw = await readFile(abs, 'utf8');
  } catch (e) {
    throw new Error(`Cannot read the instructions file ${abs}: ${(e as NodeJS.ErrnoException).code ?? (e as Error).message}`);
  }
  const text = raw.replace(/^﻿/, '').trim();
  if (text.length > MAX_INSTRUCTIONS) {
    throw new Error(`The instructions file ${abs} has ${text.length} characters; the limit is ${MAX_INSTRUCTIONS}.`);
  }
  return text;
}
