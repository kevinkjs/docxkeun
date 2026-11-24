import { describe, expect, it } from 'vitest';
import { MAX_INSTRUCTIONS, instructionsSection } from '../src/instructions.js';

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
