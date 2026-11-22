import { describe, expect, it } from 'vitest';
import { fitGroups } from '../src/render/ui/overflow.js';

describe('fitGroups', () => {
  it('keeps everything when it all fits', () => expect(fitGroups([40, 60, 80, 50], 300, 28)).toBe(4));
  it('reserves room for the more button', () => expect(fitGroups([40, 60, 80, 50], 200, 28)).toBe(2));
  it('handles no groups', () => expect(fitGroups([], 100, 28)).toBe(0));
  it('moves a group that never fits', () => expect(fitGroups([500], 100, 28)).toBe(0));
});
