import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

describe('package.json', () => {
  it('has no runtime dependencies', () => {
    const pkg = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8'));
    expect(Object.keys(pkg.dependencies ?? {})).toEqual([]);
    expect(Object.keys(pkg.peerDependencies ?? {})).toEqual([]);
    expect(Object.keys(pkg.optionalDependencies ?? {})).toEqual([]);
  });
});
