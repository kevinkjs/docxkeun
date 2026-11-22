import { describe, expect, it } from 'vitest';
import { positionMenu } from '../src/render/ui/menu.js';

const vp = { w: 800, h: 600 };
describe('positionMenu', () => {
  it('leaves a menu that fits unchanged', () => {
    expect(positionMenu({ x: 100, y: 100 }, { w: 200, h: 150 }, vp)).toEqual({ x: 100, y: 100 });
  });
  it('flips left near the right edge', () => {
    expect(positionMenu({ x: 700, y: 100 }, { w: 200, h: 150 }, vp)).toEqual({ x: 500, y: 100 });
  });
  it('flips up near the bottom', () => {
    expect(positionMenu({ x: 100, y: 550 }, { w: 200, h: 150 }, vp)).toEqual({ x: 100, y: 400 });
  });
  it('clamps a menu larger than the viewport to the margin', () => {
    expect(positionMenu({ x: 300, y: 300 }, { w: 900, h: 700 }, vp)).toEqual({ x: 4, y: 4 });
  });
});
