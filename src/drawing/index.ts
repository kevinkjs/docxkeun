/**
 * DrawingML shape geometry: preset (`a:prstGeom`) and custom (`a:custGeom`) shapes turned into
 * vector paths, plus line-end decorations. Layout positions and paints the results.
 */
export { angleToRadians, builtinGuides, evaluateFormula, evaluateGuides, radiansToAngle, resolveValue } from './formula.js';
export type { GuideDef, Guides } from './formula.js';
export { PRESET_NAMES, presetDefinition } from './presets.js';
export type { PathCommandDef, PathDef, PathFill, PathPoint, ShapeDefinition, TextRectDef } from './presets.js';
export { arcToBeziers, customGeometry, evaluateDefinition, parseCustomGeometry, presetGeometry, readAdjustments, shadeFill } from './geometry.js';
export type { ShapeGeometry, ShapePath } from './geometry.js';
export { lineEnd, readLineEnd } from './arrows.js';
export type { LineEndShape, LineEndSize, LineEndSpec, LineEndType } from './arrows.js';
