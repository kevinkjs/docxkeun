/** Part-name helpers. Part names here never have a leading slash: `word/document.xml`. */

export function normalizePartName(name: string): string {
  return name.replace(/\\/g, '/').replace(/^\/+/, '');
}

export function partKey(name: string): string {
  return normalizePartName(name).toLowerCase();
}

export function dirname(part: string): string {
  const i = part.lastIndexOf('/');
  return i < 0 ? '' : part.slice(0, i);
}

export function basename(part: string): string {
  const i = part.lastIndexOf('/');
  return i < 0 ? part : part.slice(i + 1);
}

export function extension(part: string): string {
  const b = basename(part);
  const i = b.lastIndexOf('.');
  return i < 0 ? '' : b.slice(i + 1).toLowerCase();
}

/** `_rels` part that holds relationships for a source part ('' = the package). */
export function relsPartFor(source: string): string {
  if (source === '') return '_rels/.rels';
  const dir = dirname(source);
  return `${dir ? dir + '/' : ''}_rels/${basename(source)}.rels`;
}

/** Resolve a relationship target against the directory of its source part. */
export function resolveTarget(source: string, target: string): string {
  const clean = target.split('#')[0]!;
  if (clean.startsWith('/')) return normalizePartName(clean);
  const segments = dirname(source).split('/').filter(Boolean);
  for (const seg of clean.split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') segments.pop();
    else segments.push(decodeURIComponentSafe(seg));
  }
  return segments.join('/');
}

/** Relative target from a source part to a destination part, as Word writes them. */
export function relativeTarget(source: string, destination: string): string {
  const from = dirname(source).split('/').filter(Boolean);
  const to = normalizePartName(destination).split('/');
  let common = 0;
  while (common < from.length && common < to.length - 1 && from[common] === to[common]) common++;
  const up = from.length - common;
  return [...Array<string>(up).fill('..'), ...to.slice(common)].join('/');
}

function decodeURIComponentSafe(s: string): string {
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
}
