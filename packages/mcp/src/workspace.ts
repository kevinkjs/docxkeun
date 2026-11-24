/**
 * The folders the server may touch. Every path an AI client sends is resolved here, and anything
 * outside the allowed folders (including through symlinks) is refused.
 */
import { randomUUID } from 'node:crypto';
import { readdir, readFile, realpath, rename, stat, unlink, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { basename, dirname, extname, isAbsolute, join, relative, resolve, sep } from 'node:path';

export const DOCUMENT_EXTENSIONS = ['.docx', '.docm', '.dotx', '.dotm'];

export type FileKind = 'document' | 'pdf' | 'image';

const IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff', '.emf', '.wmf'];

export class WorkspaceError extends Error {
  override name = 'WorkspaceError';
}

export interface DocumentEntry {
  path: string;
  /** Path relative to the allowed folder it is in. */
  relative: string;
  size: number;
  modified: Date;
}

export class Workspace {
  private constructor(readonly roots: string[]) {}

  /** Folders must exist; they are stored as real paths so symlinks cannot be used to escape them. */
  static async create(folders: string[]): Promise<Workspace> {
    if (folders.length === 0) throw new WorkspaceError('No folders given');
    const roots: string[] = [];
    for (const folder of folders) {
      const abs = resolve(expandHome(folder));
      let real: string;
      try {
        real = await realpath(abs);
      } catch {
        throw new WorkspaceError(`Folder does not exist: ${abs}`);
      }
      if (!(await stat(real)).isDirectory()) throw new WorkspaceError(`Not a folder: ${abs}`);
      if (real === resolve('/')) throw new WorkspaceError('Refusing to allow the whole file system; pass the folders that hold your documents');
      roots.push(real);
    }
    return new Workspace([...new Set(roots)]);
  }

  /**
   * Resolve a path from the client: absolute, `~/…`, or relative to the first allowed folder.
   * With `mustExist` the file has to exist; otherwise its folder has to.
   */
  async resolve(path: string, kind: FileKind, mustExist: boolean): Promise<string> {
    if (!path.trim()) throw new WorkspaceError('Empty path');
    const expanded = expandHome(path.trim());
    const abs = isAbsolute(expanded) ? resolve(expanded) : resolve(this.roots[0]!, expanded);
    const ext = extname(abs).toLowerCase();
    if (kind === 'document' && !DOCUMENT_EXTENSIONS.includes(ext)) {
      throw new WorkspaceError(`Not a Word document (${DOCUMENT_EXTENSIONS.join(', ')}): ${abs}`);
    }
    if (kind === 'pdf' && ext !== '.pdf') throw new WorkspaceError(`Not a .pdf path: ${abs}`);
    if (kind === 'image' && !IMAGE_EXTENSIONS.includes(ext)) throw new WorkspaceError(`Not a picture (${IMAGE_EXTENSIONS.join(', ')}): ${abs}`);

    let real: string;
    try {
      real = await realpath(abs);
    } catch {
      if (mustExist) throw new WorkspaceError(`File not found: ${abs}`);
      let dir: string;
      try {
        dir = await realpath(dirname(abs));
      } catch {
        throw new WorkspaceError(`Folder does not exist: ${dirname(abs)}`);
      }
      real = join(dir, basename(abs));
    }
    if (!this.contains(real)) {
      throw new WorkspaceError(`Outside the folders this server may use (${this.roots.join(', ')}): ${abs}`);
    }
    return real;
  }

  contains(real: string): boolean {
    return this.roots.some((root) => {
      const rel = relative(root, real);
      return rel !== '' && !rel.startsWith('..') && !isAbsolute(rel);
    });
  }

  /** A path for messages: relative to its folder when there is a single folder, absolute otherwise. */
  display(real: string): string {
    if (this.roots.length === 1) {
      const rel = relative(this.roots[0]!, real);
      if (!rel.startsWith('..') && !isAbsolute(rel)) return rel;
    }
    return real;
  }

  async read(real: string): Promise<Uint8Array> {
    return new Uint8Array(await readFile(real));
  }

  async exists(real: string): Promise<boolean> {
    try {
      await stat(real);
      return true;
    } catch {
      return false;
    }
  }

  /** Write through a temporary file and a rename, so a failed write never leaves a half-written document. */
  async write(real: string, bytes: Uint8Array): Promise<void> {
    if (!this.contains(real)) throw new WorkspaceError(`Outside the allowed folders: ${real}`);
    let mode: number | undefined;
    try {
      mode = (await stat(real)).mode & 0o777;
    } catch {
      mode = undefined;
    }
    const tmp = join(dirname(real), `.${basename(real)}.${randomUUID().slice(0, 8)}.tmp`);
    try {
      await writeFile(tmp, bytes, mode === undefined ? undefined : { mode });
      await rename(tmp, real);
    } catch (e) {
      await unlink(tmp).catch(() => {});
      throw e;
    }
  }

  /** Word documents in the allowed folders (or in `folder`, which must be inside them). */
  async listDocuments(options: { folder?: string; query?: string; limit?: number } = {}): Promise<{ entries: DocumentEntry[]; truncated: boolean }> {
    const limit = options.limit ?? 200;
    let starts = this.roots;
    if (options.folder) {
      const expanded = expandHome(options.folder);
      const abs = isAbsolute(expanded) ? resolve(expanded) : resolve(this.roots[0]!, expanded);
      let real: string;
      try {
        real = await realpath(abs);
      } catch {
        throw new WorkspaceError(`Folder does not exist: ${abs}`);
      }
      if (!this.roots.includes(real) && !this.contains(real)) throw new WorkspaceError(`Outside the allowed folders: ${abs}`);
      starts = [real];
    }
    const query = options.query?.toLowerCase();
    const entries: DocumentEntry[] = [];
    let truncated = false;

    const walk = async (dir: string, root: string, depth: number): Promise<void> => {
      if (truncated || depth > 8) return;
      let items;
      try {
        items = await readdir(dir, { withFileTypes: true });
      } catch {
        return;
      }
      items.sort((a, b) => a.name.localeCompare(b.name));
      for (const item of items) {
        if (truncated) return;
        if (item.name.startsWith('.') || item.name === 'node_modules') continue;
        const full = join(dir, item.name);
        if (item.isDirectory()) {
          await walk(full, root, depth + 1);
        } else if (item.isFile() && DOCUMENT_EXTENSIONS.includes(extname(item.name).toLowerCase())) {
          // ~$name.docx is the lock file Word keeps next to an open document.
          if (item.name.startsWith('~$')) continue;
          if (query && !item.name.toLowerCase().includes(query)) continue;
          if (entries.length >= limit) {
            truncated = true;
            return;
          }
          const info = await stat(full);
          entries.push({ path: full, relative: relative(root, full), size: info.size, modified: info.mtime });
        }
      }
    };
    for (const start of starts) {
      const root = this.roots.find((r) => start === r || start.startsWith(r + sep)) ?? start;
      await walk(start, root, 0);
    }
    return { entries, truncated };
  }
}

export function expandHome(path: string): string {
  if (path === '~') return homedir();
  if (path.startsWith('~/') || path.startsWith('~\\')) return join(homedir(), path.slice(2));
  return path;
}

/** Runs tasks one at a time (every change to a file goes through here). */
export class Queue {
  private tail: Promise<unknown> = Promise.resolve();

  /** Resolves once every task queued so far has finished. */
  idle(): Promise<void> {
    return this.tail.then(() => {});
  }

  run<T>(task: () => Promise<T>): Promise<T> {
    const next = this.tail.then(task, task);
    this.tail = next.catch(() => {});
    return next;
  }
}
