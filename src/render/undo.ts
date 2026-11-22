/**
 * One undo history for the viewer's editing: typing keeps a copy of the story's XML (cheap, one
 * part), while actions that can touch other parts (comments, lists, headers) keep a whole-document
 * checkpoint. No DOM here, so the history can be used and tested without a browser.
 */
import { EditError } from '../api/positions.js';
import type { DocumentCheckpoint, DocxDocument } from '../model/document.js';
import type { XmlElement } from '../xml/index.js';
import type { Caret, EditStory } from './editing.js';

export type UndoEntry =
  | { kind: 'story'; part: string; story: EditStory; root: XmlElement; container?: number; anchor: Caret; focus: Caret }
  | { kind: 'package'; checkpoint: DocumentCheckpoint; story: EditStory; container?: number; anchor: Caret; focus: Caret; label: string };

export class UndoHistory {
  /**
   * How long (ms) a transaction waits for the one ahead of it before it gives up with an
   * `EditError`. A change started after an `await` inside a transaction's own `run` waits for
   * that run, which may be waiting for it: this turns that hang into an error.
   */
  static queueTimeout = 10_000;

  private undoStack: UndoEntry[] = [];
  private redoStack: UndoEntry[] = [];
  private tail: Promise<void> = Promise.resolve();
  private pending = 0;
  /** A transaction's `run` is executing synchronously (before its first `await` takes effect). */
  private joinDepth = 0;

  constructor(private readonly limit = 200) {}

  /** Record a new step; anything that could be redone is dropped. */
  push(e: UndoEntry): void {
    this.undoStack.push(e);
    if (this.undoStack.length > this.limit) this.undoStack.shift();
    this.redoStack.length = 0;
  }

  /**
   * Step back. `current` builds the entry for the state being left (of the same kind as the entry
   * being restored); it goes on the redo stack. Returns the entry to restore.
   */
  undo(current: (restoring: UndoEntry) => UndoEntry): UndoEntry | undefined {
    if (this.pending) return undefined;
    const e = this.undoStack[this.undoStack.length - 1];
    if (!e) return undefined;
    this.redoStack.push(current(e));
    this.undoStack.pop();
    return e;
  }

  redo(current: (restoring: UndoEntry) => UndoEntry): UndoEntry | undefined {
    if (this.pending) return undefined;
    const e = this.redoStack[this.redoStack.length - 1];
    if (!e) return undefined;
    this.undoStack.push(current(e));
    this.redoStack.pop();
    return e;
  }

  clear(): void {
    this.undoStack.length = 0;
    this.redoStack.length = 0;
  }

  /** A transaction is running or waiting; undo and redo do nothing meanwhile. */
  get busy(): boolean {
    return this.pending > 0;
  }

  /**
   * A transaction's `run` is executing synchronously: a transaction started now is nested in it
   * and joins it. Later in the run (after an `await`) this is false, as there is no telling a
   * nested call from an unrelated one.
   */
  get joinable(): boolean {
    return this.joinDepth > 0;
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /**
   * Run a whole-document action as one undo step. A failure puts the document back as it was
   * and records nothing.
   *
   * Transactions run one at a time: one requested while another is running waits for it, so it
   * checkpoints a settled document and is an undo step of its own, and the other's failure does
   * not touch it. If the wait passes `UndoHistory.queueTimeout`, it gives up with an `EditError`
   * without running (nothing changed, nothing recorded).
   *
   * A transaction requested while another one's `run` is executing synchronously (a nested call
   * before the run's first `await`) joins it: it runs at once as part of that step, with no
   * checkpoint or entry of its own, and a failure goes to the outer `run` (which restores the
   * document if it lets the error through). Later in the run, changes go through the run's own
   * means (the editor it was given), since waiting for the outer step from inside it cannot end.
   */
  async transaction<T>(
    doc: DocxDocument,
    label: string,
    run: () => T | Promise<T>,
    where: { story: EditStory; container?: number; anchor: Caret; focus: Caret },
  ): Promise<T> {
    if (this.joinDepth) return await run();
    const before = this.tail;
    let release!: () => void;
    this.tail = new Promise<void>((r) => (release = r));
    this.pending++;
    let started = false;
    try {
      await waitFor(before, UndoHistory.queueTimeout);
      started = true;
      const checkpoint = doc.checkpoint();
      try {
        let running: T | Promise<T>;
        this.joinDepth++;
        try {
          running = run();
        } finally {
          this.joinDepth--;
        }
        const result = await running;
        this.push({ kind: 'package', checkpoint, story: where.story, ...(where.container !== undefined ? { container: where.container } : {}), anchor: { ...where.anchor }, focus: { ...where.focus }, label });
        return result;
      } catch (error) {
        doc.restore(checkpoint);
        throw error;
      }
    } finally {
      this.pending--;
      // One that gave up keeps its place: the next one still waits for the one ahead of it.
      if (started) release();
      else void before.then(release);
    }
  }
}

/** Wait for `turn`, or reject with an `EditError` after `ms`. */
async function waitFor(turn: Promise<void>, ms: number): Promise<void> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const gaveUp = new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new EditError('Another change is still running. Inside viewer.editing.transaction(), make changes through the editor it gives you.')), ms);
  });
  try {
    await Promise.race([turn, gaveUp]);
  } finally {
    clearTimeout(timer);
  }
}
