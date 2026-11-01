/**
 * AccessPolicy: who may do what. One decision function serves both the UI (viewer buttons,
 * highlighted editable regions) and the edit API (every DocxEditor call is checked).
 *
 * It combines the app's role for the current user with the document's own protection
 * (read-only / comments-only / tracked changes / forms) and its editable ranges.
 * Client-side checks stop honest mistakes; servers must re-check what they accept.
 */
import type { DocRange, EditAction, EditGuard } from '../api/editor.js';
import { TreeIndex } from '../api/positions.js';
import type { PermArea } from '../layout/types.js';
import type { DocxDocument } from '../model/document.js';
import { attr, findAll, type XmlElement } from '../xml/index.js';
import { readProtection } from './protection.js';

export type Role = 'viewer' | 'commenter' | 'editor' | 'owner';

export interface AccessUser {
  id: string;
  name?: string;
  email?: string;
  /** Group names, matched against w:edGrp (e.g. "editors", "administrators"). */
  groups?: string[];
}

export type Capability = EditAction | 'fillForm' | 'editComment' | 'read' | 'export' | 'print' | 'copy';

export interface AccessPolicyOptions {
  user: AccessUser;
  role: Role;
  /** Honour the document's own protection settings (default true). Owners always bypass it. */
  enforceProtection?: boolean;
  /** Roles allowed to export, print and copy (default: every role). */
  outputRoles?: Partial<Record<'export' | 'print' | 'copy', Role[]>>;
}

export class PermissionError extends Error {
  override name = 'PermissionError';
  constructor(readonly action: Capability, message: string) {
    super(message);
  }
}

const COMMENT_ACTIONS: Capability[] = ['comment', 'reply', 'resolveComment', 'editComment', 'deleteComment'];
const CONTENT_ACTIONS: Capability[] = ['edit', 'format', 'link', 'bookmark', 'fillForm'];

interface PermRange {
  id: string;
  editor?: string;
  group?: string;
  start: XmlElement;
  end: XmlElement;
}

export class AccessPolicy implements EditGuard {
  constructor(readonly doc: DocxDocument, readonly options: AccessPolicyOptions) {}

  get user(): AccessUser {
    return this.options.user;
  }

  get role(): Role {
    return this.options.role;
  }

  /** The reason an action is refused, or undefined when it is allowed. */
  denial(action: Capability, range?: DocRange, detail: { commentAuthor?: string } = {}): string | undefined {
    const { role } = this.options;
    if (action === 'read') return undefined;
    if (action === 'export' || action === 'print' || action === 'copy') {
      const roles = this.options.outputRoles?.[action];
      return roles && !roles.includes(role) ? `Your role (${role}) may not ${action} this document` : undefined;
    }
    if (role === 'owner') return undefined;
    if (action === 'protect') return 'Only owners can change protection and editable ranges';
    if (role === 'viewer') return 'Viewers cannot change this document';
    const isCommentAction = COMMENT_ACTIONS.includes(action);
    if (role === 'commenter' && !isCommentAction) return 'Commenters can only add and manage comments';
    if ((action === 'editComment' || (action === 'deleteComment' && role === 'commenter')) && detail.commentAuthor !== undefined && !this.isCurrentUser(detail.commentAuthor)) {
      return 'You can only change your own comments';
    }

    if (this.options.enforceProtection === false) return undefined;
    const protection = readProtection(this.doc);
    if (!protection.enforced) return undefined;
    const inEditable = range ? this.inEditableRange(range) : false;
    switch (protection.mode) {
      case 'readOnly':
        if (inEditable) return undefined;
        return 'The document is read-only outside its editable regions';
      case 'comments':
        if (isCommentAction || inEditable) return undefined;
        return 'The document only allows comments outside its editable regions';
      case 'trackedChanges':
        if (isCommentAction) return undefined;
        return 'The document requires tracked changes, which docxkit does not record yet';
      case 'forms':
        if (action === 'fillForm') return undefined;
        return 'The document only allows filling in forms';
      default:
        return undefined;
    }
  }

  can(action: Capability, range?: DocRange, detail?: { commentAuthor?: string }): boolean {
    return this.denial(action, range, detail) === undefined;
  }

  assert(action: Capability, range?: DocRange, detail?: { commentAuthor?: string }): void {
    const reason = this.denial(action, range, detail);
    if (reason) throw new PermissionError(action, reason);
  }

  private isCurrentUser(who: string): boolean {
    const w = who.trim().toLowerCase();
    const { id, name, email } = this.options.user;
    const candidates = [id, name, email].filter(Boolean).map((s) => s!.toLowerCase());
    if (candidates.includes(w)) return true;
    // DOMAIN\user
    const slash = w.lastIndexOf('\\');
    return slash >= 0 && candidates.includes(w.slice(slash + 1));
  }

  private matchesGroup(group: string): boolean {
    const g = group.toLowerCase();
    const groups = (this.options.user.groups ?? []).map((x) => x.toLowerCase());
    switch (g) {
      case 'everyone':
      case 'current':
        return true;
      case 'none':
        return false;
      case 'editors':
        return this.options.role === 'editor' || this.options.role === 'owner' || groups.includes('editors');
      case 'owners':
        return this.options.role === 'owner' || groups.includes('owners');
      default:
        return groups.includes(g);
    }
  }

  /** Whether the current user is named by a permission range. */
  mayEdit(perm: { editor?: string; group?: string }): boolean {
    if (perm.editor && this.isCurrentUser(perm.editor)) return true;
    if (perm.group && this.matchesGroup(perm.group)) return true;
    return false;
  }

  /** For DocxViewer's `highlightRange`: shade the regions this user may edit. */
  readonly highlightRange = (area: PermArea): boolean => readProtection(this.doc).enforced && this.mayEdit(area);

  private permRanges(): PermRange[] {
    const root = this.doc.documentXml.root;
    const ends = new Map(findAll(root, 'w:permEnd').map((e) => [attr(e, 'w:id') ?? '', e]));
    const out: PermRange[] = [];
    for (const start of findAll(root, 'w:permStart')) {
      const id = attr(start, 'w:id') ?? '';
      const end = ends.get(id);
      if (end) out.push({ id, editor: attr(start, 'w:ed'), group: attr(start, 'w:edGrp'), start, end });
    }
    return out;
  }

  /** True when the whole range lies inside a permission range that names this user. */
  inEditableRange(range: DocRange): boolean {
    const index = new TreeIndex(this.doc.documentXml.root);
    if (!index.contains(range.start.node) || !index.contains(range.end.node)) return false;
    const a = index.order.get(range.start.node)!;
    const b = index.order.get(range.end.node)!;
    return this.permRanges().some((p) => this.mayEdit(p) && index.order.get(p.start)! < Math.min(a, b) && Math.max(a, b) < index.order.get(p.end)!);
  }
}
