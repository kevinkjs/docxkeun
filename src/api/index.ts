export { DocxEditor, OBJECT_CHAR, paragraphText, paragraphsOf, applyFormat, runContent, lockModeOf, LockedRegionError, LOCK_TAG, type EditAction, type EditGuard, type TextFormat, type CommentInput, type Match, type LockMode, type LockedRegion, type SetListOptions, type EditorOptions, type ParagraphFormat, documentTracksChanges } from './editor.js';
export { listRevisions, acceptRevisions, rejectRevisions, type Revision, type RevisionKind, type RevisionFilter } from './revisions.js';
export { LIST_FORMATS, type ListFormat, type ListItem } from './lists.js';
export { EditError, TreeIndex, type DocPosition, type DocRange } from './positions.js';
export { ensurePart, randomHexId } from './parts.js';
export { selectionInfoFor, describeSelection, type SelectionInfo, type SelectedParagraph, type TextSpot } from './selection.js';
export type { HeaderFooterInfo, HeaderFooterKind, HeaderFooterOptions, HeaderFooterType } from './headers.js';
export type { TableOptions, CopyDropped } from './tables.js';
