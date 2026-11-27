/**
 * docxkeun chat: open a .docx, talk to Claude about it, and watch Claude's edits appear in the
 * docxkeun viewer as they happen. Runs entirely in the browser with the user's own API key.
 */
import Anthropic from '@anthropic-ai/sdk';
import type { BetaMessageParam } from '@anthropic-ai/sdk/resources/beta/messages/messages';
import { DocxDocument, FontRegistry, loadBundledFonts } from 'docxkeun';
import { describeSelection, DocxEditor, type SelectionInfo } from 'docxkeun/editor';
import { exportPdf } from 'docxkeun/pdf';
import { PasswordRequiredError, WrongPasswordError } from 'docxkeun/security';
import { ACTIVITY_STYLES, DocxViewer, type ActivityStyle } from 'docxkeun/viewer';
import { sampleDocx } from '../../../playground/sample.js';
import { MAX_INSTRUCTIONS } from '../../mcp/src/instructions.js';
import { activityLabel, numberedObjects, targetParagraphs } from '../../mcp/src/operations.js';
import { listParagraphs, readNumbersByView, renumberSelection, viewerParagraphs } from '../../mcp/src/outline.js';
import { anthropicStream, MODELS, runTurn, type ModelChoice, type StreamFactory } from './agent.js';
import { apiMessage, forgetKey, giveKey, keyProblem, NoRelayError, relayClient, relayStatus, START_THE_SERVER } from './auth.js';
import { renderMarkdown } from './markdown.js';
import { createToolbox, type OpenDocument, type ToolOutcome } from './tools.js';

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;

// ---- storage (every access guarded: storage can be blocked or unavailable) ------------------------
// The model choice, effects, appearance and instructions are remembered in this browser; the API key is not (it lives in the local relay's memory).

const MODEL = 'docxkeun-chat.model';

function read(store: () => Storage, name: string): string | undefined {
  try {
    return store().getItem(name) ?? undefined;
  } catch {
    return undefined;
  }
}
function write(store: () => Storage, name: string, value: string | undefined): void {
  try {
    if (value === undefined) store().removeItem(name);
    else store().setItem(name, value);
  } catch {
    // Not stored; nothing depends on it.
  }
}
const session = () => sessionStorage;
const local = () => localStorage;

// ---- state ------------------------------------------------------------------------------------------

const mock = import.meta.env.DEV && new URLSearchParams(location.search).has('mock');
let stream: StreamFactory | undefined;
// An earlier version kept the key in the browser: remove anything it left behind.
for (const store of [session, local]) for (const name of ['docxkeun-chat.key', 'docxkeun-chat.auth']) write(store, name, undefined);
let model: ModelChoice = MODELS.find((m) => m.id === read(local, MODEL)) ?? MODELS[0]!;
let conversation: BetaMessageParam[] = [];
let busy: AbortController | undefined;
const notes: string[] = [];
/** The user's own instructions for Claude (Instructions dialog); sent with every turn. */
let instructions = '';

const fonts = new FontRegistry();
const fontsReady = loadBundledFonts(fonts, { baseUrl: new URL(import.meta.env.BASE_URL, location.href) }).catch((e: unknown) => {
  console.warn('Bundled fonts unavailable; using system font metrics', e);
});

/** A document pane: its own viewer, bar and file. The second one shows once the user asks for it. */
interface Pane {
  index: number;
  el: HTMLElement;
  viewer: DocxViewer;
  open: OpenDocument | undefined;
}

const panes: Pane[] = [...document.querySelectorAll<HTMLElement>('.pane')].map((el, index) => {
  const pane: Pane = { index, el, open: undefined } as Pane;
  pane.viewer = new DocxViewer(el.querySelector<HTMLElement>('.viewer')!, {
    fonts,
    showComments: true,
    editable: true,
    navigation: { collapsed: true },
    comments: { mode: 'auto' },
    contextMenu: {
      items: (at, defaults) => [...defaults, '-', { label: 'Ask Claude about this', icon: '✦', run: () => putSelectionInComposer(at.selection, pane) }],
    },
  });
  return pane;
});
/** The open documents, in pane order: what Claude's tools can work on. */
const openDocuments = (): OpenDocument[] => panes.flatMap((p) => (p.open ? [p.open] : []));
const paneOf = (open: OpenDocument | undefined): Pane | undefined => (open ? panes.find((p) => p.open === open) : undefined);
const toolbox = createToolbox(openDocuments);

/** What the user selected, and in which pane. */
let selection: { info: SelectionInfo; pane: Pane } | undefined;

// How Claude's work shows in the documents; remembered in this browser.
{
  const effects = $('effects') as HTMLSelectElement;
  const apply = (value: string) => {
    for (const { viewer } of panes) {
      viewer.activity.enabled = value !== 'off';
      viewer.activity.animate = value === 'animated';
    }
  };
  try {
    const saved = localStorage.getItem('docxkeun-ai-effects');
    if (saved && ['animated', 'still', 'off'].includes(saved)) effects.value = saved;
  } catch {
    // Storage blocked: the default (animated) stays.
  }
  apply(effects.value);
  effects.addEventListener('change', () => {
    apply(effects.value);
    try {
      localStorage.setItem('docxkeun-ai-effects', effects.value);
    } catch {
      // Not remembered, still applied.
    }
  });
}

// How Claude shows on the pages (name, colours, effect, marks); remembered in this browser.
{
  type Look = { name: string; color: [string, string]; style: ActivityStyle; marks: boolean };
  const DEFAULT_LOOK: Look = { name: 'Claude', color: ['#7c5cff', '#d97757'], style: 'highlight', marks: true };
  const KEY = 'docxkeun-ai-look';
  const dialog = $('lookDialog') as HTMLDialogElement;
  const name = $('lookName') as HTMLInputElement;
  const from = $('lookColor1') as HTMLInputElement;
  const to = $('lookColor2') as HTMLInputElement;
  const style = $('lookStyle') as HTMLSelectElement;
  const marks = $('lookMarks') as HTMLInputElement;
  const sample = dialog.querySelector('.look-sample') as HTMLElement;
  for (const s of ACTIVITY_STYLES) style.add(new Option(s.label, s.style));
  const isHex = (v: unknown): v is string => typeof v === 'string' && /^#[0-9a-f]{6}$/i.test(v);
  const isStyle = (v: unknown): v is ActivityStyle => ACTIVITY_STYLES.some((s) => s.style === v);
  const saved = (): Look | undefined => {
    try {
      const s: unknown = JSON.parse(localStorage.getItem(KEY) ?? 'null');
      if (s && typeof s === 'object' && 'name' in s && 'color' in s) {
        const { name: n, color: c, style: st, marks: m } = s as { name: unknown; color: unknown; style?: unknown; marks?: unknown };
        if (typeof n === 'string' && Array.isArray(c) && c.length === 2 && isHex(c[0]) && isHex(c[1])) {
          // Saved before effects could be chosen: the defaults for those.
          return { name: n.slice(0, 40), color: [c[0], c[1]], style: isStyle(st) ? st : DEFAULT_LOOK.style, marks: typeof m === 'boolean' ? m : DEFAULT_LOOK.marks };
        }
      }
    } catch {
      // Storage blocked or unreadable: the defaults.
    }
    return undefined;
  };
  let look = saved() ?? DEFAULT_LOOK;
  const apply = () => {
    for (const { viewer } of panes) {
      viewer.activity.name = look.name.trim() || DEFAULT_LOOK.name;
      viewer.activity.color = look.color;
      viewer.activity.style = look.style;
      viewer.activity.marksShown = look.marks;
    }
    const viewer = panes[0]!.viewer;
    // The dialog's sample uses the same colours and effect, and the label text colour the viewer picked.
    sample.style.setProperty('--look-1', look.color[0]);
    sample.style.setProperty('--look-2', look.color[1]);
    sample.style.setProperty('--look-text', viewer.root.style.getPropertyValue('--docxkeun-ai-text') || '#fff');
    sample.dataset.style = look.style;
    $('lookLabel').textContent = `${viewer.activity.name} is editing…`;
    $('lookStyleNote').textContent = ACTIVITY_STYLES.find((s) => s.style === look.style)?.description ?? '';
    $('lookMarksNote').textContent = look.marks ? markNote(look.style) : 'Nothing stays on the page once a change has landed.';
  };
  const markNote = (s: ActivityStyle) =>
    s === 'highlight'
      ? 'The changed paragraph keeps a soft tint until the next change.'
      : s === 'words'
        ? 'The changed words keep these colours until the next change.'
        : s === 'underline'
          ? 'The changed words stay underlined until the next change.'
          : 'A ✦ bar in the margin marks each changed paragraph until the next change (hover it for the text before).';
  const fill = () => {
    name.value = look.name;
    from.value = look.color[0];
    to.value = look.color[1];
    style.value = look.style;
    marks.checked = look.marks;
  };
  apply();
  for (const input of [name, from, to, style, marks]) {
    input.addEventListener(input === style || input === marks ? 'change' : 'input', () => {
      look = { name: name.value, color: [from.value, to.value], style: isStyle(style.value) ? style.value : DEFAULT_LOOK.style, marks: marks.checked };
      apply();
      try {
        localStorage.setItem(KEY, JSON.stringify(look));
      } catch {
        // Not remembered, still applied.
      }
    });
  }
  $('lookReset').addEventListener('click', () => {
    look = DEFAULT_LOOK;
    fill();
    apply();
    try {
      localStorage.removeItem(KEY);
    } catch {
      // Nothing to forget.
    }
  });
  $('look').addEventListener('click', () => {
    fill();
    dialog.showModal();
    name.focus();
  });
}

// The user's own instructions for Claude; remembered in this browser, and kept when another document opens.
{
  const KEY = 'docxkeun-ai-instructions';
  const dialog = $('instructionsDialog') as HTMLDialogElement;
  const text = $('instructionsText') as HTMLTextAreaElement;
  const button = $('instructions');
  text.maxLength = MAX_INSTRUCTIONS;
  try {
    instructions = (localStorage.getItem(KEY) ?? '').slice(0, MAX_INSTRUCTIONS);
  } catch {
    // Storage blocked: none until typed.
  }
  const apply = () => {
    const shown = instructions.trim();
    button.textContent = shown ? 'Instructions ●' : 'Instructions';
    if (shown) button.setAttribute('aria-label', 'Instructions (on)');
    else button.removeAttribute('aria-label');
    button.title = shown
      ? `Your instructions for Claude (on): ${shown.slice(0, 300)}${shown.length > 300 ? '…' : ''}`
      : 'Your own instructions for Claude, added to its built-in ones';
    $('instructionsCount').textContent = `${text.value.length} / ${MAX_INSTRUCTIONS}`;
  };
  text.addEventListener('input', () => {
    instructions = text.value;
    apply();
    try {
      localStorage.setItem(KEY, instructions);
    } catch {
      // Not remembered, still applied.
    }
  });
  $('instructionsClear').addEventListener('click', () => {
    instructions = '';
    text.value = '';
    apply();
    try {
      localStorage.removeItem(KEY);
    } catch {
      // Nothing to forget.
    }
    text.focus();
  });
  button.addEventListener('click', () => {
    text.value = instructions;
    apply();
    dialog.showModal();
    text.focus();
  });
  text.value = instructions;
  apply();
}

/** Keep a pane's page, and its comment column when there are comments, inside the pane. */
function fitToPanel(viewer: DocxViewer): void {
  const layout = viewer.layout;
  if (!layout || layout.pages.length === 0 || !viewer.root.clientWidth) return;
  const pageWidth = (Math.max(...layout.pages.map((p) => p.width)) * 96) / 72;
  // Row padding (32), the comment column as the viewer sized it (0 without comments) and a scrollbar (16).
  const extra = 32 + viewer.comments.width + 16;
  const zoom = Math.max(0.3, Math.min(1, (viewer.root.clientWidth - extra) / pageWidth));
  if (Math.abs(viewer.zoom - zoom) > 0.01) viewer.zoom = zoom;
}
for (const { viewer } of panes) {
  viewer.on('layout', () => fitToPanel(viewer));
  viewer.on('chromechange', ({ part }) => {
    if (part === 'comments') fitToPanel(viewer);
  });
  // Refit whenever the pane's room changes: the window, the other pane opening or closing, and the
  // navigation panel opening or folding. Only real width changes count (not a scrollbar).
  let timer: ReturnType<typeof setTimeout> | undefined;
  const refit = () => {
    clearTimeout(timer);
    timer = setTimeout(() => fitToPanel(viewer), 50);
  };
  if (typeof ResizeObserver === 'undefined') window.addEventListener('resize', refit);
  else {
    let width = viewer.root.offsetWidth;
    new ResizeObserver(() => {
      if (viewer.root.offsetWidth === width) return;
      width = viewer.root.offsetWidth;
      refit();
    }).observe(viewer.root);
  }
}

/** Problem to show when the dialog opens at start (no relay, for example). */
let startProblem: string | undefined;
let keyFromEnv = false;
if (mock) {
  const { mockStream } = await import('./mock.js');
  stream = mockStream;
} else {
  try {
    const status = await relayStatus();
    keyFromEnv = status.fromEnv;
    if (status.hasKey) stream = anthropicStream(relayClient());
  } catch (err) {
    startProblem = err instanceof NoRelayError ? START_THE_SERVER : (err as Error).message;
  }
}

// ---- API key dialog ---------------------------------------------------------------------------------

const keyDialog = $<HTMLDialogElement>('keyDialog');
const keyInput = $<HTMLInputElement>('keyInput');
const keyError = $('keyError');

function openKeyDialog(message?: string): void {
  keyInput.value = '';
  $('forget').hidden = mock || !stream;
  $('envNote').hidden = !keyFromEnv;
  keyError.hidden = !message;
  keyError.textContent = message ?? '';
  if (!keyDialog.open) keyDialog.showModal();
  keyInput.focus();
}

// Without a key there is nothing to do: the dialog cannot be dismissed until one is given.
// (Browsers may ignore preventDefault on Escape, so a dialog closed without a key reopens.)
keyDialog.addEventListener('cancel', (e) => {
  if (!stream) e.preventDefault();
});
keyDialog.addEventListener('close', () => {
  if (!stream) setTimeout(() => openKeyDialog(), 0);
});

$('keyForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const key = keyInput.value.trim();
  if (!key) return;
  const submit = $<HTMLButtonElement>('keySubmit');
  submit.disabled = true;
  submit.textContent = 'Checking…';
  keyError.hidden = true;
  try {
    if (!mock) {
      // The relay checks the key with the API and keeps it in its memory; the page keeps nothing.
      await giveKey(key, model.id);
      stream = anthropicStream(relayClient());
      keyFromEnv = false;
    }
    keyInput.value = '';
    keyDialog.close();
    $('prompt').focus();
  } catch (err) {
    keyError.hidden = false;
    keyError.textContent = keyProblem(err, key);
  } finally {
    submit.disabled = false;
    submit.textContent = 'Continue';
  }
});

$('forget').addEventListener('click', async () => {
  try {
    await forgetKey();
    stream = undefined;
    keyFromEnv = false;
    openKeyDialog('The chat server has forgotten the key.');
  } catch (err) {
    keyError.hidden = false;
    keyError.textContent = keyProblem(err, '');
  }
});
$('key').addEventListener('click', () => openKeyDialog());

// ---- model ------------------------------------------------------------------------------------------

const modelSelect = $<HTMLSelectElement>('model');
modelSelect.replaceChildren(...MODELS.map((m) => new Option(m.label, m.id)));
modelSelect.value = model.id;
modelSelect.addEventListener('change', () => {
  model = MODELS.find((m) => m.id === modelSelect.value) ?? MODELS[0]!;
  write(local, MODEL, model.id);
});

// ---- messages ---------------------------------------------------------------------------------------

const messagesEl = $('messages');

function scrollDown(): void {
  messagesEl.scrollTop = messagesEl.scrollHeight;
}

function addMessage(kind: string): HTMLDivElement {
  const el = document.createElement('div');
  el.className = `msg ${kind}`;
  messagesEl.append(el);
  scrollDown();
  return el;
}

function notice(text: string, error = false): void {
  const el = addMessage('assistant');
  const n = document.createElement('div');
  n.className = error ? 'notice error' : 'notice';
  n.textContent = text;
  el.append(n);
}

const SUGGESTIONS = ['Summarise this document', 'Find anything risky and comment on it', 'Fix spelling and grammar mistakes', 'Turn the lettered lists into numbered lists'];
const SUGGESTIONS_FOR_TWO = ['Compare the two documents and list the differences', 'Summarise both documents', 'Comment on the second document where it differs from the first'];

function welcome(): void {
  messagesEl.replaceChildren();
  const el = addMessage('welcome');
  const p = document.createElement('p');
  const open = openDocuments();
  p.textContent =
    open.length === 2
      ? `${open[0]!.name} and ${open[1]!.name} are open side by side. Ask Claude about either or both, and say which one you mean. Select text in a document to point at it.`
      : open.length === 1
        ? `${open[0]!.name} is open. Ask Claude anything about it, or tell it what to change. Select text in the document to point at it. Use “+ Second” to work on two documents at once.`
        : 'Open a .docx (or the sample). Claude can read it, edit it, comment on it and fix lists and formatting while you watch.';
  el.append(p);
  if (open.length) {
    const box = document.createElement('div');
    box.className = 'suggestions';
    for (const s of open.length === 2 ? SUGGESTIONS_FOR_TWO : SUGGESTIONS) {
      const b = document.createElement('button');
      b.type = 'button';
      b.textContent = s;
      b.addEventListener('click', () => void send(s));
      box.append(b);
    }
    el.append(box);
  }
}

const TOOL_LABELS: Record<string, (input: Record<string, unknown>) => string> = {
  read_document: () => 'Read the document',
  find_text: (i) => `Searched for “${String(i.query ?? '')}”`,
  list_comments: () => 'Read the comments',
  list_links: () => 'Read the links and bookmarks',
};

function toolLine(container: HTMLElement, name: string, input: unknown, outcome: ToolOutcome): void {
  const line = document.createElement('div');
  const label = TOOL_LABELS[name];
  line.className = `tool ${outcome.isError ? 'error' : outcome.kind === 'read' ? 'read' : 'change'}`;
  const icon = document.createElement('span');
  icon.className = 'icon';
  icon.textContent = outcome.isError ? '⚠' : outcome.kind === 'read' ? '◦' : '✎';
  const text = document.createElement('span');
  text.textContent = outcome.isError
    ? `Couldn't ${name.replace(/_/g, ' ')}: ${outcome.content.split('\n')[0]}`
    : label
      ? label((input ?? {}) as Record<string, unknown>)
      : outcome.content.split('\n')[0]!;
  line.append(icon, text);
  // With two documents open, say which one the call was about.
  if (outcome.document && openDocuments().length > 1) {
    const doc = document.createElement('span');
    doc.className = 'tool-doc';
    doc.textContent = outcome.document.name;
    line.append(doc);
  }
  container.append(line);
}

// ---- selection in the viewer ------------------------------------------------------------------------

// What is selected (or where the caret is) goes to Claude with the next message, like an editor
// sharing the selected lines with a coding agent. The chip shows it; × leaves it out.
for (const pane of panes) {
  pane.viewer.on('selectioninfo', ({ selection: info }) => {
    if (info) showSelection(info, pane);
  });
}

function showSelection(info: SelectionInfo, pane: Pane): void {
  if (!pane.viewer.document || !pane.open) return;
  // Claude's tools number paragraphs as read_document does, text boxes included; the viewer doesn't.
  info = renumberSelection(info, readNumbersByView(listParagraphs(pane.open.doc)));
  selection = { info, pane };
  const box = $('selection');
  box.hidden = false;
  box.querySelector('span')?.remove();
  const span = document.createElement('span');
  const where = info.start.paragraph === info.end.paragraph ? `paragraph ${info.start.paragraph}` : `paragraphs ${info.start.paragraph}–${info.end.paragraph}`;
  const story = info.story === 'body' ? '' : ` of the ${info.story}`;
  // With two documents open, the chip says which one.
  const doc = openDocuments().length > 1 ? ` of ${pane.open.name}` : '';
  const text = info.text.replace(/\s+/g, ' ').trim();
  span.textContent = info.collapsed || !text ? `Cursor in ${where}${story}${doc}` : `Selected in ${where}${story}${doc}: “${text.length > 80 ? `${text.slice(0, 79)}…` : text}”`;
  box.prepend(span);
}

/** "Ask Claude about this": put the clicked selection in the chip and move on to the message box. */
function putSelectionInComposer(info: SelectionInfo | undefined, pane: Pane): void {
  if (info) showSelection(info, pane);
  $<HTMLTextAreaElement>('prompt').focus();
}

function clearSelection(): void {
  selection = undefined;
  $('selection').hidden = true;
}

{
  const close = document.createElement('button');
  close.type = 'button';
  close.textContent = '×';
  close.title = 'Do not send the selection';
  close.addEventListener('click', clearSelection);
  $('selection').append(close);
}

// ---- documents --------------------------------------------------------------------------------------

const paneEl = <T extends HTMLElement>(pane: Pane, selector: string) => pane.el.querySelector<T>(selector)!;

/** Claude tells the documents apart by name: a second file with the same name gets " (2)". */
function uniqueName(name: string, pane: Pane): string {
  const taken = (n: string) => panes.some((p) => p !== pane && p.open?.name.toLowerCase() === n.toLowerCase());
  if (!taken(name)) return name;
  const dot = name.lastIndexOf('.');
  const [stem, ext] = dot > 0 ? [name.slice(0, dot), name.slice(dot)] : [name, ''];
  for (let i = 2; ; i++) if (!taken(`${stem} (${i})${ext}`)) return `${stem} (${i})${ext}`;
}

/** Show a pane's file, or its empty state, in its bar; and the "+ Second document" button while there is room. */
function showPaneState(pane: Pane): void {
  const name = pane.open?.name ?? '';
  paneEl(pane, '.docname').textContent = name;
  paneEl(pane, '.docname').title = name;
  paneEl(pane, '.empty').hidden = !!pane.open;
  // The bar's Sample button is for an empty pane (the empty state has its own link too).
  paneEl(pane, '.pane-bar .pane-sample').hidden = !!pane.open;
  paneEl<HTMLButtonElement>(pane, '.pane-download').disabled = !pane.open;
  paneEl<HTMLButtonElement>(pane, '.pane-pdf').disabled = !pane.open;
  paneEl(panes[0]!, '.pane-add').hidden = !panes[1]!.el.hidden;
}

async function openBytes(bytes: Uint8Array, name: string, pane: Pane): Promise<void> {
  let doc: DocxDocument;
  try {
    doc = await DocxDocument.open(bytes);
  } catch (err) {
    if (!(err instanceof PasswordRequiredError)) throw err;
    const password = window.prompt(`${name} is password protected. Enter its password to open it:`);
    if (!password) return;
    try {
      doc = await DocxDocument.open(bytes, { password });
    } catch (e) {
      if (e instanceof WrongPasswordError) throw new Error('Wrong password.');
      throw e;
    }
  }
  await fontsReady;
  const previous = pane.open;
  name = uniqueName(name, pane);
  pane.el.hidden = false;
  // Marks belong to the document Claude changed.
  pane.viewer.activity.clearMarks();
  await pane.viewer.load(doc);
  pane.open = { doc, name };
  if (selection?.pane === pane) clearSelection();
  showPaneState(pane);
  // A conversation under way carries on: Claude hears what changed. Before the first message, the welcome shows what is open.
  if (!conversation.length) welcome();
  else {
    const other = panes.find((p) => p !== pane && p.open)?.open;
    notes.push(
      previous
        ? `The user replaced ${previous.name} with ${name}. Paragraph numbers you saw in ${previous.name} no longer apply; read ${name} before working on it.`
        : `The user opened ${name}${other ? ` beside ${other.name}` : ''}. Read it before working on it.`,
    );
    notice(previous ? `${name} replaced ${previous.name}.` : `${name} is open${other ? ` beside ${other.name}` : ''}.`);
  }
}

async function openFile(file: File, pane: Pane): Promise<void> {
  try {
    await openBytes(new Uint8Array(await file.arrayBuffer()), file.name, pane);
  } catch (err) {
    notice(`Could not open ${file.name}: ${(err as Error).message}`, true);
  }
}

/** Close the second document: Claude can no longer change it, and hears so. */
function closePane(pane: Pane): void {
  const closed = pane.open;
  pane.open = undefined;
  pane.viewer.activity.clearMarks();
  pane.el.hidden = true;
  if (selection?.pane === pane) clearSelection();
  showPaneState(pane);
  if (!closed) return;
  if (!conversation.length) welcome();
  else {
    notes.push(`The user closed ${closed.name}; only ${openDocuments().map((d) => d.name).join(' and ') || 'no document'} is open now.`);
    notice(`${closed.name} is closed.`);
  }
}

function download(bytes: Uint8Array, name: string, type: string): void {
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

for (const pane of panes) {
  paneEl<HTMLInputElement>(pane, '.pane-file').addEventListener('change', (e) => {
    const input = e.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = '';
    if (file) void openFile(file, pane);
  });
  for (const b of pane.el.querySelectorAll('.pane-sample')) b.addEventListener('click', async () => openBytes(await sampleDocx(), 'sample.docx', pane));
  paneEl(pane, '.pane-download').addEventListener('click', async () => {
    if (!pane.open) return;
    download(await pane.open.doc.save({ modified: new Date() }), pane.open.name, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  });
  paneEl(pane, '.pane-pdf').addEventListener('click', async () => {
    if (!pane.open) return;
    await fontsReady;
    const pdf = await exportPdf(pane.open.doc, { fonts, ...(pane.viewer.layout ? { layout: pane.viewer.layout } : {}) });
    download(pdf, pane.open.name.replace(/\.[^.]+$/, '') + '.pdf', 'application/pdf');
  });
}
// "+ Second document" shows the second pane, ready for a file; × closes it again.
paneEl(panes[0]!, '.pane-add').addEventListener('click', () => {
  panes[1]!.el.hidden = false;
  showPaneState(panes[1]!);
  paneEl<HTMLInputElement>(panes[1]!, '.pane-file').click();
});
paneEl(panes[1]!, '.pane-close').addEventListener('click', () => closePane(panes[1]!));
for (const pane of panes) showPaneState(pane);

// Drop a file on a pane to open it there; drop two files to open them side by side.
document.addEventListener('dragover', (e) => e.preventDefault());
document.addEventListener('drop', (e) => {
  e.preventDefault();
  const files = [...(e.dataTransfer?.files ?? [])].slice(0, 2);
  if (files.length === 2) {
    void openFile(files[0]!, panes[0]!).then(() => openFile(files[1]!, panes[1]!));
    return;
  }
  const on = panes.find((p) => !p.el.hidden && p.el.contains(e.target as Node));
  if (files[0]) void openFile(files[0], on ?? panes[0]!);
});

// ---- sending ----------------------------------------------------------------------------------------

const prompt = $<HTMLTextAreaElement>('prompt');

function setBusy(controller: AbortController | undefined): void {
  busy = controller;
  $<HTMLButtonElement>('send').disabled = !!controller;
  $('stop').hidden = !controller;
}

async function send(text: string): Promise<void> {
  text = text.trim();
  if (!text || busy) return;
  if (!stream) {
    openKeyDialog();
    return;
  }
  prompt.value = '';

  // What Claude should know besides the words: the open documents and the user's selection.
  const open = openDocuments();
  const context: string[] = [];
  context.push(
    open.length === 2
      ? `[Open documents: ${open[0]!.name} and ${open[1]!.name}. Name the one you work on as \`document\` in every tool call.]`
      : open.length === 1
        ? `[Open document: ${open[0]!.name}]`
        : '[No document is open.]',
  );
  const selected = selection?.pane.open ? { info: selection.info, name: selection.pane.open.name } : undefined;
  if (selected) context.push(`[${describeSelection(selected.info, { document: selected.name })}]`);
  for (const n of notes.splice(0)) context.push(`[${n}]`);
  conversation.push({ role: 'user', content: `${context.join('\n')}\n\n${text}` });

  const user = addMessage('user');
  if (selected && selected.info.text.trim()) {
    const q = document.createElement('span');
    q.className = 'quote';
    q.textContent = selected.info.text;
    user.append(q);
  }
  user.append(text);
  clearSelection();

  const reply = addMessage('assistant');
  const waiting = document.createElement('div');
  waiting.className = 'thinking';
  waiting.textContent = 'Thinking';
  reply.append(waiting);

  // Each open document as it was before this reply, for "Undo Claude's changes".
  const before = new Map<OpenDocument, Uint8Array>(await Promise.all(open.map(async (d) => [d, await d.doc.save()] as const)));
  const changed = new Set<OpenDocument>();
  let textEl: HTMLDivElement | undefined;
  let raw = '';
  let frame = 0;
  const render = () => {
    frame = 0;
    if (textEl) textEl.innerHTML = renderMarkdown(raw);
    scrollDown();
  };

  let drafted = '';
  let lastDraft = 0;
  /** The pane Claude is working in (its shimmer shows there), and its text before the call runs. */
  let working: Pane | undefined;
  let beforeTexts: { pane: Pane; snapshot: ReturnType<DocxViewer['activity']['snapshot']> } | undefined;
  const workIn = (pane: Pane | undefined) => {
    if (working && working !== pane) working.viewer.activity.idle();
    working = pane;
  };
  for (const p of panes) p.viewer.activity.clearMarks();
  const controller = new AbortController();
  setBusy(controller);
  try {
    await runTurn({
      stream,
      model,
      tools: toolbox,
      messages: conversation,
      instructions,
      signal: controller.signal,
      events: {
        step() {
          textEl = undefined;
          raw = '';
          drafted = '';
        },
        toolDraft(name, input) {
          // Show where Claude is working while it is still writing the call: in the document it names
          // (with two open, nothing shows until the call says which).
          const kind = toolbox.kind(name);
          const target = toolbox.target(input);
          const pane = paneOf(target);
          if (!kind || !target || !pane || Date.now() - lastDraft < 120) return;
          lastDraft = Date.now();
          // Claude's numbers count text box paragraphs; the viewer's don't.
          const entries = listParagraphs(target.doc);
          const args = (input ?? {}) as Record<string, unknown>;
          const objects = args.object !== undefined ? numberedObjects(new DocxEditor(target.doc), entries) : undefined;
          const paragraphs = viewerParagraphs(entries, targetParagraphs(entries, name, args, objects));
          const key = `${name}:${pane.index}:${paragraphs.join(',')}`;
          if (key === drafted) return;
          drafted = key;
          workIn(pane);
          if (kind === 'read' && !paragraphs.length) pane.viewer.activity.work([], { mode: 'read', label: activityLabel(name) });
          else pane.viewer.activity.work(paragraphs, { mode: kind === 'read' ? 'read' : 'edit', label: activityLabel(name) });
        },
        toolStart(name, input) {
          if (!drafted.startsWith(`${name}:`)) this.toolDraft?.(name, input);
          drafted = '';
          const pane = paneOf(toolbox.target(input));
          beforeTexts = pane?.viewer.document ? { pane, snapshot: pane.viewer.activity.snapshot() } : undefined;
        },
        text(delta) {
          waiting.remove();
          if (!textEl) {
            textEl = document.createElement('div');
            textEl.className = 'text';
            reply.append(textEl);
          }
          raw += delta;
          frame ||= requestAnimationFrame(render);
        },
        tool(name, input, outcome) {
          waiting.remove();
          if (frame) render();
          textEl = undefined;
          raw = '';
          toolLine(reply, name, input, outcome);
          const pane = paneOf(outcome.document);
          const wrote = !outcome.isError && outcome.kind !== 'read' && outcome.document;
          if (wrote) changed.add(outcome.document!);
          // Animate what this call changed (ink, glow, change bars) where it changed, or just stop the shimmer.
          workIn(pane);
          if (wrote && pane && beforeTexts?.pane === pane) void pane.viewer.activity.reveal(beforeTexts.snapshot);
          else pane?.viewer.activity.idle();
          working = undefined;
          beforeTexts = undefined;
          if (outcome.comment && pane) {
            const id = outcome.comment;
            // The viewer redraws on the next frame; then show the new comment.
            requestAnimationFrame(() => requestAnimationFrame(() => pane.viewer.scrollToComment(id)));
          }
          scrollDown();
        },
        notice(text, kind) {
          const n = document.createElement('div');
          n.className = kind === 'refusal' ? 'notice' : 'notice';
          n.textContent = text;
          reply.append(n);
        },
      },
    });
  } catch (err) {
    const n = document.createElement('div');
    n.className = 'notice error';
    if (err instanceof Anthropic.APIUserAbortError || controller.signal.aborted) {
      n.className = 'notice';
      n.textContent = 'Stopped.';
    } else if (err instanceof Anthropic.AuthenticationError) {
      // Also what the relay answers when it has no key (it was restarted, for example).
      const said = apiMessage(err);
      n.textContent = `The API key was not accepted${said ? ` (${said})` : ''}.`;
      stream = undefined;
      openKeyDialog(`${said ?? 'The API key was not accepted.'} Enter it again.`);
    } else if (err instanceof Anthropic.RateLimitError) {
      n.textContent = 'Rate limited by the API. Wait a moment and try again.';
    } else if (err instanceof Anthropic.APIError) {
      n.textContent = `The API returned an error (${err.status ?? 'no status'}): ${apiMessage(err) ?? err.message}`;
    } else {
      n.textContent = `Something went wrong: ${(err as Error).message}`;
    }
    reply.append(n);
  } finally {
    if (frame) render();
    waiting.remove();
    setBusy(undefined);
    for (const p of panes) if (![...changed].some((d) => p.open === d)) p.viewer.activity.idle();
  }

  // The documents this reply changed that are still open where they were.
  const undoable = [...changed].filter((d) => paneOf(d) && before.has(d));
  if (undoable.length) {
    // One click takes them back to how they were before this reply.
    for (const old of messagesEl.querySelectorAll('.turn-actions')) old.remove();
    const actions = document.createElement('div');
    actions.className = 'turn-actions';
    const undo = document.createElement('button');
    undo.type = 'button';
    undo.textContent = "Undo Claude's changes";
    const names = undoable.map((d) => d.name).join(' and ');
    undo.title = `Puts ${names} back to how ${undoable.length > 1 ? 'they were' : 'it was'} before this reply (also undoes your own typing since then)`;
    undo.addEventListener('click', async () => {
      const restoredNames: string[] = [];
      for (const d of undoable) {
        const pane = paneOf(d);
        if (!pane) continue; // Closed or replaced since: nothing to put back there.
        const restored = await DocxDocument.open(before.get(d)!);
        // The changes are gone, and so are their marks.
        pane.viewer.activity.clearMarks();
        await pane.viewer.load(restored);
        pane.open = { doc: restored, name: d.name };
        if (selection?.pane === pane) clearSelection();
        restoredNames.push(d.name);
      }
      if (restoredNames.length) {
        const list = restoredNames.join(' and ');
        notes.push(`The user undid all changes from your previous reply to ${list}; ${restoredNames.length > 1 ? 'they are' : 'it is'} back to how ${restoredNames.length > 1 ? 'they were' : 'it was'} before that reply. Read again before relying on paragraph numbers.`);
      }
      actions.replaceChildren(Object.assign(document.createElement('span'), { className: 'thinking-done', textContent: 'Changes undone.' }));
    });
    actions.append(undo);
    reply.append(actions);
  }
  scrollDown();
}

$('composer').addEventListener('submit', (e) => {
  e.preventDefault();
  void send(prompt.value);
});
prompt.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) {
    e.preventDefault();
    void send(prompt.value);
  }
});
$('stop').addEventListener('click', () => busy?.abort());

// ---- start --------------------------------------------------------------------------------------------

welcome();
if (!stream) openKeyDialog(startProblem);
