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
import { activityLabel, numberedObjects, targetParagraphs } from '../../mcp/src/operations.js';
import { listParagraphs, readNumbersByView, renumberSelection, viewerParagraphs } from '../../mcp/src/outline.js';
import { anthropicStream, MODELS, runTurn, type ModelChoice, type StreamFactory } from './agent.js';
import { apiMessage, forgetKey, giveKey, keyProblem, NoRelayError, relayClient, relayStatus, START_THE_SERVER } from './auth.js';
import { renderMarkdown } from './markdown.js';
import { createToolbox, type OpenDocument, type ToolOutcome } from './tools.js';

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;

// ---- storage (every access guarded: storage can be blocked or unavailable) ------------------------
// Only the model choice is stored. The API key lives in the local relay's memory, never in the browser.

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
let current: OpenDocument | undefined;
let conversation: BetaMessageParam[] = [];
let busy: AbortController | undefined;
let selection: SelectionInfo | undefined;
const notes: string[] = [];

const fonts = new FontRegistry();
const fontsReady = loadBundledFonts(fonts, { baseUrl: new URL(import.meta.env.BASE_URL, location.href) }).catch((e: unknown) => {
  console.warn('Bundled fonts unavailable; using system font metrics', e);
});
const viewer = new DocxViewer($('viewer'), {
  fonts,
  showComments: true,
  editable: true,
  navigation: { collapsed: true },
  comments: { mode: 'auto' },
  contextMenu: {
    items: (at, defaults) => [...defaults, '-', { label: 'Ask Claude about this', icon: '✦', run: () => putSelectionInComposer(at.selection) }],
  },
});
const toolbox = createToolbox(() => current);

// How Claude's work shows in the document; remembered in this browser.
{
  const effects = $('effects') as HTMLSelectElement;
  const apply = (value: string) => {
    viewer.activity.enabled = value !== 'off';
    viewer.activity.animate = value === 'animated';
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
    viewer.activity.name = look.name.trim() || DEFAULT_LOOK.name;
    viewer.activity.color = look.color;
    viewer.activity.style = look.style;
    viewer.activity.marksShown = look.marks;
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

/** Keep the page, and its comment column when there are comments, inside the document panel. */
function fitToPanel(): void {
  const layout = viewer.layout;
  if (!layout || layout.pages.length === 0) return;
  const pageWidth = (Math.max(...layout.pages.map((p) => p.width)) * 96) / 72;
  // Row padding (32), the comment column as the viewer sized it (0 without comments) and a scrollbar (16).
  const extra = 32 + viewer.comments.width + 16;
  const zoom = Math.max(0.3, Math.min(1, ((viewer.root.clientWidth || $('viewer').clientWidth) - extra) / pageWidth));
  if (Math.abs(viewer.zoom - zoom) > 0.01) viewer.zoom = zoom;
}
viewer.on('layout', fitToPanel);
viewer.on('chromechange', ({ part }) => {
  if (part === 'comments') fitToPanel();
});
// Refit whenever the document's room changes: the window, and the navigation panel opening or
// folding beside it. Only real width changes count (a scrollbar coming and going does not).
{
  let timer: ReturnType<typeof setTimeout> | undefined;
  const refit = () => {
    clearTimeout(timer);
    timer = setTimeout(fitToPanel, 50);
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

function welcome(): void {
  messagesEl.replaceChildren();
  const el = addMessage('welcome');
  const p = document.createElement('p');
  p.textContent = current
    ? `${current.name} is open. Ask Claude anything about it, or tell it what to change. Select text in the document to point at it.`
    : 'Open a .docx (or the sample). Claude can read it, edit it, comment on it and fix lists and formatting while you watch.';
  el.append(p);
  if (current) {
    const box = document.createElement('div');
    box.className = 'suggestions';
    for (const s of SUGGESTIONS) {
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
  container.append(line);
}

// ---- selection in the viewer ------------------------------------------------------------------------

// What is selected (or where the caret is) goes to Claude with the next message, like an editor
// sharing the selected lines with a coding agent. The chip shows it; × leaves it out.
viewer.on('selectioninfo', ({ selection: info }) => {
  if (info) showSelection(info);
});

function showSelection(info: SelectionInfo): void {
  if (!viewer.document) return;
  // Claude's tools number paragraphs as read_document does, text boxes included; the viewer doesn't.
  if (current) info = renumberSelection(info, readNumbersByView(listParagraphs(current.doc)));
  selection = info;
  const box = $('selection');
  box.hidden = false;
  box.querySelector('span')?.remove();
  const span = document.createElement('span');
  const where = info.start.paragraph === info.end.paragraph ? `paragraph ${info.start.paragraph}` : `paragraphs ${info.start.paragraph}–${info.end.paragraph}`;
  const story = info.story === 'body' ? '' : ` of the ${info.story}`;
  const text = info.text.replace(/\s+/g, ' ').trim();
  span.textContent = info.collapsed || !text ? `Cursor in ${where}${story}` : `Selected in ${where}${story}: “${text.length > 80 ? `${text.slice(0, 79)}…` : text}”`;
  box.prepend(span);
}

/** "Ask Claude about this": put the clicked selection in the chip and move on to the message box. */
function putSelectionInComposer(info: SelectionInfo | undefined): void {
  if (info) showSelection(info);
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

async function openBytes(bytes: Uint8Array, name: string): Promise<void> {
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
  // Marks belong to the document Claude changed.
  viewer.activity.clearMarks();
  await viewer.load(doc);
  current = { doc, name };
  conversation = [];
  notes.length = 0;
  clearSelection();
  $('docname').textContent = name;
  $('docname').title = name;
  $('empty').hidden = true;
  $<HTMLButtonElement>('download').disabled = false;
  $<HTMLButtonElement>('pdf').disabled = false;
  welcome();
}

async function openFile(file: File): Promise<void> {
  try {
    await openBytes(new Uint8Array(await file.arrayBuffer()), file.name);
  } catch (err) {
    notice(`Could not open ${file.name}: ${(err as Error).message}`, true);
  }
}

$<HTMLInputElement>('file').addEventListener('change', (e) => {
  const input = e.target as HTMLInputElement;
  const file = input.files?.[0];
  input.value = '';
  if (file) void openFile(file);
});
for (const id of ['sample', 'sample2']) $(id).addEventListener('click', async () => openBytes(await sampleDocx(), 'sample.docx'));
document.addEventListener('dragover', (e) => e.preventDefault());
document.addEventListener('drop', (e) => {
  e.preventDefault();
  const file = e.dataTransfer?.files[0];
  if (file) void openFile(file);
});

function download(bytes: Uint8Array, name: string, type: string): void {
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

$('download').addEventListener('click', async () => {
  if (!current) return;
  download(await current.doc.save({ modified: new Date() }), current.name, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
});
$('pdf').addEventListener('click', async () => {
  if (!current) return;
  await fontsReady;
  const pdf = await exportPdf(current.doc, { fonts, ...(viewer.layout ? { layout: viewer.layout } : {}) });
  download(pdf, current.name.replace(/\.[^.]+$/, '') + '.pdf', 'application/pdf');
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

  // What Claude should know besides the words: the open document and the user's selection.
  const context: string[] = [];
  context.push(current ? `[Open document: ${current.name}]` : '[No document is open.]');
  if (selection && current) context.push(`[${describeSelection(selection, { document: current.name })}]`);
  for (const n of notes.splice(0)) context.push(`[${n}]`);
  conversation.push({ role: 'user', content: `${context.join('\n')}\n\n${text}` });

  const user = addMessage('user');
  if (selection && current && selection.text.trim()) {
    const q = document.createElement('span');
    q.className = 'quote';
    q.textContent = selection.text;
    user.append(q);
  }
  user.append(text);
  clearSelection();

  const reply = addMessage('assistant');
  const waiting = document.createElement('div');
  waiting.className = 'thinking';
  waiting.textContent = 'Thinking';
  reply.append(waiting);

  const doc = current?.doc;
  const before = doc ? await doc.save() : undefined;
  let changed = false;
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
  let beforeTexts: ReturnType<typeof viewer.activity.snapshot> | undefined;
  viewer.activity.clearMarks();
  const controller = new AbortController();
  setBusy(controller);
  try {
    await runTurn({
      stream,
      model,
      tools: toolbox,
      messages: conversation,
      signal: controller.signal,
      events: {
        step() {
          textEl = undefined;
          raw = '';
          drafted = '';
        },
        toolDraft(name, input) {
          // Show where Claude is working while it is still writing the call.
          const kind = toolbox.kind(name);
          if (!kind || !current || Date.now() - lastDraft < 120) return;
          lastDraft = Date.now();
          // Claude's numbers count text box paragraphs; the viewer's don't.
          const entries = listParagraphs(current.doc);
          const args = (input ?? {}) as Record<string, unknown>;
          const objects = args.object !== undefined ? numberedObjects(new DocxEditor(current.doc), entries) : undefined;
          const paragraphs = viewerParagraphs(entries, targetParagraphs(entries, name, args, objects));
          const key = `${name}:${paragraphs.join(',')}`;
          if (key === drafted) return;
          drafted = key;
          if (kind === 'read' && !paragraphs.length) viewer.activity.work([], { mode: 'read', label: activityLabel(name) });
          else viewer.activity.work(paragraphs, { mode: kind === 'read' ? 'read' : 'edit', label: activityLabel(name) });
        },
        toolStart(name, input) {
          if (!drafted.startsWith(`${name}:`)) this.toolDraft?.(name, input);
          drafted = '';
          beforeTexts = viewer.document ? viewer.activity.snapshot() : undefined;
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
          if (!outcome.isError && outcome.kind !== 'read') changed = true;
          // Animate what this call changed (ink, glow, change bars), or just stop the shimmer.
          if (!outcome.isError && outcome.kind !== 'read' && beforeTexts) void viewer.activity.reveal(beforeTexts);
          else viewer.activity.idle();
          beforeTexts = undefined;
          if (outcome.comment) {
            const id = outcome.comment;
            // The viewer redraws on the next frame; then show the new comment.
            requestAnimationFrame(() => requestAnimationFrame(() => viewer.scrollToComment(id)));
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
    if (!changed) viewer.activity.idle();
  }

  if (changed && before && current?.doc === doc) {
    // One click takes the document back to how it was before this reply.
    for (const old of messagesEl.querySelectorAll('.turn-actions')) old.remove();
    const actions = document.createElement('div');
    actions.className = 'turn-actions';
    const undo = document.createElement('button');
    undo.type = 'button';
    undo.textContent = "Undo Claude's changes";
    undo.title = 'Puts the document back to how it was before this reply (also undoes your own typing since then)';
    undo.addEventListener('click', async () => {
      if (!current) return;
      const restored = await DocxDocument.open(before);
      // The changes are gone, and so are their marks.
      viewer.activity.clearMarks();
      await viewer.load(restored);
      current = { doc: restored, name: current.name };
      notes.push('The user undid all changes from your previous reply; the document is back to how it was before that reply. Read it again before relying on paragraph numbers.');
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
