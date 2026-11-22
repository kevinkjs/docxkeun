/** A small titled popover with a text input (OK/Cancel), used for link and bookmark prompts. */
import { ensureUiStyle } from './theme.js';
import { positionMenu } from './menu.js';

export interface PromptOptions {
  title: string;
  value?: string;
  placeholder?: string;
  /** Viewport point (the popover is `position: fixed`). */
  at: { x: number; y: number };
  /** Return a message to show and keep the popover open, or undefined to accept. */
  validate?: (v: string) => string | undefined;
  /** Where to append the popover (default `document.body`). */
  container?: HTMLElement;
  /** Aborting closes the popover as if cancelled (resolves `undefined`). */
  signal?: AbortSignal;
}

const STYLE_ID = 'docxkeun-popover-style';
const CSS = `
.dk-popover{position:fixed;z-index:2147483000;width:260px;max-width:calc(100vw - 8px);box-sizing:border-box;padding:10px;display:flex;flex-direction:column;gap:8px;background:var(--dk-ui-bg);color:var(--dk-ui-fg);border:1px solid var(--dk-ui-line);border-radius:var(--dk-ui-radius);box-shadow:var(--dk-ui-shadow);font:var(--dk-ui-font)}
.dk-popover-title{font-weight:600}
.dk-popover input{height:28px;box-sizing:border-box;width:100%;padding:0 8px;border:1px solid var(--dk-ui-line);border-radius:4px;background:var(--dk-ui-bg);color:inherit;font:inherit}
.dk-popover input:focus-visible,.dk-popover button:focus-visible{outline:2px solid var(--dk-ui-accent);outline-offset:1px}
.dk-popover-error{color:#b3261e;font-size:12px}
.dk-popover-error:empty{display:none}
.dk-popover-actions{display:flex;justify-content:flex-end;gap:6px}
.dk-popover button{height:28px;padding:0 12px;border:1px solid var(--dk-ui-line);border-radius:4px;background:var(--dk-ui-bg);color:inherit;font:inherit;cursor:pointer}
.dk-popover button:hover{background:var(--dk-ui-hover)}
.dk-popover button[data-role=ok]{background:var(--dk-ui-accent);border-color:var(--dk-ui-accent);color:#fff}
`;

function ensureStyle(): void {
  ensureUiStyle();
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement('style');
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.appendChild(style);
}

/** Resolves the entered text, or `undefined` on Esc, Cancel or an outside click. */
export function promptPopover(options: PromptOptions): Promise<string | undefined> {
  if (options.signal?.aborted) return Promise.resolve(undefined);
  ensureStyle();
  const previous = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  const pop = document.createElement('div');
  pop.className = 'dk-popover docxkeun-ui-root';
  pop.setAttribute('role', 'dialog');
  pop.setAttribute('aria-label', options.title);
  const title = document.createElement('div');
  title.className = 'dk-popover-title';
  title.textContent = options.title;
  const input = document.createElement('input');
  input.type = 'text';
  input.value = options.value ?? '';
  if (options.placeholder) input.placeholder = options.placeholder;
  input.setAttribute('aria-label', options.title);
  const error = document.createElement('div');
  error.className = 'dk-popover-error';
  error.setAttribute('role', 'alert');
  const actions = document.createElement('div');
  actions.className = 'dk-popover-actions';
  const cancel = document.createElement('button');
  cancel.type = 'button';
  cancel.textContent = 'Cancel';
  cancel.dataset.role = 'cancel';
  const ok = document.createElement('button');
  ok.type = 'button';
  ok.textContent = 'OK';
  ok.dataset.role = 'ok';
  actions.append(cancel, ok);
  pop.append(title, input, error, actions);
  pop.style.left = '0px';
  pop.style.top = '0px';
  (options.container ?? document.body).appendChild(pop);
  const rect = pop.getBoundingClientRect();
  const p = positionMenu(options.at, { w: rect.width, h: rect.height }, { w: window.innerWidth, h: window.innerHeight });
  pop.style.left = `${p.x}px`;
  pop.style.top = `${p.y}px`;

  return new Promise((resolve) => {
    const finish = (value: string | undefined): void => {
      document.removeEventListener('mousedown', onOutside, true);
      options.signal?.removeEventListener('abort', onAbort);
      pop.remove();
      if (previous?.isConnected) previous.focus();
      resolve(value);
    };
    const submit = (): void => {
      const message = options.validate?.(input.value);
      if (message) {
        error.textContent = message;
        input.setAttribute('aria-invalid', 'true');
        input.focus();
        return;
      }
      finish(input.value);
    };
    const onOutside = (e: Event): void => {
      if (!pop.contains(e.target as Node)) finish(undefined);
    };
    const onAbort = (): void => finish(undefined);
    document.addEventListener('mousedown', onOutside, true);
    options.signal?.addEventListener('abort', onAbort);
    pop.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        finish(undefined);
      } else if (e.key === 'Enter' && e.target === input) {
        e.preventDefault();
        submit();
      }
    });
    cancel.addEventListener('click', () => finish(undefined));
    ok.addEventListener('click', submit);
    input.focus();
    input.select();
  });
}
