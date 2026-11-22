/**
 * Connecting to Claude through the chat's local relay (server/relay.mjs). The page hands the user's
 * key to the relay once; the relay checks it, keeps it in memory and adds it to every API call it
 * forwards. The page itself sends no key, and the browser stores none. Nothing here shows or logs it.
 */
import Anthropic, { type ClientOptions } from '@anthropic-ai/sdk';

/** Every request to the relay carries this, so the relay can tell its own page from other sites. */
export const RELAY_HEADER = { 'X-Docxkeun-Relay': '1' } as const;

function origin(): string {
  return globalThis.location?.origin ?? 'http://localhost';
}

/** A client whose requests go to the relay; the relay adds the key. */
export function relayClient(extra: Partial<ClientOptions> = {}): Anthropic {
  return new Anthropic({
    apiKey: 'held-by-the-local-relay',
    baseURL: new URL('/anthropic', origin()).href,
    // The SDK refuses to run in a browser unless asked; it talks only to the relay on this computer.
    dangerouslyAllowBrowser: true,
    maxRetries: 2,
    // No credential leaves the page: the relay adds the real key.
    defaultHeaders: { ...RELAY_HEADER, Authorization: null },
    ...extra,
  });
}

export interface RelayStatus {
  hasKey: boolean;
  /** The relay was started with ANTHROPIC_API_KEY. */
  fromEnv: boolean;
}

/** The relay answered a key request with an error (its own or the API's, passed through). */
export class KeyCheckError extends Error {
  override name = 'KeyCheckError';
  constructor(
    readonly status: number,
    readonly apiMessage: string | undefined,
  ) {
    super(apiMessage ?? `HTTP ${status}`);
  }
}

/** The relay is not there: the page was not opened through `pnpm dev:chat` or `pnpm start:chat`. */
export class NoRelayError extends Error {
  override name = 'NoRelayError';
}

type FetchLike = typeof globalThis.fetch;

async function relayFetch(path: string, init: RequestInit, fetchImpl: FetchLike): Promise<Response> {
  try {
    return await fetchImpl(new URL(path, origin()).href, { ...init, headers: { ...RELAY_HEADER, ...(init.headers as Record<string, string> | undefined) } });
  } catch {
    throw new NoRelayError('The docxkeun chat server is not running.');
  }
}

async function errorFrom(res: Response): Promise<KeyCheckError> {
  let message: string | undefined;
  try {
    const body = (await res.json()) as { error?: { message?: unknown } };
    if (typeof body?.error?.message === 'string') message = body.error.message;
  } catch {
    message = undefined;
  }
  return new KeyCheckError(res.status, message);
}

export async function relayStatus(fetchImpl: FetchLike = fetch): Promise<RelayStatus> {
  const res = await relayFetch('/relay/status', { method: 'GET' }, fetchImpl);
  if (!res.ok || !(res.headers.get('content-type') ?? '').includes('application/json')) throw new NoRelayError('No relay at this address.');
  return (await res.json()) as RelayStatus;
}

/** Give the relay a key; it checks it with the API first and keeps it only if the API accepts it. */
export async function giveKey(key: string, model: string, fetchImpl: FetchLike = fetch): Promise<void> {
  const res = await relayFetch('/relay/key', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ key, model }) }, fetchImpl);
  if (!res.ok) throw await errorFrom(res);
}

export async function forgetKey(fetchImpl: FetchLike = fetch): Promise<void> {
  const res = await relayFetch('/relay/key', { method: 'DELETE' }, fetchImpl);
  if (!res.ok) throw await errorFrom(res);
}

/** The API's own explanation from an SDK error, if it gave one. */
export function apiMessage(err: unknown): string | undefined {
  if (!(err instanceof Anthropic.APIError)) return undefined;
  const body = err.error as { error?: { message?: unknown }; message?: unknown } | undefined;
  const message = body?.error?.message ?? body?.message;
  return typeof message === 'string' && message.trim() ? message.trim() : undefined;
}

/** What the key looks like it is, when that explains a rejection. */
export function keyHint(key: string): string | undefined {
  if (key.startsWith('sk-ant-oat')) return 'This looks like a Claude subscription token (from Claude Code or claude.ai), not an API key. Create an API key in the Claude Console.';
  if (key.startsWith('sk-ant-admin')) return 'This is an Admin API key. The chat needs a regular API key from the Claude Console.';
  if (!key.startsWith('sk-ant-')) return 'Anthropic API keys start with "sk-ant-".';
  if (/\s/.test(key)) return 'The key contains spaces or line breaks; copy it again in one piece.';
  return undefined;
}

export const START_THE_SERVER =
  'This page needs the docxkeun chat server on your computer. Start it with `pnpm dev:chat` (or `pnpm start:chat` after `pnpm build:chat`) and open the address it prints.';

/** A message for the key dialog (never containing the key). */
export function keyProblem(err: unknown, key: string): string {
  if (err instanceof NoRelayError) return START_THE_SERVER;
  if (err instanceof KeyCheckError) {
    const said = err.apiMessage ? ` (the API said: “${err.apiMessage}”)` : '';
    switch (err.status) {
      case 401:
        return `The API did not accept this key${said}. ${keyHint(key) ?? 'Check that you copied all of it, and that it has not been disabled in the Claude Console.'}`;
      case 403:
        return `This key may not be used here${said}.`;
      case 404:
        return 'The selected model is not available to this key. Pick another model and try again.';
      case 502:
        return `The chat server could not reach the Claude API${err.apiMessage ? `: ${err.apiMessage}` : ''}. Check your connection.`;
      default:
        return `Could not check the key (HTTP ${err.status})${said}.`;
    }
  }
  return `Could not check the key: ${(err as Error).message}`;
}
