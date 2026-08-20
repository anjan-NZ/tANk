import type { ProviderId } from "./registry";

export type FailureKind =
  | "rate_limit" // 429, or quota exhausted for the day
  | "auth" // missing or rejected key
  | "model" // model id retired or not available on this key
  | "server" // provider side 5xx
  | "network" // fetch could not reach the provider (CORS, offline)
  | "tools_unsupported" // model or endpoint refuses function calling
  | "other";

export class ProviderError extends Error {
  readonly kind: FailureKind;
  readonly status: number;
  readonly provider: ProviderId;
  readonly model: string;
  /** seconds the provider asked us to wait, when it said so */
  readonly retryAfterSec: number | null;

  constructor(init: {
    message: string;
    kind: FailureKind;
    status: number;
    provider: ProviderId;
    model: string;
    retryAfterSec?: number | null;
  }) {
    super(init.message);
    this.name = "ProviderError";
    this.kind = init.kind;
    this.status = init.status;
    this.provider = init.provider;
    this.model = init.model;
    this.retryAfterSec = init.retryAfterSec ?? null;
  }

  /** true when trying a different model or provider could plausibly work */
  get switchable(): boolean {
    return this.kind === "rate_limit" || this.kind === "model" || this.kind === "server" || this.kind === "auth";
  }
}

/** Defensive: provider errors sometimes echo the request, so scrub the key before it is shown. */
export function redact(text: string, apiKey: string): string {
  if (!apiKey) return text;
  const tail = apiKey.slice(-4);
  return text.split(apiKey).join("****" + tail);
}

export function classify(status: number, body: string): FailureKind {
  // Before the tools check: this one names the model, so no amount of retrying without
  // tools will help. Gemini says it about previews that are not open for general use.
  if (/not enabled for models\/|is not supported for|multiturn/i.test(body)) return "model";

  // A 400 about tools means "retry me without tools", not "give up".
  if (
    (status === 400 || status === 422 || status === 404) &&
    /tool|function[_ ]?call|tool_choice/i.test(body) &&
    /not support|unsupported|invalid|unrecognized|unknown field|does not/i.test(body)
  )
    return "tools_unsupported";
  if (status === 429) return "rate_limit";
  if (status === 401 || status === 403) return "auth";
  if (status === 402) return "rate_limit"; // credits exhausted behaves like a quota wall
  if (
    status === 404 ||
    /model_not_found|does not exist|decommissioned|multiturn|not enabled for|not supported for/i.test(
      body
    )
  )
    return "model";
  if (status >= 500) return "server";
  if (/rate limit|quota|exhausted|resource_exhausted/i.test(body)) return "rate_limit";
  return "other";
}

/** These come back as durations, not plain seconds: "60", "980ms", "2m59.56s". */
function durationSecs(v: string | null): number | undefined {
  if (!v) return undefined;
  const parts = [...v.matchAll(/([\d.]+)\s*(ms|s|m|h)?/g)];
  let total = 0;
  let found = false;
  for (const p of parts) {
    const f = parseFloat(p[1]);
    if (!Number.isFinite(f)) continue;
    found = true;
    total += p[2] === "h" ? f * 3600 : p[2] === "m" ? f * 60 : p[2] === "ms" ? f / 1000 : f;
  }
  if (found) return total;
  // Retry-After is allowed to be an HTTP date instead of a count of seconds.
  const at = Date.parse(v);
  return Number.isFinite(at) ? Math.max(0, (at - Date.now()) / 1000) : undefined;
}

/** Groq and friends report what is left in the current window on every response. */
export function rateFrom(headers: Headers): {
  remainingRequests?: number;
  remainingTokens?: number;
  resetSeconds?: number;
} {
  const n = (v: string | null) => {
    if (!v) return undefined;
    const f = parseFloat(v);
    return Number.isFinite(f) ? f : undefined;
  };
  return {
    remainingRequests: n(headers.get("x-ratelimit-remaining-requests")),
    remainingTokens: n(headers.get("x-ratelimit-remaining-tokens")),
    resetSeconds:
      durationSecs(headers.get("x-ratelimit-reset-tokens")) ??
      durationSecs(headers.get("x-ratelimit-reset-requests")),
  };
}

/** Providers report the wait in a header, or inside the message ("try again in 42.1s"). */
export function retryAfterFrom(headers: Headers, body: string): number | null {
  const h = durationSecs(headers.get("retry-after") ?? headers.get("x-ratelimit-reset-requests"));
  if (h !== undefined) return Math.ceil(h);

  const m = body.match(/try again in ([\d.]+\s*(?:h|m|s|ms)(?:[\d.]+\s*(?:m|s|ms))*)/i);
  const fromBody = durationSecs(m?.[1] ?? null);
  return fromBody === undefined ? null : Math.ceil(fromBody);
}
