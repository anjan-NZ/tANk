import type { ProviderId } from "../providers/registry";

/**
 * Tracks what each model has left, taken from the rate limit headers providers return,
 * so the loop can switch before a 429 instead of after one.
 */
export interface RateInfo {
  remainingRequests?: number;
  remainingTokens?: number;
  /** seconds until the window resets */
  resetSeconds?: number;
}

interface Health {
  cooldownUntil: number;
  remainingRequests?: number;
  remainingTokens?: number;
  lastUsed: number;
  fails: number;
  calls: number;
  tokens: number;
  limitHits: number;
}

const health = new Map<string, Health>();

export const idOf = (provider: ProviderId, model: string) => provider + "|" + model;

function entry(id: string): Health {
  const h =
    health.get(id) ?? { cooldownUntil: 0, lastUsed: 0, fails: 0, calls: 0, tokens: 0, limitHits: 0 };
  health.set(id, h);
  return h;
}

export function noteSuccess(id: string, rate?: RateInfo, tokens?: number): void {
  const h = entry(id);
  h.lastUsed = Date.now();
  h.fails = 0;
  h.cooldownUntil = 0;
  h.calls++;
  if (tokens) h.tokens += tokens;
  if (rate?.remainingRequests !== undefined) h.remainingRequests = rate.remainingRequests;
  if (rate?.remainingTokens !== undefined) h.remainingTokens = rate.remainingTokens;
}

export function noteRateLimit(id: string, resetSeconds: number | null): void {
  const h = entry(id);
  h.fails++;
  h.limitHits++;
  // Trust the provider's own reset when it gives one, otherwise back off progressively.
  const wait = resetSeconds ?? Math.min(60, 5 * h.fails);
  h.cooldownUntil = Date.now() + wait * 1000;
  h.remainingRequests = 0;
  h.remainingTokens = 0;
}

export function noteFailure(id: string): void {
  const h = entry(id);
  h.fails++;
  h.cooldownUntil = Date.now() + Math.min(30, 5 * h.fails) * 1000;
}

export function coolingFor(id: string): number {
  const h = health.get(id);
  if (!h) return 0;
  return Math.max(0, Math.ceil((h.cooldownUntil - Date.now()) / 1000));
}

/** Providers cap requests and tokens per minute, and one turn costs a few thousand tokens. */
export function runningLow(id: string, estimatedTokens: number): boolean {
  const h = health.get(id);
  if (!h) return false;
  if (h.remainingRequests !== undefined && h.remainingRequests <= 1) return true;
  if (h.remainingTokens !== undefined && h.remainingTokens < estimatedTokens) return true;
  return false;
}

export function statusLine(id: string): string {
  const h = health.get(id);
  if (!h) return "";
  const bits: string[] = [];
  if (h.remainingRequests !== undefined) bits.push(h.remainingRequests + " requests left");
  if (h.remainingTokens !== undefined) bits.push(h.remainingTokens.toLocaleString() + " tokens left");
  const cool = coolingFor(id);
  if (cool) bits.push("resting " + cool + "s");
  return bits.join(", ");
}

/** Least recently used model that is not cooling down and has room for this request. */
export function pickHealthiest<T extends { provider: ProviderId; model: string }>(
  candidates: T[],
  estimatedTokens: number
): T | null {
  const usable = candidates.filter((c) => {
    const id = idOf(c.provider, c.model);
    return !coolingFor(id) && !runningLow(id, estimatedTokens);
  });
  if (!usable.length) return null;
  return usable.sort((a, b) => {
    const ha = health.get(idOf(a.provider, a.model));
    const hb = health.get(idOf(b.provider, b.model));
    return (ha?.lastUsed ?? 0) - (hb?.lastUsed ?? 0);
  })[0];
}

export interface UsageRow {
  id: string;
  provider: string;
  model: string;
  calls: number;
  tokens: number;
  limitHits: number;
  remainingRequests?: number;
  remainingTokens?: number;
  cooling: number;
  lastUsed: number;
}

/** Per model usage for the stats panel. */
export function snapshot(): UsageRow[] {
  return [...health.entries()]
    .map(([id, h]) => {
      const [provider, model] = id.split("|");
      return {
        id,
        provider,
        model,
        calls: h.calls,
        tokens: h.tokens,
        limitHits: h.limitHits,
        remainingRequests: h.remainingRequests,
        remainingTokens: h.remainingTokens,
        cooling: coolingFor(id),
        lastUsed: h.lastUsed,
      };
    })
    .sort((a, b) => b.lastUsed - a.lastUsed);
}

export function totals(): { calls: number; tokens: number; limitHits: number } {
  return [...health.values()].reduce(
    (t, h) => ({ calls: t.calls + h.calls, tokens: t.tokens + h.tokens, limitHits: t.limitHits + h.limitHits }),
    { calls: 0, tokens: 0, limitHits: 0 }
  );
}

/** Approximate request size. Only used to decide whether a model has room. */
export function estimateTokens(systemPrompt: string, historyChars: number, toolChars: number): number {
  return Math.round((systemPrompt.length + historyChars + toolChars) / 3.7) + 600;
}
