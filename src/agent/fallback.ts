import type { Settings } from "../store/settings";
import { PROVIDERS, getProvider, type ProviderId } from "../providers/registry";
import type { ProviderError } from "../providers/error";

export interface Candidate {
  provider: ProviderId;
  model: string;
  label: string;
}

/** Fallback order: the user's chain, then other keyed providers, then other models here. */
export function candidatesFor(settings: Settings, exclude: { provider: ProviderId; model: string }): Candidate[] {
  const out: Candidate[] = [];
  const seen = new Set<string>([exclude.provider + "|" + exclude.model]);

  const add = (provider: ProviderId, model: string) => {
    const id = provider + "|" + model;
    if (seen.has(id) || !model) return;
    if (!settings.keys[provider]) return;
    seen.add(id);
    out.push({ provider, model, label: getProvider(provider).label + " · " + model });
  };

  for (const f of settings.fallbacks) add(f.provider, f.model);

  for (const p of PROVIDERS) {
    if (p.id === exclude.provider) continue;
    const cached = settings.modelCache[p.id];
    add(p.id, (cached ?? p.models)[0]);
  }

  const own = settings.modelCache[exclude.provider] ?? getProvider(exclude.provider).models;
  for (const m of own.slice(0, 3)) add(exclude.provider, m);

  return out;
}

export function failureLine(err: ProviderError): string {
  const who = getProvider(err.provider).label + " · " + err.model;
  switch (err.kind) {
    case "rate_limit":
      return (
        who +
        " has run out for now" +
        (err.retryAfterSec ? ", it resets in about " + err.retryAfterSec + "s" : "") +
        "."
      );
    case "auth":
      return who + " would not accept that API key.";
    case "model":
      return who + " is not on this key. The model was probably retired.";
    case "server":
      return who + " is having trouble at their end.";
    default:
      return who + " failed: " + err.message;
  }
}

/** Match a free-typed or clicked answer back to a candidate. */
export function matchChoice(answer: string, candidates: Candidate[]): Candidate | null {
  const a = answer.trim().toLowerCase();
  if (!a) return null;
  const byLabel = candidates.find((c) => c.label.toLowerCase() === a);
  if (byLabel) return byLabel;
  const byModel = candidates.find((c) => a.includes(c.model.toLowerCase()));
  if (byModel) return byModel;
  const byProvider = candidates.find((c) => a.includes(getProvider(c.provider).label.toLowerCase()));
  if (byProvider) return byProvider;
  const n = parseInt(a, 10);
  if (Number.isFinite(n) && n >= 1 && n <= candidates.length) return candidates[n - 1];
  return null;
}
