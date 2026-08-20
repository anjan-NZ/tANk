import type { Provider } from "./registry";

/** Providers retire model ids often, so the pane can ask them what exists right now. */
const NOT_CHAT = /whisper|tts|embed|guard|orpheus|allam|rerank|vision-ocr|image/i;

/** Advertises chat, does not reliably do a conversation with tools. */
const UNPROVEN = /preview|experimental|-exp|antigravity|learnlm|nightly|alpha|beta|deprecated/i;

/**
 * Which live id to move to when the configured one has been retired. The list a provider
 * returns is alphabetical, and its first entry is as likely to be a half-built preview as
 * anything usable, so prefer the ids this project already vouches for.
 */
export function autoPick(
  curated: string[],
  live: string[],
  avoid: ReadonlySet<string>
): string | undefined {
  const usable = live.filter((id) => !avoid.has(id));
  return (
    usable.find((id) => curated.includes(id)) ??
    usable.find((id) => !UNPROVEN.test(id)) ??
    usable[0]
  );
}

export async function fetchModels(provider: Provider, apiKey: string): Promise<string[]> {
  if (!apiKey) throw new Error("Add an API key first.");

  if (provider.kind === "gemini") {
    const res = await fetch(provider.baseUrl + "/models", {
      headers: { "x-goog-api-key": apiKey },
    });
    if (!res.ok) throw new Error("Gemini " + res.status + ": " + (await res.text()).slice(0, 200));
    const json = await res.json();
    return (json.models ?? [])
      .filter((m: any) => (m.supportedGenerationMethods ?? []).includes("generateContent"))
      .map((m: any) => String(m.name ?? "").replace(/^models\//, ""))
      .filter((id: string) => id && !NOT_CHAT.test(id))
      .sort();
  }

  const res = await fetch(provider.baseUrl + "/models", {
    headers: { Authorization: "Bearer " + apiKey },
  });
  if (!res.ok)
    throw new Error(provider.label + " " + res.status + ": " + (await res.text()).slice(0, 200));
  const json = await res.json();
  return (json.data ?? [])
    .map((m: any) => String(m.id ?? ""))
    .filter((id: string) => id && !NOT_CHAT.test(id))
    .sort();
}
