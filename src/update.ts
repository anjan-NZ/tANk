declare const __APP_VERSION__: string;
declare const __REPO_SLUG__: string;

export const APP_VERSION = typeof __APP_VERSION__ === "string" ? __APP_VERSION__ : "dev";
export const REPO_SLUG = typeof __REPO_SLUG__ === "string" ? __REPO_SLUG__ : "";

export interface UpdateCheck {
  current: string;
  latest?: string;
  newer: boolean;
  url?: string;
  notes?: string;
  error?: string;
  checkedAt: number;
}

/** Compares two version strings like 1.2.10 without pulling in a library. */
function isNewer(latest: string, current: string): boolean {
  const parts = (v: string) =>
    v.replace(/^v/, "").split(/[.\-+]/).map((p) => (/^\d+$/.test(p) ? Number(p) : 0));
  const a = parts(latest);
  const b = parts(current);
  for (let i = 0; i < Math.max(a.length, b.length); i++) {
    const x = a[i] ?? 0;
    const y = b[i] ?? 0;
    if (x !== y) return x > y;
  }
  return false;
}

/**
 * Asks GitHub for the newest release. The pane itself is loaded from the web, so a new
 * version of the chat arrives on its own; only a change to the manifest or the installer
 * needs the user to run the setup again.
 */
export async function checkForUpdate(signal?: AbortSignal): Promise<UpdateCheck> {
  const now = Date.now();
  const base: UpdateCheck = { current: APP_VERSION, newer: false, checkedAt: now };

  if (!REPO_SLUG) return { ...base, error: "This build does not know where to look for updates." };

  try {
    const res = await fetch("https://api.github.com/repos/" + REPO_SLUG + "/releases/latest", {
      headers: { Accept: "application/vnd.github+json" },
      signal,
    });
    if (res.status === 404) return { ...base, error: "No releases published yet." };
    if (!res.ok) return { ...base, error: "GitHub replied " + res.status + "." };

    const json = await res.json();
    const latest = String(json.tag_name ?? "").replace(/^v/, "");
    if (!latest) return { ...base, error: "That release has no version number." };

    return {
      ...base,
      latest,
      newer: isNewer(latest, APP_VERSION),
      url: json.html_url,
      notes: typeof json.body === "string" ? json.body.slice(0, 400) : undefined,
    };
  } catch (err) {
    if (signal?.aborted) return base;
    return { ...base, error: err instanceof Error ? err.message : String(err) };
  }
}
