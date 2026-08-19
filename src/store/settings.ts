import type { CavemanLevel } from "../types";
import type { ProviderId } from "../providers/registry";

const KEY = "tank.settings.v1"; // renamed from officeai.settings.v2

export interface Settings {
  provider: ProviderId;
  model: string;
  caveman: CavemanLevel;
  /** provider id -> api key. Never leaves this machine except to that provider. */
  keys: Partial<Record<ProviderId, string>>;
  maxRowsPerRead: number;
  /** ask before overwriting cells that already have something in them */
  editMode: "ask" | "auto";
  /** model ids last fetched from each provider, so the picker stays current */
  modelCache: Partial<Record<ProviderId, string[]>>;
  /** preferred order to fall back through when a model is rate limited */
  fallbacks: Array<{ provider: ProviderId; model: string }>;
  /** switch on a rate limit without asking first */
  autoRotate: boolean;
  /** "provider|model" entries that need the prompt-based tool protocol */
  promptToolModels: string[];
  theme: "system" | "light" | "dark";
  /** drop the narrow convenience tools and lean on run_office_script instead */
  leanTools: boolean;
  /** look for a new release now and then. It only looks; nothing installs on its own. */
  autoCheckUpdates: boolean;
  /** show the tool calls and their raw results in the transcript */
  showToolDetails: boolean;
  /** show the model switching chatter (rate limits, rotations) in the transcript */
  showNotices: boolean;
}

const DEFAULTS: Settings = {
  provider: "groq",
  model: "openai/gpt-oss-120b",
  caveman: "lite",
  keys: {},
  maxRowsPerRead: 400,
  editMode: "ask",
  modelCache: {},
  fallbacks: [],
  autoRotate: false,
  promptToolModels: [],
  theme: "system",
  leanTools: false,
  autoCheckUpdates: true,
  showToolDetails: false,
  showNotices: false,
};

/**
 * Local convenience only: keys put in .env.local (gitignored) seed the settings on
 * first run so there is nothing to paste. A key typed in Settings always wins.
 */
function seededKeys(): Settings["keys"] {
  const env = import.meta.env as Record<string, string | undefined>;
  const seeds: Settings["keys"] = {};
  // Never bake a key into a production bundle.
  if (!import.meta.env.DEV) return seeds;
  if (env.VITE_GROQ_API_KEY) seeds.groq = env.VITE_GROQ_API_KEY;
  if (env.VITE_GEMINI_API_KEY) seeds.gemini = env.VITE_GEMINI_API_KEY;
  if (env.VITE_OPENROUTER_API_KEY) seeds.openrouter = env.VITE_OPENROUTER_API_KEY;
  if (env.VITE_CEREBRAS_API_KEY) seeds.cerebras = env.VITE_CEREBRAS_API_KEY;
  return seeds;
}

export function loadSettings(): Settings {
  const base: Settings = { ...DEFAULTS, keys: seededKeys() };
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return base;
    const stored = JSON.parse(raw) as Partial<Settings>;
    return { ...base, ...stored, keys: { ...base.keys, ...(stored.keys ?? {}) } };
  } catch {
    return base;
  }
}

export function saveSettings(s: Settings): void {
  localStorage.setItem(KEY, JSON.stringify(s));
}

/** Wipe every key and setting this add-in has stored on the machine. */
export function clearAllSettings(): void {
  localStorage.removeItem(KEY);
}
