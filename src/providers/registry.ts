export type ProviderId =
  | "groq"
  | "gemini"
  | "openrouter"
  | "cerebras"
  | "mistral"
  | "sambanova"
  | "nvidia"
  | "github"
  | "zai";
export type ProviderKind = "openai" | "gemini";

export interface Provider {
  id: ProviderId;
  label: string;
  kind: ProviderKind;
  baseUrl: string;
  keyUrl: string;
  models: string[];
  /** free tier in one line, shown in the Help table */
  freeLimit: string;
  note?: string;
}

/** Free-tier friendly providers. All keys stay in this machine's localStorage. */
export const PROVIDERS: Provider[] = [
  {
    id: "groq",
    label: "Groq",
    kind: "openai",
    baseUrl: "https://api.groq.com/openai/v1",
    keyUrl: "https://console.groq.com/keys",
    models: ["openai/gpt-oss-120b", "openai/gpt-oss-20b", "qwen/qwen3.6-27b", "groq/compound"],
    freeLimit: "30 req/min, 1,000 req/day, no card",
    note: "Fastest free tier. gpt-oss-120b is the strongest tool caller here.",
  },
  {
    id: "gemini",
    label: "Google Gemini",
    kind: "gemini",
    baseUrl: "https://generativelanguage.googleapis.com/v1beta",
    keyUrl: "https://aistudio.google.com/apikey",
    models: ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.0-flash"],
    freeLimit: "15 req/min, 1,500 req/day",
    note: "Big free quota, strong on long sheets.",
  },
  {
    id: "openrouter",
    label: "OpenRouter",
    kind: "openai",
    baseUrl: "https://openrouter.ai/api/v1",
    keyUrl: "https://openrouter.ai/keys",
    models: [
      "deepseek/deepseek-chat-v3.1:free",
      "qwen/qwen3-235b-a22b:free",
      "meta-llama/llama-3.3-70b-instruct:free",
    ],
    freeLimit: "20 req/min, 50 req/day on :free models",
    note: "Router over many models; :free variants are rate limited.",
  },
  {
    id: "cerebras",
    label: "Cerebras",
    kind: "openai",
    baseUrl: "https://api.cerebras.ai/v1",
    keyUrl: "https://cloud.cerebras.ai",
    models: ["llama-3.3-70b", "qwen-3-32b"],
    freeLimit: "1M tokens/day, card on file required",
    note: "Very fast, small free daily quota.",
  },
  {
    id: "mistral",
    label: "Mistral",
    kind: "openai",
    baseUrl: "https://api.mistral.ai/v1",
    keyUrl: "https://console.mistral.ai/api-keys",
    models: ["mistral-medium-latest", "mistral-small-latest", "open-mistral-nemo"],
    freeLimit: "~1B tokens/month on the Experiment plan, no card",
    note: "Experiment plan, no card. ~1B tokens/month. Reliable tool calling.",
  },
  {
    id: "sambanova",
    label: "SambaNova",
    kind: "openai",
    baseUrl: "https://api.sambanova.ai/v1",
    keyUrl: "https://cloud.sambanova.ai/apis",
    models: ["DeepSeek-V3.1", "Meta-Llama-3.3-70B-Instruct", "Llama-4-Maverick-17B-128E-Instruct"],
    freeLimit: "20 req/min, 200K tokens/day, no card",
    note: "No card. 20 RPM, 200K tokens/day. Fast big models.",
  },
  {
    id: "nvidia",
    label: "NVIDIA NIM",
    kind: "openai",
    baseUrl: "https://integrate.api.nvidia.com/v1",
    keyUrl: "https://build.nvidia.com",
    models: [
      "meta/llama-3.3-70b-instruct",
      "deepseek-ai/deepseek-v3.1",
      "qwen/qwen3-235b-a22b",
    ],
    freeLimit: "~40 req/min, free developer program",
    note: "Free with the developer program, ~40 RPM. Wide model catalogue.",
  },
  {
    id: "github",
    label: "GitHub Models",
    kind: "openai",
    baseUrl: "https://models.github.ai/inference",
    keyUrl: "https://github.com/settings/tokens",
    models: ["openai/gpt-4o", "openai/gpt-4o-mini", "meta/Llama-3.3-70B-Instruct"],
    freeLimit: "10 req/min, 50 req/day, key is a GitHub PAT",
    note: "Key is a GitHub PAT with models:read. 10 RPM, 50 requests/day.",
  },
  {
    id: "zai",
    label: "Z.AI (GLM)",
    kind: "openai",
    baseUrl: "https://open.bigmodel.cn/api/paas/v4",
    keyUrl: "https://open.bigmodel.cn/usercenter/apikeys",
    models: ["glm-4.7-flash", "glm-4.6"],
    freeLimit: "Flash tier free, 1 request at a time",
    note: "Flash tier is free, one request at a time. Hosted in China, expect latency.",
  },
];

export function getProvider(id: ProviderId): Provider {
  const p = PROVIDERS.find((x) => x.id === id);
  if (!p) throw new Error(`Unknown provider: ${id}`);
  return p;
}
