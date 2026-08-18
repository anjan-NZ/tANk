import type { Msg, ToolCall } from "../types";
import type { ToolDef } from "../agent/tools";
import type { ChatRequest, ChatResponse } from "./index";
import { ProviderError, classify, rateFrom, redact, retryAfterFrom } from "./error";

interface OaiMsg {
  role: string;
  content: string | null;
  tool_calls?: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }>;
  tool_call_id?: string;
}

function toOai(msgs: Msg[]): OaiMsg[] {
  return msgs.map((m) => {
    if (m.role === "tool") {
      return { role: "tool", content: m.content, tool_call_id: m.toolCallId };
    }
    if (m.role === "assistant" && m.toolCalls?.length) {
      return {
        role: "assistant",
        content: m.content || null,
        tool_calls: m.toolCalls.map((t) => ({
          id: t.id,
          type: "function" as const,
          function: { name: t.name, arguments: JSON.stringify(t.args ?? {}) },
        })),
      };
    }
    return { role: m.role, content: m.content };
  });
}

function toOaiTools(tools: ToolDef[]) {
  return tools.map((t) => ({
    type: "function",
    function: { name: t.name, description: t.description, parameters: t.parameters },
  }));
}

export async function chatOpenAi(req: ChatRequest): Promise<ChatResponse> {
  const body = {
    model: req.model,
    messages: [{ role: "system", content: req.system }, ...toOai(req.msgs)],
    ...(req.promptTools ? {} : { tools: toOaiTools(req.tools), tool_choice: "auto" }),
    temperature: 0.2,
  };

  let res: Response;
  try {
    res = await fetch(req.provider.baseUrl + "/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + req.apiKey,
        ...(req.provider.id === "openrouter"
          ? { "HTTP-Referer": "https://localhost:3000", "X-Title": "tANk" }
          : {}),
      },
      body: JSON.stringify(body),
      signal: req.signal,
    });
  } catch (err) {
    if (req.signal?.aborted) throw err;
    throw new ProviderError({
      message: req.provider.label + " unreachable: " + (err instanceof Error ? err.message : String(err)),
      kind: "network",
      status: 0,
      provider: req.provider.id,
      model: req.model,
    });
  }

  if (!res.ok) {
    const text = redact(await res.text().catch(() => ""), req.apiKey);
    throw new ProviderError({
      message: `${req.provider.label} ${res.status}: ${text.slice(0, 300)}`,
      kind: classify(res.status, text),
      status: res.status,
      provider: req.provider.id,
      model: req.model,
      retryAfterSec: retryAfterFrom(res.headers, text),
    });
  }

  const json = await res.json();
  const choice = json.choices?.[0]?.message ?? {};
  const toolCalls: ToolCall[] = (choice.tool_calls ?? []).map((t: any, i: number) => ({
    id: t.id ?? "call_" + i,
    name: t.function?.name ?? "",
    args: safeParse(t.function?.arguments),
  }));

  return {
    content: dedupe(choice.content ?? ""),
    toolCalls,
    rate: rateFrom(res.headers),
    usage: {
      promptTokens: json.usage?.prompt_tokens,
      completionTokens: json.usage?.completion_tokens,
      totalTokens: json.usage?.total_tokens,
    },
  };
}

/** gpt-oss and a few others sometimes emit the same answer twice in one message. */
function dedupe(text: string): string {
  const t = (text ?? "").trim();
  if (t.length < 40) return t;
  const half = Math.floor(t.length / 2);
  const a = t.slice(0, half).trim();
  const b = t.slice(half).trim();
  return a && a === b ? a : t;
}

function safeParse(s: unknown): Record<string, unknown> {
  if (typeof s !== "string" || !s.trim()) return {};
  try {
    return JSON.parse(s);
  } catch {
    // Some small models emit single quotes or trailing commas; try one cheap repair.
    try {
      return JSON.parse(s.replace(/,\s*([}\]])/g, "$1"));
    } catch {
      return {};
    }
  }
}
