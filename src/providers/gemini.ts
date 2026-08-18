import type { Msg, ToolCall } from "../types";
import type { JsonSchema, ToolDef } from "../agent/tools";
import type { ChatRequest, ChatResponse } from "./index";
import { ProviderError, classify, rateFrom, redact, retryAfterFrom } from "./error";

/** Gemini takes an OpenAPI subset: upper-case type names, every array needs typed items. */
function geminiSchema(node: any): any {
  if (!node || typeof node !== "object") return { type: "STRING" };
  const out: any = {};
  if (node.type) out.type = String(node.type).toUpperCase();
  if (node.description) out.description = node.description;
  if (node.enum) out.enum = node.enum;
  if (out.type === "OBJECT" || node.properties) {
    out.type = "OBJECT";
    out.properties = {};
    for (const [k, v] of Object.entries(node.properties ?? {})) out.properties[k] = geminiSchema(v);
    if (node.required?.length) out.required = node.required;
    // Gemini rejects an OBJECT with no properties, so give it a harmless one.
    if (!Object.keys(out.properties).length) {
      out.properties = { _unused: { type: "STRING" } };
    }
  }
  if (out.type === "ARRAY" || node.items) {
    out.type = "ARRAY";
    out.items = geminiSchema(node.items);
  }
  if (!out.type) out.type = "STRING";
  return out;
}

function toGeminiTools(tools: ToolDef[]) {
  return [
    {
      function_declarations: tools.map((t) => ({
        name: t.name,
        description: t.description,
        parameters: geminiSchema(t.parameters as JsonSchema),
      })),
    },
  ];
}

function toContents(msgs: Msg[]) {
  const out: any[] = [];
  for (const m of msgs) {
    if (m.role === "user") {
      out.push({ role: "user", parts: [{ text: m.content }] });
    } else if (m.role === "assistant") {
      const parts: any[] = [];
      if (m.content) parts.push({ text: m.content });
      for (const t of m.toolCalls ?? []) parts.push({ functionCall: { name: t.name, args: t.args ?? {} } });
      if (parts.length) out.push({ role: "model", parts });
    } else if (m.role === "tool") {
      out.push({
        role: "user",
        parts: [
          { functionResponse: { name: m.toolName ?? "tool", response: { result: m.content } } },
        ],
      });
    }
  }
  return out;
}

export async function chatGemini(req: ChatRequest): Promise<ChatResponse> {
  const url =
    req.provider.baseUrl + "/models/" + encodeURIComponent(req.model) + ":generateContent";

  const body = {
    system_instruction: { parts: [{ text: req.system }] },
    contents: toContents(req.msgs),
    ...(req.promptTools ? {} : { tools: toGeminiTools(req.tools) }),
    generationConfig: { temperature: 0.2 },
  };

  let res: Response;
  try {
    res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-goog-api-key": req.apiKey },
      body: JSON.stringify(body),
      signal: req.signal,
    });
  } catch (err) {
    if (req.signal?.aborted) throw err;
    throw new ProviderError({
      message: "Gemini unreachable: " + (err instanceof Error ? err.message : String(err)),
      kind: "network",
      status: 0,
      provider: req.provider.id,
      model: req.model,
    });
  }

  if (!res.ok) {
    const text = redact(await res.text().catch(() => ""), req.apiKey);
    throw new ProviderError({
      message: `Gemini ${res.status}: ${text.slice(0, 300)}`,
      kind: classify(res.status, text),
      status: res.status,
      provider: req.provider.id,
      model: req.model,
      retryAfterSec: retryAfterFrom(res.headers, text),
    });
  }

  const json = await res.json();
  const parts = json.candidates?.[0]?.content?.parts ?? [];
  let content = "";
  const toolCalls: ToolCall[] = [];
  parts.forEach((p: any, i: number) => {
    if (p.text) content += p.text;
    if (p.functionCall) {
      toolCalls.push({
        id: "gcall_" + i + "_" + Date.now(),
        name: p.functionCall.name,
        args: p.functionCall.args ?? {},
      });
    }
  });

  return {
    content,
    toolCalls,
    rate: rateFrom(res.headers),
    usage: {
      promptTokens: json.usageMetadata?.promptTokenCount,
      completionTokens: json.usageMetadata?.candidatesTokenCount,
      totalTokens: json.usageMetadata?.totalTokenCount,
    },
  };
}
