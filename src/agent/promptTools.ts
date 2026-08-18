import type { Msg, ToolCall } from "../types";
import type { ToolDef } from "./tools";

/**
 * For models without function calling: the tool list goes in the system prompt and calls
 * come back as a JSON line in the reply.
 */
export function toolCatalogPrompt(tools: ToolDef[]): string {
  const lines = tools.map(
    (t) => "- " + t.name + "(" + Object.keys(t.parameters.properties).join(", ") + "): " + t.description
  );
  return [
    "",
    "## Calling tools",
    "This model has no native tool calling, so use the text protocol below.",
    "To call a tool, reply with ONLY a JSON object on its own, no prose around it:",
    '{"tool": "read_range", "args": {"sheet": "TB", "address": "A1:F120"}}',
    "You may call one tool per reply. The result comes back as a message starting with TOOL RESULT.",
    "When you are done and want to speak to the user, reply with plain text and no JSON.",
    "",
    "Available tools:",
    ...lines,
  ].join("\n");
}

const FENCE = /```(?:json)?\s*([\s\S]*?)```/i;

/** Pull a tool call out of a text reply. Returns the call plus whatever prose was around it. */
export function parsePromptedCall(text: string): { call: ToolCall | null; prose: string } {
  if (!text || !text.includes("{")) return { call: null, prose: text };

  const fenced = text.match(FENCE);
  const candidates: string[] = [];
  if (fenced) candidates.push(fenced[1].trim());

  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");
  if (start >= 0 && end > start) candidates.push(text.slice(start, end + 1));

  for (const raw of candidates) {
    try {
      const obj = JSON.parse(raw);
      const name = obj.tool ?? obj.name ?? obj.function;
      if (typeof name !== "string") continue;
      const args = obj.args ?? obj.arguments ?? obj.parameters ?? {};
      return {
        call: {
          id: "pcall_" + Math.random().toString(36).slice(2, 8),
          name,
          args: typeof args === "string" ? JSON.parse(args) : args,
        },
        prose: text.replace(fenced ? fenced[0] : raw, "").trim(),
      };
    } catch {
      // not JSON, keep looking
    }
  }
  return { call: null, prose: text };
}

/** Collapse tool roles into plain turns, for models that only know user/assistant. */
export function flattenForPromptMode(msgs: Msg[]): Msg[] {
  return msgs.map((m) => {
    if (m.role === "tool") {
      return {
        ...m,
        role: "user" as const,
        content: "TOOL RESULT " + (m.toolName ?? "") + ":\n" + m.content,
        toolCallId: undefined,
      };
    }
    if (m.role === "assistant" && m.toolCalls?.length) {
      const call = m.toolCalls[0];
      return {
        ...m,
        content:
          (m.content ? m.content + "\n" : "") +
          JSON.stringify({ tool: call.name, args: call.args }),
        toolCalls: undefined,
      };
    }
    return m;
  });
}
