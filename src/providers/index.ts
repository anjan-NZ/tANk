import type { Msg, ToolCall } from "../types";
import type { ToolDef } from "../agent/tools";
import type { Provider } from "./registry";
import { chatOpenAi } from "./openaiCompat";
import { chatGemini } from "./gemini";

export interface ChatRequest {
  provider: Provider;
  apiKey: string;
  model: string;
  system: string;
  msgs: Msg[];
  tools: ToolDef[];
  /** true when the model has no native tool calling and the catalogue is in the prompt */
  promptTools?: boolean;
  signal?: AbortSignal;
}

export interface ChatResponse {
  content: string;
  toolCalls: ToolCall[];
  /** what the provider says is left in this minute's window */
  rate?: { remainingRequests?: number; remainingTokens?: number; resetSeconds?: number };
  /** what this call actually cost, as counted by the provider */
  usage?: { promptTokens?: number; completionTokens?: number; totalTokens?: number };
}

export function callModel(req: ChatRequest): Promise<ChatResponse> {
  if (!req.apiKey) throw new Error("No API key set for " + req.provider.label + ". Open Settings.");
  return req.provider.kind === "gemini" ? chatGemini(req) : chatOpenAi(req);
}
