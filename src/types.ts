export type Role = "system" | "user" | "assistant" | "tool";

export interface ToolCall {
  id: string;
  name: string;
  args: Record<string, unknown>;
}

export interface Msg {
  id: string;
  role: Role;
  content: string;
  toolCalls?: ToolCall[];
  /** set on role==="tool": which call this answers */
  toolCallId?: string;
  toolName?: string;
  /** UI-only: transient status line under an assistant turn */
  status?: string;
  /** shown in the transcript, never replayed to the model */
  uiOnly?: boolean;
  /** routing chatter: which model was swapped in and why. Hidden unless asked for. */
  notice?: boolean;
  error?: boolean;
}

export interface RangeRef {
  sheet: string;
  address: string;
}

export type ScopeSource = "selection" | "user" | "auto" | "workbook";

export interface Scope {
  ranges: RangeRef[];
  label: string;
  source: ScopeSource;
}

export type CavemanLevel = "off" | "lite" | "ultra";

export interface AskRequest {
  question: string;
  options: string[];
}

export interface UndoEntry {
  label: string;
  at: number;
  cells: Array<{ sheet: string; address: string; formulas: unknown[][]; numberFormat: string[][] }>;
}
