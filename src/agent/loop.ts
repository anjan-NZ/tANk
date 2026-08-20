import type { Msg, Scope } from "../types";
import type { SelectionInfo } from "../office/scope";
import type { Settings } from "../store/settings";
import { getProvider, type ProviderId } from "../providers/registry";
import { callModel, type ChatResponse } from "../providers";
import { ProviderError } from "../providers/error";
import { candidatesFor, failureLine, matchChoice } from "./fallback";
import {
  coolingFor,
  estimateTokens,
  idOf,
  noteFailure,
  noteRateLimit,
  noteSuccess,
  pickHealthiest,
  runningLow,
  statusLine,
} from "./budget";
import { toolsFor, execTool } from "./tools";
import { buildSystemPrompt } from "./systemPrompt";
import { flattenForPromptMode, parsePromptedCall, toolCatalogPrompt } from "./promptTools";
import { autoPick, fetchModels } from "../providers/models";

const MAX_STEPS = 12;
const OLD_TOOL_RESULT_CHARS = 600;
const MAX_SWITCHES = 4;
const MAX_WAITS = 3;
const MAX_RELISTS = 2;

export interface TurnDeps {
  settings: Settings;
  getScope: () => Scope | null;
  getSelection: () => SelectionInfo | null;
  setScope: (s: Scope) => void;
  ask: (question: string, options: string[]) => Promise<string>;
  confirmEdit: (summary: string) => Promise<boolean>;
  push: (m: Msg) => void;
  setStatus: (s: string | null) => void;
  /** persist a model switch the user accepted */
  onSwitch: (provider: ProviderId, model: string) => void;
  /** remember that this model needs the prompt-based tool protocol */
  onPromptMode: (provider: ProviderId, model: string) => void;
  /** a provider just told us what it actually has; worth keeping */
  onModels: (provider: ProviderId, models: string[]) => void;
  signal?: AbortSignal;
}

export function newId(): string {
  return Math.random().toString(36).slice(2, 10);
}

/** Old tool results dominate the context, so only the last two are kept in full. */
function compact(msgs: Msg[]): Msg[] {
  const toolIdx = msgs.map((m, i) => (m.role === "tool" ? i : -1)).filter((i) => i >= 0);
  const keepFull = new Set(toolIdx.slice(-2));
  return msgs.map((m, i) => {
    if (m.role !== "tool" || keepFull.has(i) || m.content.length <= OLD_TOOL_RESULT_CHARS) return m;
    return {
      ...m,
      content: m.content.slice(0, OLD_TOOL_RESULT_CHARS) + "\n... [older result trimmed]",
    };
  });
}

const sleep = (ms: number, signal?: AbortSignal) =>
  new Promise<void>((resolve, reject) => {
    const t = setTimeout(resolve, ms);
    signal?.addEventListener("abort", () => {
      clearTimeout(t);
      reject(new DOMException("aborted", "AbortError"));
    });
  });

export async function runTurn(history: Msg[], deps: TurnDeps): Promise<void> {
  const { settings } = deps;
  const convo: Msg[] = [...history];
  let active = { provider: settings.provider, model: settings.model };
  let switches = 0;
  let waits = 0;
  const deadModels = new Set<string>();
  let unproven: { provider: ProviderId; model: string } | null = null;
  const relists = new Map<ProviderId, number>();
  const promptModels = new Set(settings.promptToolModels);
  const isPromptMode = () => promptModels.has(active.provider + "|" + active.model);

  /** One model call, switching models when the provider refuses. */
  async function callWithFallback(): Promise<ChatResponse | null> {
    for (;;) {
      const prompted = isPromptMode();

      // Step aside before the wall: if this model has no room left in its window,
      // move to the freshest one that does. No question asked, it is not a real failure.
      const activeId = idOf(active.provider, active.model);
      const cooling = coolingFor(activeId);
      const tools = toolsFor(undefined, settings.leanTools);
      const estimate = estimateTokens(
        buildSystemPrompt(settings.caveman, deps.getScope(), deps.getSelection()),
        convo.reduce((n, m) => n + m.content.length, 0),
        JSON.stringify(tools).length
      );

      if (cooling || runningLow(activeId, estimate)) {
        const spare = pickHealthiest(candidatesFor(settings, active), estimate);
        if (spare) {
          deps.push({
            id: newId(),
            role: "assistant",
            notice: true,
            content:
              getProvider(active.provider).label +
              " is out of room for the moment" +
              (statusLine(activeId) ? " (" + statusLine(activeId) + ")" : "") +
              ", carrying on with " +
              spare.label +
              ".",
          });
          active = { provider: spare.provider, model: spare.model };
          switches++;
        } else if (cooling) {
          deps.setStatus("waiting " + cooling + "s for " + getProvider(active.provider).label);
          await sleep(cooling * 1000 + 300, deps.signal);
        }
      }

      try {
        const res = await callModel({
          provider: getProvider(active.provider),
          apiKey: settings.keys[active.provider] ?? "",
          model: active.model,
          system: buildSystemPrompt(
            settings.caveman,
            deps.getScope(),
            deps.getSelection(),
            prompted ? toolCatalogPrompt(toolsFor(undefined, settings.leanTools)) : undefined
          ),
          msgs: prompted ? flattenForPromptMode(compact(convo)) : compact(convo),
          tools: toolsFor(undefined, settings.leanTools),
          promptTools: prompted,
          signal: deps.signal,
        });

        noteSuccess(idOf(active.provider, active.model), res.rate, res.usage?.totalTokens);

        // Only now is a healed model worth remembering. Saving it at the moment of the
        // switch wrote a broken id into the user's settings and made it the default.
        if (unproven && unproven.provider === active.provider && unproven.model === active.model) {
          deps.onSwitch(unproven.provider, unproven.model);
          unproven = null;
        }

        // In prompt mode the call arrives as JSON inside the text.
        if (prompted && !res.toolCalls.length) {
          const { call, prose } = parsePromptedCall(res.content);
          if (call) return { content: prose, toolCalls: [call] };
        }
        return res;
      } catch (err) {
        if (err instanceof ProviderError && err.kind === "tools_unsupported" && !prompted) {
          promptModels.add(active.provider + "|" + active.model);
          deps.onPromptMode(active.provider, active.model);
          deps.push({
            id: newId(),
            role: "assistant",
            content:
              getProvider(active.provider).label +
              " · " +
              active.model +
              " cannot call tools, so I will ask it in plain text instead. Carrying on.",
          });
          continue;
        }
        // A retired model id is not a reason to abandon a provider that still works.
        // Ask it what it has now and carry on with that. Capped, because a provider can
        // list several ids that each turn out to refuse the job.
        if (
          err instanceof ProviderError &&
          err.kind === "model" &&
          (relists.get(active.provider) ?? 0) < MAX_RELISTS
        ) {
          relists.set(active.provider, (relists.get(active.provider) ?? 0) + 1);
          deadModels.add(idOf(active.provider, active.model));
          const provider = getProvider(active.provider);
          const live = await fetchModels(provider, settings.keys[active.provider] ?? "").catch(
            () => [] as string[]
          );
          // Keep it. The rotation builds its candidates from this, so a stale list is
          // what kept offering retired ids over and over.
          if (live.length) deps.onModels(active.provider, live);
          const replacement = autoPick(
            provider.models,
            live,
            new Set(live.filter((m) => deadModels.has(idOf(active.provider, m))))
          );
          if (replacement) {
            deps.push({
              id: newId(),
              role: "assistant",
              notice: true,
              content:
                getProvider(active.provider).label +
                " cannot use " +
                active.model +
                ", using " +
                replacement +
                " instead.",
            });
            active = { provider: active.provider, model: replacement };
            unproven = { provider: active.provider, model: replacement };
            continue;
          }
        }

        if (!(err instanceof ProviderError) || !err.switchable) throw err;

        const failedId = idOf(active.provider, active.model);
        if (err.kind === "rate_limit") noteRateLimit(failedId, err.retryAfterSec);
        else noteFailure(failedId);

        if (switches >= MAX_SWITCHES) throw err;

        const candidates = candidatesFor(settings, active);

        // A rate limit with somewhere healthy to go is not worth interrupting for.
        if (err.kind === "rate_limit") {
          const spare = pickHealthiest(candidates, 4000);
          if (spare) {
            active = { provider: spare.provider, model: spare.model };
            switches++;
            deps.push({
              id: newId(),
              role: "assistant",
              notice: true,
              content: failureLine(err) + " Carrying on with " + spare.label + ".",
            });
            continue;
          }
          // Nowhere to go, but the window reopens soon: wait it out, a few times at most.
          if (err.retryAfterSec && err.retryAfterSec <= 30 && waits < MAX_WAITS) {
            waits++;
            deps.setStatus("every model is busy, waiting " + err.retryAfterSec + "s");
            await sleep(err.retryAfterSec * 1000 + 500, deps.signal);
            continue;
          }
        }
        const waitable = err.kind === "rate_limit" && err.retryAfterSec && err.retryAfterSec <= 120;

        if (!candidates.length && !waitable) throw err;

        if (settings.autoRotate && candidates.length) {
          active = { provider: candidates[0].provider, model: candidates[0].model };
          switches++;
          deps.push({
            id: newId(),
            role: "assistant",
            content: failureLine(err) + "\nRotated to " + candidates[0].label + " automatically.",
          });
          deps.onSwitch(active.provider, active.model);
          continue;
        }

        const options = candidates.slice(0, 3).map((c) => c.label);
        if (waitable) options.push("Wait " + err.retryAfterSec + "s and retry");
        options.push("Stop");

        const answer = await deps.ask(
          failureLine(err) + " Want me to carry on with another model?",
          options
        );

        if (/^stop$/i.test(answer.trim())) {
          deps.push({ id: newId(), role: "assistant", content: "Stopped.", error: true });
          return null;
        }

        if (/^wait/i.test(answer.trim()) && err.retryAfterSec) {
          deps.setStatus("waiting " + err.retryAfterSec + "s for the limit to reset");
          await sleep(err.retryAfterSec * 1000 + 500, deps.signal);
          continue;
        }

        const chosen = matchChoice(answer, candidates);
        if (!chosen) {
          deps.push({
            id: newId(),
            role: "assistant",
            content: "Did not recognise that choice, stopping. " + err.message,
            error: true,
          });
          return null;
        }

        active = { provider: chosen.provider, model: chosen.model };
        switches++;
        deps.onSwitch(active.provider, active.model);
        deps.push({
          id: newId(),
          role: "assistant",
          notice: true,
          content: "Switched to " + chosen.label + ".",
        });
      }
    }
  }

  for (let step = 0; step < MAX_STEPS; step++) {
    deps.setStatus(step === 0 ? "thinking" : "thinking (step " + (step + 1) + ")");

    const res = await callWithFallback();
    if (!res) {
      deps.setStatus(null);
      return;
    }

    if (!res.toolCalls.length) {
      const final: Msg = { id: newId(), role: "assistant", content: res.content || "(no reply)" };
      convo.push(final);
      deps.push(final);
      deps.setStatus(null);
      return;
    }

    const assistant: Msg = {
      id: newId(),
      role: "assistant",
      content: res.content,
      toolCalls: res.toolCalls,
    };
    convo.push(assistant);
    deps.push(assistant);

    for (const call of res.toolCalls) {
      deps.setStatus(call.name);
      let result: string;
      try {
        result = await execTool(call.name, call.args as Record<string, any>, {
          maxRows: settings.maxRowsPerRead,
          ask: deps.ask,
          setScope: deps.setScope,
          editMode: settings.editMode,
          confirmEdit: deps.confirmEdit,
        });
      } catch (err) {
        result = "ERROR: " + (err instanceof Error ? err.message : String(err));
      }
      const toolMsg: Msg = {
        id: newId(),
        role: "tool",
        content: result,
        toolCallId: call.id,
        toolName: call.name,
        error: result.startsWith("ERROR:"),
      };
      convo.push(toolMsg);
      deps.push(toolMsg);
    }
  }

  deps.push({
    id: newId(),
    role: "assistant",
    content: "Stopped after " + MAX_STEPS + " steps. Tell me how to narrow the task.",
    error: true,
  });
  deps.setStatus(null);
}
