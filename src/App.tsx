import { useCallback, useEffect, useRef, useState } from "react";
import type { AskRequest, Msg, Scope } from "./types";
import { loadSettings, saveSettings, type Settings } from "./store/settings";
import { runTurn, newId } from "./agent/loop";
import { proposeScope, readSelection, scopeFromSelection, type SelectionInfo } from "./office/scope";
import { undoLast, undoDepth } from "./office/undo";
import { detectHost, HOST_LABEL } from "./office/host";
import { checkForUpdate } from "./update";
import Icon from "./ui/Icon";
import SettingsPanel from "./ui/SettingsPanel";
import HelpPanel from "./ui/HelpPanel";
import UsagePanel from "./ui/UsagePanel";
import MessageList from "./ui/MessageList";
import Composer from "./ui/Composer";

const HOST = detectHost();
const inExcel = () => HOST === "excel";

const WELCOME_BY_HOST: Record<string, string> = {
  excel:
    "Pick some cells and ask, or just ask and I will work out where to look.\n" +
    'People usually start with things like "why does this TB not tally", ' +
    '"add a variance column and flag anything over 10%", or "compare TB with the P&L".',
  word:
    "Select a paragraph and ask, or tell me what the document needs.\n" +
    'Things people ask: "tighten this paragraph", "turn these notes into headings", ' +
    '"comment on the wording in the scope section".',
  powerpoint:
    "Tell me what the deck needs and I will build it.\n" +
    'Things people ask: "make five slides on the audit findings", ' +
    '"shorten the bullets on slide 3", "what is in this deck".',
  none: "Open me from inside Excel, Word or PowerPoint and I can work on the open file.",
};

const WELCOME: Msg = {
  id: "welcome",
  role: "assistant",
  content: WELCOME_BY_HOST[HOST],
};

export default function App() {
  const [settings, setSettings] = useState<Settings>(loadSettings);
  const [msgs, setMsgs] = useState<Msg[]>([WELCOME]);
  const [scope, setScope] = useState<Scope | null>(null);
  const [pinned, setPinned] = useState(false);
  const [selLabel, setSelLabel] = useState<string>("");
  const [status, setStatus] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [ask, setAsk] = useState<AskRequest | null>(null);
  const [view, setView] = useState<"chat" | "settings" | "help" | "usage">("chat");
  const [undoCount, setUndoCount] = useState(0);
  const [updateReady, setUpdateReady] = useState<string | null>(null);

  const askResolver = useRef<((answer: string) => void) | null>(null);
  const abortRef = useRef<AbortController | null>(null);
  const scopeRef = useRef<Scope | null>(null);
  scopeRef.current = scope;
  const selRef = useRef<SelectionInfo | null>(null);
  const askUserRef = useRef<(q: string, o: string[]) => Promise<string>>(
    () => Promise.resolve("")
  );

  useEffect(() => saveSettings(settings), [settings]);

  // Looks for a new release at most once a day, and only if the user left the tick on.
  // Finding one just lights a dot on the settings button; nothing downloads or installs.
  useEffect(() => {
    if (!settings.autoCheckUpdates) return;
    const last = Number(localStorage.getItem("tank.lastUpdateCheck") ?? 0);
    if (Date.now() - last < 86400000) return;
    localStorage.setItem("tank.lastUpdateCheck", String(Date.now()));
    checkForUpdate().then((u) => {
      if (u.newer && u.latest) setUpdateReady(u.latest);
    });
  }, [settings.autoCheckUpdates]);

  useEffect(() => {
    const root = document.documentElement;
    if (settings.theme === "system") root.removeAttribute("data-theme");
    else root.setAttribute("data-theme", settings.theme);
  }, [settings.theme]);

  const refreshSelection = useCallback(async () => {
    if (!inExcel()) return;
    try {
      const sel = await readSelection();
      selRef.current = sel;
      setSelLabel(sel.sheet + "!" + sel.address);
      if (!pinned && !(sel.rows === 1 && sel.cols === 1)) {
        setScope(scopeFromSelection(sel));
      }
    } catch {
      /* selection can be unavailable mid-edit; ignore */
    }
  }, [pinned]);

  useEffect(() => {
    if (!inExcel()) return;
    refreshSelection();
    let handler: any;
    Excel.run(async (ctx) => {
      handler = ctx.workbook.onSelectionChanged.add(async () => {
        refreshSelection();
      });
      await ctx.sync();
    }).catch(() => undefined);
    return () => {
      if (handler) Excel.run(handler.context, async (ctx) => {
        handler.remove();
        await ctx.sync();
      }).catch(() => undefined);
    };
  }, [refreshSelection]);

  const confirmEdit = useCallback(async (summary: string): Promise<boolean> => {
    const answer = await askUserRef.current(summary, ["Yes, change it", "No, leave it"]);
    return /^yes/i.test(answer.trim());
  }, []);

  const askUser = useCallback((question: string, options: string[]): Promise<string> => {
    setAsk({ question, options });
    setStatus(null);
    return new Promise<string>((resolve) => {
      askResolver.current = (answer) => {
        setAsk(null);
        askResolver.current = null;
        push({ id: newId(), role: "user", content: answer });
        resolve(answer);
      };
    });
  }, []);

  const push = (m: Msg) => setMsgs((prev) => [...prev, m]);
  askUserRef.current = askUser;

  async function send(text: string) {
    const trimmed = text.trim();
    if (!trimmed) return;

    // A pending ask_user question takes the next thing typed as its answer.
    if (askResolver.current) {
      askResolver.current(trimmed);
      return;
    }
    if (busy) return;

    const userMsg: Msg = { id: newId(), role: "user", content: trimmed };
    const history = [...msgs.filter((m) => m.id !== "welcome"), userMsg];
    push(userMsg);
    setBusy(true);
    abortRef.current = new AbortController();

    try {
      if (inExcel()) {
        try {
          selRef.current = await readSelection();
        } catch {
          /* mid-edit, keep the last known selection */
        }
      }

      if (inExcel() && !pinned) {
        const p = await proposeScope();
        if (p.scope) {
          setScope(p.scope);
          scopeRef.current = p.scope;
        } else if (p.needsConfirm && p.proposal) {
          const answer = await askUser(
            "Not much is selected. Should I work on " + p.proposal.label + "?",
            ["Yes, use that", "Use the whole sheet", "I will select it myself"]
          );
          if (/^yes/i.test(answer)) {
            setScope(p.proposal);
            scopeRef.current = p.proposal;
            setPinned(true);
          } else if (/whole/i.test(answer)) {
            const sel = await readSelection();
            const whole: Scope = {
              ranges: [{ sheet: sel.sheet, address: "" }],
              label: sel.sheet + " (used range)",
              source: "auto",
            };
            setScope(whole);
            scopeRef.current = whole;
          } else {
            setScope(null);
            scopeRef.current = null;
          }
        }
      }

      await runTurn(history, {
        settings,
        getScope: () => scopeRef.current,
        getSelection: () => selRef.current,
        setScope: (s) => {
          scopeRef.current = s;
          setScope(s);
          setPinned(true);
        },
        ask: askUser,
        confirmEdit,
        push,
        setStatus,
        onSwitch: (provider, model) =>
          setSettings((prev) => ({ ...prev, provider, model })),
        onPromptMode: (provider, model) =>
          setSettings((prev) =>
            prev.promptToolModels.includes(provider + "|" + model)
              ? prev
              : { ...prev, promptToolModels: [...prev.promptToolModels, provider + "|" + model] }
          ),
        signal: abortRef.current.signal,
      });
    } catch (err) {
      push({
        id: newId(),
        role: "assistant",
        content: err instanceof Error ? err.message : String(err),
        error: true,
      });
    } finally {
      setBusy(false);
      setStatus(null);
      setUndoCount(undoDepth());
    }
  }

  async function onUndo() {
    const label = await undoLast();
    setUndoCount(undoDepth());
    push({
      id: newId(),
      role: "assistant",
      content: label ? "Reverted: " + label : "Nothing to undo.",
    });
  }

  const address = inExcel()
    ? scope
      ? scope.label
      : selLabel || "nothing selected"
    : HOST_LABEL[HOST] + " document";

  return (
    <div className="app">
      <header className="topbar">
        <img className="mark" src="assets/logo.png" alt="" width={18} height={21} />
        <h1 className="brand" translate="no">
          tANk
        </h1>
        <div className="actions">
          <button
            type="button"
            className="icon"
            disabled={!undoCount}
            onClick={onUndo}
            aria-label={"Undo the last change" + (undoCount ? " (" + undoCount + " stored)" : "")}
            title="Put back whatever tANk last changed"
          >
            <Icon name="undo" />
            {undoCount > 0 && <span className="badge">{undoCount}</span>}
          </button>
          <button
            type="button"
            className="icon"
            aria-label="Usage and limits"
            aria-pressed={view === "usage"}
            title="How much of each free tier you have used"
            onClick={() => setView((v) => (v === "usage" ? "chat" : "usage"))}
          >
            <Icon name="stats" />
          </button>
          <button
            type="button"
            className="icon"
            aria-label="Help, free limits and privacy"
            aria-pressed={view === "help"}
            title="Free limits, privacy and how tANk works"
            onClick={() => setView((v) => (v === "help" ? "chat" : "help"))}
          >
            <Icon name="help" />
          </button>
          <button
            type="button"
            className="icon"
            aria-label={updateReady ? "Settings, version " + updateReady + " is available" : "Settings"}
            aria-pressed={view === "settings"}
            title={
              updateReady
                ? "Version " + updateReady + " is available. Open Settings to see it."
                : "Providers, keys and behaviour"
            }
            onClick={() => setView((v) => (v === "settings" ? "chat" : "settings"))}
          >
            <Icon name="gear" />
            {updateReady && <span className="newdot" aria-hidden="true" />}
          </button>
        </div>
      </header>

      <div className="context">
        <span className={"ctxdot" + (pinned ? " pinned" : "")} aria-hidden="true" />
        <span className="ctxlabel">{inExcel() ? (pinned ? "Pinned" : "Working on") : "Open in"}</span>
        <code className="addr" title={address}>
          {address}
        </code>
        {inExcel() && (
          <button
            type="button"
            className="ghost"
            onClick={async () => {
              if (pinned) {
                setPinned(false);
                refreshSelection();
                return;
              }
              const sel = await readSelection();
              setScope(scopeFromSelection(sel));
              setPinned(true);
            }}
            title={
              pinned
                ? "Follow the cursor again"
                : "Keep working on this range even when you click elsewhere"
            }
          >
            {pinned ? "Unpin" : "Pin"}
          </button>
        )}
      </div>

      {view === "settings" && <SettingsPanel settings={settings} onChange={setSettings} />}
      {view === "help" && <HelpPanel settings={settings} />}
      {view === "usage" && <UsagePanel settings={settings} />}

      {view === "chat" && (
        <>
          <MessageList
            msgs={msgs}
            status={status}
            showDetails={settings.showToolDetails}
            onToggleDetails={() =>
              setSettings({ ...settings, showToolDetails: !settings.showToolDetails })
            }
          />

          <div className="modes">
            <label className="mode">
              <span className="modelabel">Style</span>
              <select
                value={settings.caveman}
                title="How tANk talks to you. Shorter styles use fewer tokens. It never changes what goes into cells."
                onChange={(e) =>
                  setSettings({ ...settings, caveman: e.target.value as Settings["caveman"] })
                }
              >
                <option value="off">Normal</option>
                <option value="lite">Short</option>
                <option value="ultra">Caveman</option>
              </select>
            </label>
            <label className="mode">
              <span className="modelabel">Edits</span>
              <select
                className={settings.editMode === "auto" ? "risky" : undefined}
                value={settings.editMode}
                title="Whether tANk may replace cells that already have something in them"
                onChange={(e) =>
                  setSettings({ ...settings, editMode: e.target.value as Settings["editMode"] })
                }
              >
                <option value="ask">Ask me</option>
                <option value="auto">Auto</option>
              </select>
            </label>
          </div>

          <Composer
            onSend={send}
            busy={busy}
            ask={ask}
            onAnswer={(a) => askResolver.current?.(a)}
            onStop={() => abortRef.current?.abort()}
          />
        </>
      )}
    </div>
  );
}

