import { useEffect, useState } from "react";
import { PROVIDERS, getProvider, type ProviderId } from "../providers/registry";
import { fetchModels } from "../providers/models";
import { clearAllSettings, type Settings } from "../store/settings";
import { APP_VERSION, checkForUpdate, type UpdateCheck } from "../update";

export default function SettingsPanel({
  settings,
  onChange,
}: {
  settings: Settings;
  onChange: (s: Settings) => void;
}) {
  const provider = getProvider(settings.provider);
  const [refreshing, setRefreshing] = useState(false);
  const [modelMsg, setModelMsg] = useState<string | null>(null);
  const [confirmClear, setConfirmClear] = useState(false);
  const [update, setUpdate] = useState<UpdateCheck | null>(null);
  const [checking, setChecking] = useState(false);

  async function runUpdateCheck() {
    setChecking(true);
    setUpdate(await checkForUpdate());
    setChecking(false);
  }

  // A quiet look on open, once a day at most, and only if the user allows it.
  useEffect(() => {
    if (!settings.autoCheckUpdates) return;
    const last = Number(localStorage.getItem("tank.lastUpdateCheck") ?? 0);
    if (Date.now() - last < 86400000) return;
    localStorage.setItem("tank.lastUpdateCheck", String(Date.now()));
    checkForUpdate().then(setUpdate);
  }, [settings.autoCheckUpdates]);

  const models = settings.modelCache[settings.provider] ?? provider.models;

  /** Task panes cannot just follow a link, Office has to open the browser for them. */
  function openKeyPage(url: string) {
    const ui = (Office as unknown as { context?: { ui?: { openBrowserWindow?: (u: string) => void } } })
      ?.context?.ui;
    if (ui?.openBrowserWindow) ui.openBrowserWindow(url);
    else window.open(url, "_blank", "noopener");
  }
  const currentPair = settings.provider + "|" + settings.model;

  function pickProvider(id: ProviderId) {
    const p = getProvider(id);
    const cached = settings.modelCache[id];
    onChange({ ...settings, provider: id, model: (cached ?? p.models)[0] });
    setModelMsg(null);
  }

  async function refreshModels() {
    setRefreshing(true);
    setModelMsg(null);
    try {
      const ids = await fetchModels(provider, settings.keys[settings.provider] ?? "");
      onChange({
        ...settings,
        modelCache: { ...settings.modelCache, [settings.provider]: ids },
        model: ids.includes(settings.model) ? settings.model : (ids[0] ?? settings.model),
      });
      setModelMsg(ids.length + " models available");
    } catch (err) {
      setModelMsg(err instanceof Error ? err.message : String(err));
    } finally {
      setRefreshing(false);
    }
  }

  return (
    <div className="sheet">
      <h2>Model</h2>
      <div className="card">
        <label className="field">
          <span>Provider</span>
          <select
            value={settings.provider}
            onChange={(e) => pickProvider(e.target.value as ProviderId)}
          >
            {PROVIDERS.map((p) => (
              <option key={p.id} value={p.id}>
                {p.label}
              </option>
            ))}
          </select>
        </label>
        {provider.note && <p className="note">{provider.note}</p>}

        <label className="field">
          <span>Model</span>
          <select
            value={settings.model}
            onChange={(e) => onChange({ ...settings, model: e.target.value })}
          >
            {(models.includes(settings.model) ? models : [settings.model, ...models]).map((m) => (
              <option key={m} value={m}>
                {m}
              </option>
            ))}
          </select>
        </label>

        <div className="row">
          <button type="button" className="mini" onClick={refreshModels} disabled={refreshing}>
            {refreshing ? "Refreshing…" : "Refresh models"}
          </button>
          <span className="note" aria-live="polite">
            {modelMsg}
          </span>
        </div>

        <label className="field">
          <span>API key for {provider.label}</span>
          <input
            type="text"
            className="secret"
            name="tank-provider-token"
            autoComplete="off"
            autoCorrect="off"
            autoCapitalize="off"
            spellCheck={false}
            data-lpignore="true"
            data-1p-ignore=""
            data-form-type="other"
            placeholder={"paste your " + provider.label + " key…"}
            value={settings.keys[settings.provider] ?? ""}
            onChange={(e) =>
              onChange({
                ...settings,
                keys: { ...settings.keys, [settings.provider]: e.target.value },
              })
            }
          />
        </label>

        <div className="row">
          <button
            type="button"
            className="mini"
            disabled={!settings.keys[settings.provider]}
            onClick={() =>
              onChange({ ...settings, keys: { ...settings.keys, [settings.provider]: "" } })
            }
          >
            Clear this key
          </button>
          <button
            type="button"
            className="mini danger"
            title="Erase every key and setting tANk has stored on this PC"
            onClick={() => {
              if (!confirmClear) {
                setConfirmClear(true);
                return;
              }
              clearAllSettings();
              location.reload();
            }}
          >
            {confirmClear ? "Tap again to erase" : "Erase everything"}
          </button>
          {confirmClear && (
            <button type="button" className="mini" onClick={() => setConfirmClear(false)}>
              Cancel
            </button>
          )}
        </div>

        <div className="row">
          <button type="button" className="mini" onClick={() => openKeyPage(provider.keyUrl)}>
            Get a {provider.label} key
          </button>
          <span className="note">free, opens in your browser</span>
        </div>
        <p className="note">
          Kept on this PC and sent to {provider.label} only. Uninstalling tANk does not reach these,
          so press Erase everything first if you want them gone.
        </p>

        {Object.values(settings.keys).filter(Boolean).length < 2 && (
          <p className="note">
            Only one provider has a key. Add a second and tANk can keep working when this one runs
            out for the minute.
          </p>
        )}
      </div>

      <h2>Updates</h2>
      <div className="card">
        <div className="row">
          <span className="note">
            Version <b>{APP_VERSION}</b>
          </span>
          <button type="button" className="mini" onClick={runUpdateCheck} disabled={checking}>
            {checking ? "Checking…" : "Check for updates"}
          </button>
        </div>

        {update?.error && <p className="note">{update.error}</p>}

        {update && !update.error && update.newer && (
          <>
            <p className="note">
              Version <b>{update.latest}</b> is out. Yours is {update.current}.
            </p>
            {update.url && (
              <div className="row">
                <button type="button" className="mini" onClick={() => openKeyPage(update.url as string)}>
                  Open the download page
                </button>
              </div>
            )}
            <p className="note">
              Close Office, run the new setup, and your keys and settings stay as they are.
            </p>
          </>
        )}

        {update && !update.error && !update.newer && update.latest && (
          <p className="note">Up to date. {update.latest} is the newest release.</p>
        )}

        <label className="check">
          <input
            type="checkbox"
            checked={settings.autoCheckUpdates}
            onChange={(e) => onChange({ ...settings, autoCheckUpdates: e.target.checked })}
          />
          <span className="note">
            Look for new versions once a day. It only looks and tells you. Nothing is ever
            downloaded or installed without you saying so.
          </span>
        </label>

        <p className="note">
          The pane itself refreshes from the web, so small fixes reach you without doing anything.
          A new setup is only needed when the add-in itself changes.
        </p>
      </div>

      <h2>Keys</h2>
      <div className="card">
        <p className="note">
          Each of these is free and none asks for a card except Cerebras. Two or three is plenty.
        </p>
        <ul className="keylist">
          {PROVIDERS.map((p) => {
            const has = Boolean(settings.keys[p.id]);
            return (
              <li key={p.id} className={has ? "has" : undefined}>
                <span className={"dot" + (has ? " on" : "")} aria-hidden="true" />
                <span className="keyname">
                  {p.label}
                  <span className="sub">{p.freeLimit}</span>
                </span>
                <button
                  type="button"
                  className="mini"
                  onClick={() => openKeyPage(p.keyUrl)}
                  aria-label={"Open the " + p.label + " key page"}
                >
                  {has ? "Manage" : "Get key"}
                </button>
                {!has && (
                  <button
                    type="button"
                    className="mini"
                    onClick={() => pickProvider(p.id)}
                    aria-label={"Switch to " + p.label + " to paste its key"}
                  >
                    Paste
                  </button>
                )}
              </li>
            );
          })}
        </ul>
      </div>

      <h2>When a model runs out</h2>
      <div className="card">
        {settings.fallbacks.length === 0 ? (
          <p className="note">
            Nothing set, so tANk offers whichever other providers have a key. Add entries here to fix
            the order it tries them in.
          </p>
        ) : (
          settings.fallbacks.map((f, i) => (
            <div className="row" key={f.provider + f.model + i}>
              <span className="note">
                {i + 1}. {getProvider(f.provider).label} · {f.model}
              </span>
              <button
                type="button"
                className="mini"
                onClick={() =>
                  onChange({ ...settings, fallbacks: settings.fallbacks.filter((_, j) => j !== i) })
                }
              >
                Remove
              </button>
            </div>
          ))
        )}

        <div className="row">
          <button
            type="button"
            className="mini"
            onClick={() =>
              onChange({
                ...settings,
                fallbacks: [
                  ...settings.fallbacks.filter(
                    (f) => !(f.provider === settings.provider && f.model === settings.model)
                  ),
                  { provider: settings.provider, model: settings.model },
                ],
              })
            }
          >
            Add current model
          </button>
        </div>

        <label className="check">
          <input
            type="checkbox"
            checked={settings.autoRotate}
            onChange={(e) => onChange({ ...settings, autoRotate: e.target.checked })}
          />
          <span className="note">Switch to the next one without asking me</span>
        </label>
      </div>

      <h2>Behaviour</h2>
      <div className="card">
        <label className="field">
          <span>Theme</span>
          <select
            value={settings.theme}
            onChange={(e) => onChange({ ...settings, theme: e.target.value as Settings["theme"] })}
          >
            <option value="system">Follow Windows</option>
            <option value="light">Light</option>
            <option value="dark">Dark</option>
          </select>
        </label>

        <label className="field">
          <span>Rows read at a time</span>
          <input
            type="number"
            inputMode="numeric"
            min={50}
            max={5000}
            value={settings.maxRowsPerRead}
            onChange={(e) => onChange({ ...settings, maxRowsPerRead: Number(e.target.value) || 400 })}
          />
        </label>
        <p className="note">
          Reading more rows costs more tokens, and a very big read can go past what a free tier
          allows in one go.
        </p>

        <label className="check">
          <input
            type="checkbox"
            checked={settings.leanTools}
            onChange={(e) => onChange({ ...settings, leanTools: e.target.checked })}
          />
          <span className="note">
            <b>Send a smaller toolbox</b>
            <br />
            Every message carries a description of what tANk can do, about 5,000 tokens in Excel.
            This drops the convenience tools and keeps the general ones, saving roughly 1,400 tokens
            a message. Nothing becomes impossible, it just writes a short script instead. Worth
            turning on if you keep hitting a provider's per minute limit.
          </span>
        </label>

        <label className="check">
          <input
            type="checkbox"
            checked={settings.showToolDetails}
            onChange={(e) => onChange({ ...settings, showToolDetails: e.target.checked })}
          />
          <span className="note">Show every step tANk takes, with the raw data it read</span>
        </label>

        <label className="check">
          <input
            type="checkbox"
            checked={settings.promptToolModels.includes(currentPair)}
            onChange={(e) =>
              onChange({
                ...settings,
                promptToolModels: e.target.checked
                  ? [...settings.promptToolModels, currentPair]
                  : settings.promptToolModels.filter((x) => x !== currentPair),
              })
            }
          />
          <span className="note">
            <b>This model cannot run tools directly</b>
            <br />
            Some smaller models ignore the list of things tANk can do. With this on, tANk writes
            those instructions into the prompt instead and carries out whatever the model asks for.
            Slower, but it works. tANk ticks this itself the first time a model refuses, so only
            touch it if a model chats back happily and never touches your sheet.
          </span>
        </label>
      </div>
    </div>
  );
}
