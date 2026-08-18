import { useEffect, useState } from "react";
import { snapshot, totals, type UsageRow } from "../agent/budget";
import { PROVIDERS, type ProviderId } from "../providers/registry";
import type { Settings } from "../store/settings";

const label = (id: string) => PROVIDERS.find((p) => p.id === id)?.label ?? id;

function fmt(n?: number): string {
  if (n === undefined) return "-";
  return n >= 1000 ? Math.round(n / 100) / 10 + "k" : String(Math.round(n));
}

function ago(ts: number): string {
  if (!ts) return "not used";
  const s = Math.round((Date.now() - ts) / 1000);
  if (s < 60) return s + "s ago";
  if (s < 3600) return Math.round(s / 60) + "m ago";
  return Math.round(s / 3600) + "h ago";
}

function State({ row }: { row: UsageRow }) {
  if (row.cooling) return <span className="state resting">resting {row.cooling}s</span>;
  if (row.remainingRequests !== undefined && row.remainingRequests <= 1)
    return <span className="state resting">at its limit</span>;
  return <span className="state ready">ready</span>;
}

export default function UsagePanel({ settings }: { settings: Settings }) {
  const [rows, setRows] = useState<UsageRow[]>(snapshot());
  const [sum, setSum] = useState(totals());

  // Cooldowns count down, so keep the view honest while it is open.
  useEffect(() => {
    const t = setInterval(() => {
      setRows(snapshot());
      setSum(totals());
    }, 2000);
    return () => clearInterval(t);
  }, []);

  const unused = PROVIDERS.filter(
    (p) => settings.keys[p.id as ProviderId] && !rows.some((r) => r.provider === p.id)
  );

  return (
    <div className="sheet">
      <h2>This session</h2>
      <div className="card">
        <div className="bigstats">
          <div>
            <strong>{sum.calls}</strong>
            <span>calls</span>
          </div>
          <div>
            <strong>{fmt(sum.tokens)}</strong>
            <span>tokens</span>
          </div>
          <div>
            <strong>{sum.limitHits}</strong>
            <span>limits hit</span>
          </div>
        </div>
        <p className="note">
          Counted from what each provider reports back. It resets when the pane is closed, and it
          costs nothing: these are your own free tiers.
        </p>
      </div>

      <h2>Models used</h2>
      {rows.length === 0 ? (
        <div className="card">
          <p className="note">Nothing yet. Ask something and the numbers will appear here.</p>
        </div>
      ) : (
        <div className="tablewrap">
          <table className="ptable">
            <thead>
              <tr>
                <th>Model</th>
                <th>Calls</th>
                <th>Tokens</th>
                <th>Left this minute</th>
                <th>State</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.id}>
                  <td>
                    {label(r.provider)}
                    <div className="sub mono">{r.model}</div>
                    <div className="sub">{ago(r.lastUsed)}</div>
                  </td>
                  <td className="numcell">{r.calls}</td>
                  <td className="numcell">{fmt(r.tokens)}</td>
                  <td className="numcell">
                    {r.remainingRequests === undefined && r.remainingTokens === undefined
                      ? "not reported"
                      : fmt(r.remainingRequests) + " calls / " + fmt(r.remainingTokens) + " tok"}
                  </td>
                  <td>
                    <State row={r} />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {unused.length > 0 && (
        <>
          <h2>Standing by</h2>
          <div className="card">
            <p className="note">
              These have a key and have not been needed yet: {unused.map((p) => p.label).join(", ")}.
              tANk will move to them if the model in use runs out.
            </p>
          </div>
        </>
      )}

      <p className="note">
        Gemini and a few others do not publish what is left in the window, so their "left this
        minute" stays blank until they refuse a call.
      </p>
    </div>
  );
}
