import { useEffect, useRef } from "react";
import type { Msg } from "../types";

function ToolRow({ m }: { m: Msg }) {
  return (
    <details className={"step" + (m.error ? " failed" : "")}>
      <summary>
        <span className="stepname" translate="no">
          {m.toolName}
        </span>
        <span className="stepstate">{m.error ? "failed" : "done"}</span>
      </summary>
      <pre>{m.content}</pre>
    </details>
  );
}

function CallRow({ name, args }: { name: string; args: Record<string, unknown> }) {
  const preview = JSON.stringify(args);
  return (
    <details className="step">
      <summary>
        <span className="stepname" translate="no">
          {name}
        </span>
        <span className="stepargs">
          {preview.length > 64 ? preview.slice(0, 64) + "…" : preview}
        </span>
      </summary>
      <pre>{JSON.stringify(args, null, 2)}</pre>
    </details>
  );
}

export default function MessageList({
  msgs,
  status,
  showDetails,
  onToggleDetails,
}: {
  msgs: Msg[];
  status: string | null;
  showDetails: boolean;
  onToggleDetails: () => void;
}) {
  const endRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    endRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [msgs.length, status, showDetails]);

  const hiddenSteps = showDetails ? 0 : msgs.filter((m) => m.role === "tool").length;

  return (
    <div className="thread" role="log" aria-live="polite" aria-label="Conversation">
      {msgs.map((m) => {
        if (m.role === "tool") return showDetails ? <ToolRow key={m.id} m={m} /> : null;

        if (m.role === "user")
          return (
            <p key={m.id} className="said you">
              {m.content}
            </p>
          );

        // A turn that was only a tool call has nothing of its own to say.
        if (!showDetails && !m.content) return null;

        return (
          <div key={m.id} className={"said tank" + (m.error ? " bad" : "")}>
            {m.content && <p className="text">{m.content}</p>}
            {showDetails &&
              m.toolCalls?.map((c) => <CallRow key={c.id} name={c.name} args={c.args} />)}
          </div>
        );
      })}

      {status && (
        <p className="working">
          <span className="pulse" aria-hidden="true" />
          {status}…
        </p>
      )}

      {(hiddenSteps > 0 || showDetails) && (
        <button type="button" className="linky" onClick={onToggleDetails}>
          {showDetails
            ? "Hide the steps"
            : "Show " + hiddenSteps + (hiddenSteps === 1 ? " step" : " steps")}
        </button>
      )}

      <div ref={endRef} />
    </div>
  );
}
