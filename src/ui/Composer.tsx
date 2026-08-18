import { useState } from "react";
import type { AskRequest } from "../types";
import Icon from "./Icon";

export default function Composer({
  onSend,
  busy,
  ask,
  onAnswer,
  onStop,
}: {
  onSend: (text: string) => void;
  busy: boolean;
  ask: AskRequest | null;
  onAnswer: (answer: string) => void;
  onStop: () => void;
}) {
  const [text, setText] = useState("");

  function submit() {
    const t = text.trim();
    if (!t) return;
    setText("");
    onSend(t);
  }

  return (
    <div className="composer">
      {ask && (
        <div className="askcard" role="group" aria-label="tANk needs an answer">
          <p className="askq">{ask.question}</p>
          {ask.options.length > 0 && (
            <div className="chips">
              {ask.options.map((o) => (
                <button type="button" key={o} className="chip" onClick={() => onAnswer(o)}>
                  {o}
                </button>
              ))}
            </div>
          )}
        </div>
      )}

      <div className="inputrow">
        <label className="sr-only" htmlFor="tank-input">
          Message tANk
        </label>
        <textarea
          id="tank-input"
          name="message"
          value={text}
          spellCheck={false}
          placeholder={ask ? "Type an answer, or pick one above…" : "Ask about the selected cells…"}
          onChange={(e) => setText(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              submit();
            }
          }}
          rows={1}
        />
        {busy && !ask ? (
          <button type="button" className="send stop" onClick={onStop} aria-label="Stop">
            <span className="square" aria-hidden="true" />
          </button>
        ) : (
          <button
            type="button"
            className="send"
            onClick={submit}
            disabled={!text.trim()}
            aria-label="Send message"
          >
            <Icon name="send" />
          </button>
        )}
      </div>

      <p className="disclaimer">tANk is AI. It gets things wrong, so check its work.</p>
    </div>
  );
}
