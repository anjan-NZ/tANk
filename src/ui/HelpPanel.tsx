import { PROVIDERS } from "../providers/registry";
import type { Settings } from "../store/settings";

export default function HelpPanel({ settings }: { settings: Settings }) {
  return (
    <div className="sheet help">
      <h2>Privacy</h2>
      <div className="privacy">
        <p>
          tANk runs on this PC. There is no tANk server, no account, and nothing is logged anywhere.
        </p>
        <ul className="tips">
          <li>
            Your keys sit in this add-in's storage on this computer. Each key only ever goes to the
            provider it belongs to.
          </li>
          <li>
            Your workbook is not uploaded. When you ask something, the cells in the current scope go
            to whichever provider you picked, because the model has to see the numbers to answer.
            Nothing is sent while you are typing or clicking around.
          </li>
          <li>The chat only lives in this pane. Close it and the conversation is gone.</li>
          <li>
            If a provider should not see a client's data, do not give it a key. tANk can only reach
            providers you have added a key for, and Settings clears a key any time.
          </li>
        </ul>
      </div>

      <h2>Running out mid-job</h2>
      <p className="note">
        Free tiers cap how much you can send per minute, not just per day, and one question can take
        several rounds of thinking. tANk watches what each provider says it has left and moves to
        another model before it hits the wall, then tells you in the chat. If everything is busy and
        the wait is short it just waits. Give it keys for two or three providers and you will rarely
        notice a limit at all.
      </p>

      <h2>Free tiers</h2>
      <p className="note">
        These are your own accounts. A green dot means a key is saved here. When one runs out of
        quota, tANk stops and asks whether to move to another.
      </p>

      <div className="tablewrap">
        <table className="ptable">
          <thead>
            <tr>
              <th>Provider</th>
              <th>Free limit</th>
              <th>Notes</th>
            </tr>
          </thead>
          <tbody>
            {PROVIDERS.map((p) => (
              <tr key={p.id}>
                <td>
                  <span className={"dot" + (settings.keys[p.id] ? " on" : "")} />
                  {p.label}
                  <div className="sub mono">{p.keyUrl.replace(/^https:\/\//, "")}</div>
                </td>
                <td>{p.freeLimit}</td>
                <td>{p.note}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <h2>Word and PowerPoint</h2>
      <p className="note">
        The same pane opens in Word and PowerPoint, with tools that suit them. In Word it reads the
        document, rewrites the selection, inserts headings, lists and tables, runs find and replace,
        and can leave comments instead of editing. In PowerPoint it lists the deck, builds title and
        bullet slides, edits the text in a shape, and adds or deletes slides.
      </p>

      <h2>What tANk works on</h2>
      <div className="tablewrap">
        <table className="ptable">
          <tbody>
            <tr>
              <td>You selected a range</td>
              <td>That is what it works on. It will not ask.</td>
            </tr>
            <tr>
              <td>You selected one cell</td>
              <td>
                If that cell sits inside a block of data it suggests the whole block and waits for a
                yes. A cell on its own is used as it is, so "put it here" works.
              </td>
            </tr>
            <tr>
              <td>Nothing selected</td>
              <td>It looks at the sheets in the file and asks where you want it to work.</td>
            </tr>
            <tr>
              <td>Work across sheets</td>
              <td>Just say so, like "compare TB with the P&amp;L". It can name each sheet itself.</td>
            </tr>
            <tr>
              <td>Pin selection</td>
              <td>Locks the scope so it stops changing every time you click elsewhere.</td>
            </tr>
          </tbody>
        </table>
      </div>

      <h2>Reply style</h2>
      <p className="note">
        The Style box at the top changes how tANk talks to you. It never changes what goes into your
        cells, which is always written in plain, normal English.
      </p>
      <div className="tablewrap">
        <table className="ptable">
          <tbody>
            <tr>
              <td>Normal</td>
              <td>Ordinary sentences, full explanations.</td>
            </tr>
            <tr>
              <td>Short</td>
              <td>Same answer with the padding cut out. Usually two or three lines.</td>
            </tr>
            <tr>
              <td>Caveman</td>
              <td>Shortest of the three. Reads like notes rather than sentences.</td>
            </tr>
          </tbody>
        </table>
      </div>
      <p className="note">Shorter replies use fewer tokens, so a free tier lasts longer.</p>

      <h2>Before it changes anything</h2>
      <p className="note">
        The Edits box at the top decides what happens when tANk is about to write over cells that
        already have something in them.
      </p>
      <div className="tablewrap">
        <table className="ptable">
          <tbody>
            <tr>
              <td>Edits: ask me</td>
              <td>
                It stops and shows what it wants to do, like "Write Sheet1!A1:C20? 42 of 60 cells
                there already have something in them." Nothing happens until you say yes. Empty
                cells are filled without asking.
              </td>
            </tr>
            <tr>
              <td>Edits: auto</td>
              <td>
                It writes straight away, over anything. Faster, and easy to regret. Undo still
                works, one step at a time.
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <h2>Worth knowing</h2>
      <ul className="tips">
        <li>
          Ask for highlighting by a rule ("colour anything over 10", "flag duplicates", "top 5 in
          green") and you get real conditional formatting, so it keeps working as the numbers
          change. Plain fills are only used for fixed styling like a header row.
        </li>
        <li>
          The Undo button puts back whatever tANk last changed. Excel's own Ctrl+Z will not, because
          add-in edits skip Excel's undo history.
        </li>
        <li>
          When a model runs out of quota, tANk tells you which one and offers the others you have
          keys for. It waits for your answer unless you tick "rotate automatically" in Settings.
        </li>
        <li>
          It reads at most {settings.maxRowsPerRead} rows at a time and says so when a range is
          longer. Raise that in Settings for big jobs, lower it to save tokens.
        </li>
        <li>
          If a model id stops working, use Refresh models in Settings. Providers retire models often
          and the list here is only a starting point.
        </li>
        <li>
          The steps it takes are hidden. The line under the chat opens them up if you want to see
          exactly which cells it read and wrote.
        </li>
        <li>
          Every model works, even ones that cannot run tools. Normally tANk hands the model a menu
          of things it can do to your sheet. A model that cannot read that menu gets the same
          instructions written into the prompt, and tANk carries out what it asks for. It switches
          over on its own the first time a model refuses.
        </li>
      </ul>

      <h2>Removing tANk</h2>
      <p className="note">
        Windows Settings, then Apps, find tANk and uninstall it. That takes away the button, the
        registration and the files.
      </p>
      <p className="note">
        Your API keys are held by this pane, in Office's own web storage, which no uninstaller can
        pick apart without wiping what every other add-in has saved. So if you want them gone,
        press <b>Erase everything</b> in Settings before you uninstall. It clears the keys and every
        setting at once, and takes effect immediately.
      </p>

      <p className="note eduonly">
        tANk is a personal project, shared for educational purposes only. It comes with no warranty
        and has no connection to Microsoft or to any of the AI providers. Check whatever it writes
        before you rely on it for real work.
      </p>
    </div>
  );
}
