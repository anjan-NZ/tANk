import { detectHost } from "./host";

/** A script drives the open file through `context`. Everything here is a way out of that. */
const ESCAPES =
  /(?<![.\w$])(globalThis|window|document|self|parent|top|eval|localStorage|sessionStorage|indexedDB|XMLHttpRequest|WebSocket|EventSource|fetch|importScripts|navigator)\b|(?<![.\w$])(Function|import)\s*\(|constructor\s*\[/;

/** Blanks strings and comments. Scanned, not matched: the // in a URL is not a comment. */
function codeOnly(src: string): string {
  let out = "";
  let i = 0;

  while (i < src.length) {
    const c = src[i];
    const next = src[i + 1];

    if (c === "/" && next === "/") {
      while (i < src.length && src[i] !== "\n") i++;
      out += " ";
    } else if (c === "/" && next === "*") {
      i += 2;
      while (i < src.length && !(src[i] === "*" && src[i + 1] === "/")) i++;
      i += 2;
      out += " ";
    } else if (c === '"' || c === "'") {
      i++;
      while (i < src.length && src[i] !== c) i += src[i] === "\\" ? 2 : 1;
      i++;
      out += '""';
    } else if (c === "`") {
      i++;
      while (i < src.length && src[i] !== "`") {
        if (src[i] === "\\") {
          i += 2;
        } else if (src[i] === "$" && src[i + 1] === "{") {
          i += 2;
          const start = i;
          for (let depth = 1; i < src.length && depth > 0; ) {
            if (src[i] === "{") depth++;
            else if (src[i] === "}") depth--;
            if (depth > 0) i++;
          }
          out += "(" + src.slice(start, i) + ")";
          i++;
        } else {
          i++;
        }
      }
      i++;
      out += "``";
    } else {
      out += c;
      i++;
    }
  }

  return out;
}

/** Names the sandbox refuses, or null when the snippet keeps to `context`. */
export function escapesSandbox(code: string): string | null {
  return codeOnly(code).match(ESCAPES)?.[0] ?? null;
}

/** Shadowed inside the script, so the bare identifiers resolve to undefined. */
const SHADOWED = [
  "fetch",
  "XMLHttpRequest",
  "WebSocket",
  "EventSource",
  "localStorage",
  "sessionStorage",
  "indexedDB",
  "navigator",
  "document",
  "window",
  "globalThis",
  "self",
];

/**
 * Runs a model written Office.js snippet against the open file, which covers everything
 * the API can do without a tool per feature. The body runs with `context` in scope and
 * must return a string.
 */
export async function runOfficeScript(args: { code: string }): Promise<string> {
  const host = detectHost();
  const body = (args.code ?? "").trim();
  if (!body) return "no code given";

  const escape = escapesSandbox(body);
  if (escape)
    return (
      "Refused: a script may only use `context` to work on the open file, and this one reaches for `" +
      escape +
      "`. Rewrite it using the Office.js API alone."
    );

  const wrapped = "return (async (context) => {\n" + body + "\n})(context);";

  const call = async (context: unknown): Promise<string> => {
    // eslint-disable-next-line no-new-func
    const fn = new Function("context", ...SHADOWED, wrapped) as (
      c: unknown,
      ...rest: undefined[]
    ) => Promise<unknown>;
    const out = await fn(context, ...SHADOWED.map(() => undefined));
    if (out === undefined || out === null) return "done";
    return typeof out === "string" ? out : JSON.stringify(out);
  };

  if (host === "word") return Word.run(async (ctx) => call(ctx));
  if (host === "powerpoint") return PowerPoint.run(async (ctx) => call(ctx));
  return Excel.run(async (ctx) => call(ctx));
}

/**
 * Evaluates a formula in a scratch cell at the far corner of the sheet, reads the result
 * and clears it, so every worksheet function is usable without leaving anything behind.
 */
export async function evaluateFormula(args: {
  sheet?: string;
  formula: string;
  spill?: boolean;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = args.sheet
      ? ctx.workbook.worksheets.getItem(args.sheet)
      : ctx.workbook.worksheets.getActiveWorksheet();
    ws.load("name");
    // The far corner cannot spill, and a different sheet would break unqualified references.
    const used = ws.getUsedRangeOrNullObject(true);
    used.load(["rowIndex", "rowCount", "isNullObject"]);
    await ctx.sync();

    const below = used.isNullObject ? 0 : used.rowIndex + used.rowCount + 4;
    const scratch =
      args.spill && below < 1_000_000 ? ws.getRangeByIndexes(below, 0, 1, 1) : ws.getRange("XFD1048576");
    const formula = args.formula.startsWith("=") ? args.formula : "=" + args.formula;
    let values: unknown[][];
    try {
      scratch.formulas = [[formula]];
      await ctx.sync();

      const area = args.spill ? scratch.getSpillingToRangeOrNullObject() : scratch;
      area.load(["values", "isNullObject"]);
      scratch.load("values");
      await ctx.sync();

      values =
        args.spill && !(area as Excel.Range).isNullObject
          ? ((area as Excel.Range).values as unknown[][])
          : (scratch.values as unknown[][]);
    } finally {
      scratch.clear(Excel.ClearApplyTo.all);
      await ctx.sync();
    }

    const text = values
      .map((row) => row.map((v) => (v === null || v === undefined ? "" : String(v))).join("\t"))
      .join("\n");
    return formula + "  =>  " + text;
  });
}
