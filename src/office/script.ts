import { detectHost } from "./host";

/**
 * Runs a model written Office.js snippet against the open file, which covers everything
 * the API can do without a tool per feature. The body runs with `context` in scope and
 * must return a string.
 */
export async function runOfficeScript(args: { code: string }): Promise<string> {
  const host = detectHost();
  const body = (args.code ?? "").trim();
  if (!body) return "no code given";

  const wrapped = "return (async (context) => {\n" + body + "\n})(context);";

  const call = async (context: unknown): Promise<string> => {
    // eslint-disable-next-line no-new-func
    const fn = new Function("context", wrapped) as (c: unknown) => Promise<unknown>;
    const out = await fn(context);
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
    await ctx.sync();

    const scratch = ws.getRange("XFD1048576");
    const formula = args.formula.startsWith("=") ? args.formula : "=" + args.formula;
    scratch.formulas = [[formula]];
    await ctx.sync();

    const area = args.spill ? scratch.getSpillingToRangeOrNullObject() : scratch;
    area.load(["values", "isNullObject"]);
    scratch.load("values");
    await ctx.sync();

    const values =
      args.spill && !(area as Excel.Range).isNullObject
        ? ((area as Excel.Range).values as unknown[][])
        : (scratch.values as unknown[][]);

    scratch.clear(Excel.ClearApplyTo.all);
    await ctx.sync();

    const text = values
      .map((row) => row.map((v) => (v === null || v === undefined ? "" : String(v))).join("\t"))
      .join("\n");
    return formula + "  =>  " + text;
  });
}
