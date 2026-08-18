/** Audit work: reconciliations, sampling, duplicate and gap hunting, formula review, ageing. */

function sheetOf(ctx: Excel.RequestContext, name?: string): Excel.Worksheet {
  return name ? ctx.workbook.worksheets.getItem(name) : ctx.workbook.worksheets.getActiveWorksheet();
}

function stripSheet(address: string): string {
  return address.includes("!") ? address.split("!")[1] : address;
}

function num(v: unknown): number | null {
  if (typeof v === "number") return v;
  if (typeof v === "string" && v.trim() && !isNaN(Number(v))) return Number(v);
  return null;
}

function key(v: unknown): string {
  return String(v ?? "").trim().toLowerCase();
}

/** Deterministic PRNG so a sample can be reproduced from its seed. */
function rng(seed: number) {
  let s = seed >>> 0 || 1;
  return () => {
    s ^= s << 13;
    s ^= s >>> 17;
    s ^= s << 5;
    return ((s >>> 0) % 1000000) / 1000000;
  };
}

async function readGrid(
  ctx: Excel.RequestContext,
  sheet: string | undefined,
  address: string | undefined
): Promise<{ name: string; address: string; values: unknown[][] }> {
  const ws = sheetOf(ctx, sheet);
  ws.load("name");
  const range = address ? ws.getRange(address) : ws.getUsedRange(true);
  range.load(["values", "address"]);
  await ctx.sync();
  return { name: ws.name, address: stripSheet(range.address), values: range.values as unknown[][] };
}

export async function reconcile(args: {
  sheetA?: string;
  addressA: string;
  sheetB?: string;
  addressB: string;
  keyColumn?: number;
  amountColumn?: number;
  hasHeaders?: boolean;
  tolerance?: number;
  maxReported?: number;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const a = await readGrid(ctx, args.sheetA, args.addressA);
    const b = await readGrid(ctx, args.sheetB, args.addressB);

    const kc = args.keyColumn ?? 0;
    const ac = args.amountColumn ?? 1;
    const tol = args.tolerance ?? 0.01;
    const cap = args.maxReported ?? 40;
    const skip = args.hasHeaders === false ? 0 : 1;

    const index = (rows: unknown[][]) => {
      const m = new Map<string, { amount: number; row: number }>();
      rows.slice(skip).forEach((r, i) => {
        const k = key(r[kc]);
        if (!k) return;
        const amt = num(r[ac]) ?? 0;
        const seen = m.get(k);
        m.set(k, { amount: (seen?.amount ?? 0) + amt, row: i + skip + 1 });
      });
      return m;
    };

    const ma = index(a.values);
    const mb = index(b.values);

    const onlyA: string[] = [];
    const onlyB: string[] = [];
    const diff: string[] = [];
    let matched = 0;
    let diffTotal = 0;

    for (const [k, va] of ma) {
      const vb = mb.get(k);
      if (!vb) {
        onlyA.push(k + " (" + va.amount + ", row " + va.row + ")");
        continue;
      }
      const d = va.amount - vb.amount;
      if (Math.abs(d) > tol) {
        diff.push(k + ": " + va.amount + " vs " + vb.amount + " = " + d.toFixed(2));
        diffTotal += d;
      } else matched++;
    }
    for (const [k, vb] of mb) if (!ma.has(k)) onlyB.push(k + " (" + vb.amount + ", row " + vb.row + ")");

    const sumA = [...ma.values()].reduce((s, v) => s + v.amount, 0);
    const sumB = [...mb.values()].reduce((s, v) => s + v.amount, 0);

    const cut = (arr: string[]) =>
      arr.length > cap ? arr.slice(0, cap).join("\n") + "\n… and " + (arr.length - cap) + " more" : arr.join("\n");

    return [
      "A = " + a.name + "!" + a.address + " (" + ma.size + " keys, total " + sumA.toFixed(2) + ")",
      "B = " + b.name + "!" + b.address + " (" + mb.size + " keys, total " + sumB.toFixed(2) + ")",
      "difference in totals: " + (sumA - sumB).toFixed(2),
      "matched within tolerance: " + matched,
      "",
      "ONLY IN A (" + onlyA.length + "):",
      cut(onlyA) || "none",
      "",
      "ONLY IN B (" + onlyB.length + "):",
      cut(onlyB) || "none",
      "",
      "AMOUNT DIFFERENCES (" + diff.length + ", net " + diffTotal.toFixed(2) + "):",
      cut(diff) || "none",
    ].join("\n");
  });
}

export async function tieOut(args: {
  sheet?: string;
  address?: string;
  debitColumn: number;
  creditColumn: number;
  hasHeaders?: boolean;
  tolerance?: number;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const g = await readGrid(ctx, args.sheet, args.address);
    const skip = args.hasHeaders === false ? 0 : 1;
    const tol = args.tolerance ?? 0.5;

    let dr = 0;
    let cr = 0;
    const blank: number[] = [];
    g.values.slice(skip).forEach((r, i) => {
      const d = num(r[args.debitColumn]) ?? 0;
      const c = num(r[args.creditColumn]) ?? 0;
      dr += d;
      cr += c;
      if (!d && !c) blank.push(i + skip + 1);
    });

    const diff = dr - cr;
    return [
      g.name + "!" + g.address,
      "debits " + dr.toFixed(2),
      "credits " + cr.toFixed(2),
      "difference " + diff.toFixed(2) + (Math.abs(diff) <= tol ? " (within tolerance, it ties)" : " (DOES NOT TIE)"),
      blank.length ? "rows with no debit and no credit: " + blank.slice(0, 30).join(", ") : "no empty rows",
    ].join("\n");
  });
}

export async function sampleRows(args: {
  sheet?: string;
  address?: string;
  size: number;
  method?: "random" | "systematic" | "highValue" | "monetaryUnit";
  amountColumn?: number;
  seed?: number;
  hasHeaders?: boolean;
  writeToSheet?: string;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const g = await readGrid(ctx, args.sheet, args.address);
    const skip = args.hasHeaders === false ? 0 : 1;
    const header = skip ? (g.values[0] as unknown[]) : [];
    const body = g.values.slice(skip);
    const n = Math.min(args.size ?? 25, body.length);
    const method = args.method ?? "random";
    const seed = args.seed ?? 42;
    const ac = args.amountColumn ?? 1;

    let picked: number[] = [];

    if (method === "systematic") {
      const step = body.length / n;
      for (let i = 0; i < n; i++) picked.push(Math.floor(i * step));
    } else if (method === "highValue") {
      picked = body
        .map((r, i) => ({ i, v: Math.abs(num(r[ac]) ?? 0) }))
        .sort((x, y) => y.v - x.v)
        .slice(0, n)
        .map((x) => x.i);
    } else if (method === "monetaryUnit") {
      // Probability proportional to size: bigger amounts are likelier to be hit.
      const weights = body.map((r) => Math.abs(num(r[ac]) ?? 0));
      const total = weights.reduce((a, b) => a + b, 0);
      if (total <= 0) return "monetaryUnit sampling needs a column of amounts.";
      const step = total / n;
      const rand = rng(seed)() * step;
      let cum = 0;
      let next = rand;
      for (let i = 0; i < body.length && picked.length < n; i++) {
        cum += weights[i];
        while (cum >= next && picked.length < n) {
          picked.push(i);
          next += step;
        }
      }
    } else {
      const rand = rng(seed);
      const pool = body.map((_, i) => i);
      for (let i = pool.length - 1; i > 0; i--) {
        const j = Math.floor(rand() * (i + 1));
        [pool[i], pool[j]] = [pool[j], pool[i]];
      }
      picked = pool.slice(0, n).sort((a, b) => a - b);
    }

    const rows = picked.map((i) => [i + skip + 1, ...(body[i] as unknown[])]);

    if (args.writeToSheet) {
      const existing = ctx.workbook.worksheets.getItemOrNullObject(args.writeToSheet);
      existing.load("isNullObject");
      await ctx.sync();
      const out = existing.isNullObject
        ? ctx.workbook.worksheets.add(args.writeToSheet)
        : existing;
      const grid = [["Source row", ...header.map((h) => String(h ?? ""))], ...rows] as (
        | string
        | number
        | boolean
      )[][];
      out
        .getRange("A1")
        .getResizedRange(grid.length - 1, Math.max(...grid.map((r) => r.length)) - 1)
        .values = grid;
      out.getRange("A1").getResizedRange(0, grid[0].length - 1).format.font.bold = true;
      out.activate();
      await ctx.sync();
      return (
        "picked " +
        picked.length +
        " items by " +
        method +
        " (seed " +
        seed +
        ") from " +
        body.length +
        " rows and wrote them to " +
        args.writeToSheet
      );
    }

    return (
      "picked " +
      picked.length +
      " of " +
      body.length +
      " rows by " +
      method +
      " (seed " +
      seed +
      "). Source rows: " +
      picked.map((i) => i + skip + 1).join(", ")
    );
  });
}

export async function duplicatesAndGaps(args: {
  sheet?: string;
  address?: string;
  column: number;
  hasHeaders?: boolean;
  checkGaps?: boolean;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const g = await readGrid(ctx, args.sheet, args.address);
    const skip = args.hasHeaders === false ? 0 : 1;
    const body = g.values.slice(skip);

    const seen = new Map<string, number[]>();
    const numbers: number[] = [];
    body.forEach((r, i) => {
      const raw = r[args.column];
      const k = key(raw);
      if (!k) return;
      seen.set(k, [...(seen.get(k) ?? []), i + skip + 1]);
      const n = num(raw);
      if (n !== null) numbers.push(n);
    });

    const dupes = [...seen.entries()].filter(([, rows]) => rows.length > 1);
    const lines: string[] = [
      g.name + "!" + g.address + ", column index " + args.column,
      "duplicates: " + dupes.length,
      ...dupes.slice(0, 30).map(([k, rows]) => "  " + k + " on rows " + rows.join(", ")),
    ];

    if (args.checkGaps !== false && numbers.length > 1) {
      const sorted = [...new Set(numbers)].sort((a, b) => a - b);
      const gaps: string[] = [];
      for (let i = 1; i < sorted.length; i++) {
        const step = sorted[i] - sorted[i - 1];
        if (step > 1) gaps.push(step === 2 ? String(sorted[i - 1] + 1) : sorted[i - 1] + 1 + "-" + (sorted[i] - 1));
      }
      lines.push(
        "range " + sorted[0] + " to " + sorted[sorted.length - 1],
        "missing numbers: " + (gaps.length ? gaps.slice(0, 40).join(", ") : "none")
      );
    }

    return lines.join("\n");
  });
}

export async function scanFormulas(args: { sheet?: string; address?: string }): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const range = args.address ? ws.getRange(args.address) : ws.getUsedRange(true);
    range.load(["formulas", "values", "address", "rowCount", "columnCount"]);
    await ctx.sync();

    const formulas = range.formulas as unknown[][];
    const values = range.values as unknown[][];
    const first = range.address.includes("!") ? range.address.split("!")[1] : range.address;
    const anchor = first.split(":")[0];
    const col0 = anchor.replace(/\d+/g, "");
    const row0 = parseInt(anchor.replace(/\D+/g, ""), 10) || 1;
    const colLetter = (i: number) => {
      let n = col0.split("").reduce((a, c) => a * 26 + (c.charCodeAt(0) - 64), 0) + i;
      let s = "";
      while (n > 0) {
        const r = (n - 1) % 26;
        s = String.fromCharCode(65 + r) + s;
        n = Math.floor((n - 1) / 26);
      }
      return s;
    };
    const addr = (r: number, c: number) => colLetter(c) + (row0 + r);

    const errors: string[] = [];
    const hardcoded: string[] = [];
    const inconsistent: string[] = [];

    values.forEach((row, r) =>
      row.forEach((v, c) => {
        if (typeof v === "string" && /^#(REF|VALUE|DIV\/0|N\/A|NAME\?|NUM|NULL)/.test(v))
          errors.push(addr(r, c) + " " + v);
      })
    );

    // A column that is mostly formulas but has a typed constant is the classic review finding.
    const cols = formulas[0]?.length ?? 0;
    for (let c = 0; c < cols; c++) {
      const cells = formulas.map((row) => String(row[c] ?? ""));
      const withFormula = cells.filter((f) => f.startsWith("=")).length;
      if (withFormula < 3) continue;
      const shapes = new Map<string, number>();
      cells.forEach((f, r) => {
        if (!f.startsWith("=")) {
          if (String(values[r][c] ?? "").trim() !== "") hardcoded.push(addr(r, c) + " = " + values[r][c]);
          return;
        }
        const shape = f.replace(/\d+/g, "#");
        shapes.set(shape, (shapes.get(shape) ?? 0) + 1);
      });
      if (shapes.size > 1) {
        const sorted = [...shapes.entries()].sort((a, b) => b[1] - a[1]);
        const odd = sorted.slice(1).reduce((s, x) => s + x[1], 0);
        inconsistent.push(
          "column " + colLetter(c) + ": " + shapes.size + " different formula patterns, " + odd + " odd cell(s)"
        );
      }
    }

    return [
      ws.name + "!" + first,
      "error cells (" + errors.length + "): " + (errors.slice(0, 25).join(", ") || "none"),
      "constants inside formula columns (" + hardcoded.length + "): " +
        (hardcoded.slice(0, 25).join(", ") || "none"),
      "inconsistent formulas: " + (inconsistent.join("; ") || "none"),
    ].join("\n");
  });
}

export async function ageing(args: {
  sheet?: string;
  address?: string;
  dateColumn: number;
  amountColumn: number;
  nameColumn?: number;
  asOf?: string;
  buckets?: number[];
  hasHeaders?: boolean;
  writeToSheet?: string;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const g = await readGrid(ctx, args.sheet, args.address);
    const skip = args.hasHeaders === false ? 0 : 1;
    const body = g.values.slice(skip);
    const cuts = args.buckets?.length ? [...args.buckets].sort((a, b) => a - b) : [30, 60, 90, 180];
    const asOf = args.asOf ? new Date(args.asOf) : new Date();

    const labels = [
      "0-" + cuts[0],
      ...cuts.slice(1).map((c, i) => cuts[i] + 1 + "-" + c),
      "over " + cuts[cuts.length - 1],
    ];

    const byName = new Map<string, number[]>();
    let skipped = 0;

    for (const r of body) {
      const amt = num(r[args.amountColumn]);
      if (amt === null) {
        skipped++;
        continue;
      }
      // Excel serial dates count days from 1899-12-30.
      const raw = r[args.dateColumn];
      const d =
        typeof raw === "number"
          ? new Date(Date.UTC(1899, 11, 30) + raw * 86400000)
          : new Date(String(raw));
      if (isNaN(d.getTime())) {
        skipped++;
        continue;
      }
      const days = Math.floor((asOf.getTime() - d.getTime()) / 86400000);
      let bucket = cuts.findIndex((c) => days <= c);
      if (bucket === -1) bucket = cuts.length;

      const name = args.nameColumn !== undefined ? String(r[args.nameColumn] ?? "(blank)") : "All";
      const row = byName.get(name) ?? new Array(labels.length).fill(0);
      row[bucket] += amt;
      byName.set(name, row);
    }

    const grid: (string | number)[][] = [
      [args.nameColumn !== undefined ? "Name" : "Total", ...labels, "Total"],
      ...[...byName.entries()]
        .sort((a, b) => b[1].reduce((x, y) => x + y, 0) - a[1].reduce((x, y) => x + y, 0))
        .map(([name, row]) => [name, ...row.map((v) => Number(v.toFixed(2))), Number(row.reduce((a, b) => a + b, 0).toFixed(2))]),
    ];

    if (args.writeToSheet) {
      const existing = ctx.workbook.worksheets.getItemOrNullObject(args.writeToSheet);
      existing.load("isNullObject");
      await ctx.sync();
      const out = existing.isNullObject ? ctx.workbook.worksheets.add(args.writeToSheet) : existing;
      const target = out.getRange("A1").getResizedRange(grid.length - 1, grid[0].length - 1);
      target.values = grid as (string | number | boolean)[][];
      out.getRange("A1").getResizedRange(0, grid[0].length - 1).format.font.bold = true;
      target.format.autofitColumns();
      out.activate();
      await ctx.sync();
      return "ageing written to " + args.writeToSheet + " (" + (grid.length - 1) + " rows, " + skipped + " skipped)";
    }

    return grid.map((r) => r.join("\t")).join("\n") + (skipped ? "\n(" + skipped + " rows skipped)" : "");
  });
}
