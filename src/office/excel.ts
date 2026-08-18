import { snapshot } from "./undo";

/** Values come back as TSV: far cheaper in tokens than JSON matrices. */
function toTsv(rows: unknown[][]): string {
  return rows
    .map((r) => r.map((c) => (c === null || c === undefined ? "" : String(c))).join("\t"))
    .join("\n");
}

function sheetOf(ctx: Excel.RequestContext, name?: string): Excel.Worksheet {
  return name ? ctx.workbook.worksheets.getItem(name) : ctx.workbook.worksheets.getActiveWorksheet();
}

function stripSheet(address: string): string {
  return address.includes("!") ? address.split("!")[1] : address;
}

export interface SheetInfo {
  name: string;
  usedRange: string | null;
  rows: number;
  cols: number;
  headers: string[];
}

export async function listSheets(): Promise<SheetInfo[]> {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();

    const probes = sheets.items.map((s) => {
      const used = s.getUsedRangeOrNullObject(true);
      used.load(["address", "rowCount", "columnCount", "isNullObject"]);
      return { name: s.name, used };
    });
    await ctx.sync();

    const heads = probes.map((p) => {
      if (p.used.isNullObject) return null;
      const first = p.used.getRow(0);
      first.load("values");
      return first;
    });
    await ctx.sync();

    return probes.map((p, i) => ({
      name: p.name,
      usedRange: p.used.isNullObject ? null : stripSheet(p.used.address),
      rows: p.used.isNullObject ? 0 : p.used.rowCount,
      cols: p.used.isNullObject ? 0 : p.used.columnCount,
      headers: heads[i] ? (heads[i]!.values[0] as unknown[]).map((v) => String(v ?? "")) : [],
    }));
  });
}

export interface ReadResult {
  sheet: string;
  address: string;
  rows: number;
  cols: number;
  truncated: boolean;
  values?: string;
  formulas?: string;
}

export async function readRange(args: {
  sheet?: string;
  address?: string;
  mode?: "values" | "formulas" | "both";
  maxRows: number;
}): Promise<ReadResult> {
  const mode = args.mode ?? "values";
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const target = args.address ? ws.getRange(args.address) : ws.getUsedRange(true);
    target.load(["address", "rowCount", "columnCount"]);
    await ctx.sync();

    const truncated = target.rowCount > args.maxRows;
    const slice = truncated ? target.getResizedRange(args.maxRows - target.rowCount, 0) : target;
    if (mode !== "formulas") slice.load("values");
    if (mode !== "values") slice.load("formulas");
    await ctx.sync();

    return {
      sheet: ws.name,
      address: stripSheet(target.address),
      rows: target.rowCount,
      cols: target.columnCount,
      truncated,
      values: mode !== "formulas" ? toTsv(slice.values as unknown[][]) : undefined,
      formulas: mode !== "values" ? toTsv(slice.formulas as unknown[][]) : undefined,
    };
  });
}

export interface TargetProbe {
  sheet: string;
  address: string;
  cells: number;
  occupied: number;
}

/** What would this write hit? Used to ask before overwriting anything that already has data. */
export async function probeTarget(args: {
  sheet?: string;
  address: string;
  rows?: number;
  cols?: number;
}): Promise<TargetProbe> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const anchor = ws.getRange(args.address);
    anchor.load(["rowCount", "columnCount"]);
    await ctx.sync();

    const rows = args.rows ?? anchor.rowCount;
    const cols = args.cols ?? anchor.columnCount;
    const target =
      anchor.rowCount === rows && anchor.columnCount === cols
        ? anchor
        : anchor.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
    target.load(["address", "values", "rowCount", "columnCount"]);
    await ctx.sync();

    const flat = (target.values as unknown[][]).flat();
    return {
      sheet: ws.name,
      address: stripSheet(target.address),
      cells: target.rowCount * target.columnCount,
      occupied: flat.filter((v) => v !== "" && v !== null && v !== undefined).length,
    };
  });
}

export async function writeRange(args: {
  sheet?: string;
  address: string;
  values: unknown[][];
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    await ctx.sync();

    const rows = args.values.length;
    const cols = rows ? Math.max(...args.values.map((r) => r.length)) : 0;
    if (!rows || !cols) throw new Error("values is empty");
    // Normalise ragged rows: Office.js rejects a grid whose rows differ in length.
    const grid = args.values.map((r) =>
      Array.from({ length: cols }, (_, i) => (r[i] === undefined ? "" : r[i]))
    );

    const anchor = ws.getRange(args.address);
    anchor.load(["rowCount", "columnCount"]);
    await ctx.sync();

    const target =
      anchor.rowCount === rows && anchor.columnCount === cols
        ? anchor
        : anchor.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
    target.load("address");
    await ctx.sync();

    const addr = stripSheet(target.address);
    await snapshot(ctx, ws.name, addr, "write " + ws.name + "!" + addr);
    target.values = grid as (string | number | boolean)[][];
    await ctx.sync();
    return "wrote " + rows + "x" + cols + " to " + ws.name + "!" + addr;
  });
}

export async function setFormula(args: {
  sheet?: string;
  address: string;
  formula: string;
  fill?: boolean;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const target = ws.getRange(args.address);
    target.load(["address", "rowCount", "columnCount"]);
    await ctx.sync();

    const addr = stripSheet(target.address);
    await snapshot(ctx, ws.name, addr, "formula " + ws.name + "!" + addr);

    const single = target.rowCount === 1 && target.columnCount === 1;
    if (single || args.fill === false) {
      const grid = Array.from({ length: target.rowCount }, () =>
        Array.from({ length: target.columnCount }, () => args.formula)
      );
      target.formulas = grid;
      await ctx.sync();
    } else {
      // copyFrom adjusts relative references per cell; a bulk formulas assignment does not.
      const top = target.getCell(0, 0);
      top.formulas = [[args.formula]];
      await ctx.sync();
      target.copyFrom(top, Excel.RangeCopyType.formulas);
      await ctx.sync();
    }
    return "set formula on " + ws.name + "!" + addr;
  });
}

export interface FormatArgs {
  sheet?: string;
  address: string;
  bold?: boolean;
  italic?: boolean;
  fillColor?: string;
  fontColor?: string;
  numberFormat?: string;
  autofitColumns?: boolean;
  wrapText?: boolean;
  horizontalAlignment?: "Left" | "Center" | "Right";
  borders?: "all" | "bottom" | "none";
}

export async function formatRange(args: FormatArgs): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const target = ws.getRange(args.address);
    target.load(["address", "rowCount", "columnCount"]);
    await ctx.sync();

    const addr = stripSheet(target.address);
    await snapshot(ctx, ws.name, addr, "format " + ws.name + "!" + addr);

    const fmt = target.format;
    if (args.bold !== undefined) fmt.font.bold = args.bold;
    if (args.italic !== undefined) fmt.font.italic = args.italic;
    if (args.fontColor) fmt.font.color = args.fontColor;
    if (args.fillColor) fmt.fill.color = args.fillColor;
    if (args.wrapText !== undefined) fmt.wrapText = args.wrapText;
    if (args.horizontalAlignment) fmt.horizontalAlignment = args.horizontalAlignment;
    if (args.numberFormat) {
      const grid = Array.from({ length: target.rowCount }, () =>
        Array.from({ length: target.columnCount }, () => args.numberFormat as string)
      );
      target.numberFormat = grid;
    }
    if (args.borders && args.borders !== "none") {
      const edges =
        args.borders === "all"
          ? ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"]
          : ["EdgeBottom"];
      for (const e of edges) {
        const b = fmt.borders.getItem(e as Excel.BorderIndex);
        b.style = Excel.BorderLineStyle.continuous;
        b.weight = Excel.BorderWeight.thin;
      }
    }
    if (args.autofitColumns) fmt.autofitColumns();
    await ctx.sync();
    return "formatted " + ws.name + "!" + addr;
  });
}

export async function addSheet(args: { name: string; activate?: boolean }): Promise<string> {
  return Excel.run(async (ctx) => {
    const existing = ctx.workbook.worksheets.getItemOrNullObject(args.name);
    existing.load("isNullObject");
    await ctx.sync();
    if (!existing.isNullObject) return 'sheet "' + args.name + '" already exists';

    const ws = ctx.workbook.worksheets.add(args.name);
    if (args.activate !== false) ws.activate();
    await ctx.sync();
    return 'created sheet "' + args.name + '"';
  });
}

export async function findText(args: { query: string; sheet?: string }): Promise<string> {
  return Excel.run(async (ctx) => {
    const targets: Excel.Worksheet[] = [];
    if (args.sheet) {
      targets.push(ctx.workbook.worksheets.getItem(args.sheet));
    } else {
      const all = ctx.workbook.worksheets;
      all.load("items/name");
      await ctx.sync();
      targets.push(...all.items);
    }

    const hits = targets.map((ws) => {
      ws.load("name");
      // Worksheet.findAllOrNullObject needs ExcelApi 1.9; on older hosts this throws
      // and the caller surfaces the error to the model.
      const found = ws.findAllOrNullObject(args.query, {
        completeMatch: false,
        matchCase: false,
      });
      found.load(["areas/items/address", "isNullObject"]);
      return { ws, found };
    });
    await ctx.sync();

    const lines: string[] = [];
    for (const h of hits) {
      if (h.found.isNullObject) continue;
      const addrs = h.found.areas.items.map((a: Excel.Range) => stripSheet(a.address)).join(", ");
      if (addrs) lines.push(h.ws.name + ": " + addrs);
    }
    return lines.length ? lines.join("\n") : 'no match for "' + args.query + '"';
  });
}

export async function insertChart(args: {
  sheet?: string;
  dataRange: string;
  chartType?: string;
  title?: string;
  anchorCell?: string;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const data = ws.getRange(args.dataRange);
    const chart = ws.charts.add(
      (args.chartType ?? "ColumnClustered") as Excel.ChartType,
      data,
      Excel.ChartSeriesBy.auto
    );
    if (args.title) chart.title.text = args.title;
    if (args.anchorCell) chart.setPosition(args.anchorCell);
    await ctx.sync();
    return "chart added on " + ws.name + " from " + args.dataRange;
  });
}

export async function sortRange(args: {
  sheet?: string;
  address: string;
  columnIndex: number;
  ascending?: boolean;
  hasHeaders?: boolean;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const target = ws.getRange(args.address);
    target.sort.apply(
      [{ key: args.columnIndex, ascending: args.ascending ?? true }],
      false,
      args.hasHeaders ?? true
    );
    await ctx.sync();
    return "sorted " + ws.name + "!" + args.address + " by column index " + args.columnIndex;
  });
}
