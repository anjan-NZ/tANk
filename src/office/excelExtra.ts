/** The rest of the everyday Excel jobs: structure, tables, filters, cleanup. */

function sheetOf(ctx: Excel.RequestContext, name?: string): Excel.Worksheet {
  return name ? ctx.workbook.worksheets.getItem(name) : ctx.workbook.worksheets.getActiveWorksheet();
}

export async function clearRange(args: {
  sheet?: string;
  address: string;
  what?: "contents" | "formats" | "all";
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const map = {
      contents: Excel.ClearApplyTo.contents,
      formats: Excel.ClearApplyTo.formats,
      all: Excel.ClearApplyTo.all,
    };
    ws.getRange(args.address).clear(map[args.what ?? "contents"]);
    await ctx.sync();
    return "cleared " + (args.what ?? "contents") + " in " + ws.name + "!" + args.address;
  });
}

export async function insertRowsColumns(args: {
  sheet?: string;
  address: string;
  what: "rows" | "columns";
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const range = ws.getRange(args.address);
    range.insert(args.what === "rows" ? Excel.InsertShiftDirection.down : Excel.InsertShiftDirection.right);
    await ctx.sync();
    return "inserted " + args.what + " at " + ws.name + "!" + args.address;
  });
}

export async function deleteRowsColumns(args: {
  sheet?: string;
  address: string;
  what: "rows" | "columns";
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const range = ws.getRange(args.address);
    range.delete(args.what === "rows" ? Excel.DeleteShiftDirection.up : Excel.DeleteShiftDirection.left);
    await ctx.sync();
    return "deleted " + args.what + " at " + ws.name + "!" + args.address;
  });
}

export async function mergeCells(args: {
  sheet?: string;
  address: string;
  unmerge?: boolean;
  across?: boolean;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const range = ws.getRange(args.address);
    if (args.unmerge) range.unmerge();
    else range.merge(args.across ?? false);
    await ctx.sync();
    return (args.unmerge ? "unmerged " : "merged ") + ws.name + "!" + args.address;
  });
}

export async function freezePanes(args: { sheet?: string; rows?: number; columns?: number }): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    if (!args.rows && !args.columns) {
      ws.freezePanes.freezeAt(null as unknown as Excel.Range);
      await ctx.sync();
      return "unfroze the panes on " + ws.name;
    }
    if (args.rows) ws.freezePanes.freezeRows(args.rows);
    if (args.columns) ws.freezePanes.freezeColumns(args.columns);
    await ctx.sync();
    return "froze " + (args.rows ?? 0) + " row(s) and " + (args.columns ?? 0) + " column(s) on " + ws.name;
  });
}

export async function createTable(args: {
  sheet?: string;
  address: string;
  hasHeaders?: boolean;
  name?: string;
  style?: string;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const table = ws.tables.add(args.address, args.hasHeaders ?? true);
    if (args.name) table.name = args.name;
    if (args.style) table.style = args.style;
    table.load("name");
    await ctx.sync();
    return 'made "' + table.name + '" a table on ' + ws.name + "!" + args.address;
  });
}

export async function listTables(args: { sheet?: string }): Promise<string> {
  return Excel.run(async (ctx) => {
    const tables = args.sheet ? sheetOf(ctx, args.sheet).tables : ctx.workbook.tables;
    tables.load("items/name,items/id");
    await ctx.sync();
    if (!tables.items.length) return "no tables";
    const ranges = tables.items.map((t) => {
      const r = t.getRange();
      r.load("address");
      return { name: t.name, r };
    });
    await ctx.sync();
    return ranges.map((x) => x.name + " on " + x.r.address).join("\n");
  });
}

export async function autoFilter(args: {
  sheet?: string;
  address: string;
  columnIndex?: number;
  values?: string[];
  clear?: boolean;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");

    if (args.clear) {
      ws.autoFilter.remove();
      await ctx.sync();
      return "removed the filter on " + ws.name;
    }

    if (args.columnIndex !== undefined && args.values?.length) {
      ws.autoFilter.apply(ws.getRange(args.address), args.columnIndex, {
        filterOn: Excel.FilterOn.values,
        values: args.values,
      });
    } else {
      ws.autoFilter.apply(ws.getRange(args.address));
    }
    await ctx.sync();
    return "applied a filter to " + ws.name + "!" + args.address;
  });
}

export async function removeDuplicates(args: {
  sheet?: string;
  address: string;
  columns: number[];
  hasHeaders?: boolean;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const result = ws
      .getRange(args.address)
      .removeDuplicates(args.columns ?? [0], args.hasHeaders ?? true);
    result.load(["removed", "uniqueRemaining"]);
    await ctx.sync();
    return "removed " + result.removed + " duplicate row(s), " + result.uniqueRemaining + " left";
  });
}

export async function copyRange(args: {
  sheet?: string;
  from: string;
  to: string;
  toSheet?: string;
  what?: "all" | "formulas" | "values" | "formats";
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const fromWs = sheetOf(ctx, args.sheet);
    const toWs = args.toSheet ? ctx.workbook.worksheets.getItem(args.toSheet) : fromWs;
    fromWs.load("name");
    toWs.load("name");
    const map = {
      all: Excel.RangeCopyType.all,
      formulas: Excel.RangeCopyType.formulas,
      values: Excel.RangeCopyType.values,
      formats: Excel.RangeCopyType.formats,
    };
    toWs.getRange(args.to).copyFrom(fromWs.getRange(args.from), map[args.what ?? "all"]);
    await ctx.sync();
    return "copied " + fromWs.name + "!" + args.from + " to " + toWs.name + "!" + args.to;
  });
}

export async function setSizes(args: {
  sheet?: string;
  address: string;
  columnWidth?: number;
  rowHeight?: number;
  autofit?: boolean;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const range = ws.getRange(args.address);
    if (args.autofit) {
      range.format.autofitColumns();
      range.format.autofitRows();
    }
    if (args.columnWidth) range.format.columnWidth = args.columnWidth;
    if (args.rowHeight) range.format.rowHeight = args.rowHeight;
    await ctx.sync();
    return "resized " + ws.name + "!" + args.address;
  });
}

export async function manageSheet(args: {
  action: "rename" | "delete" | "activate" | "hide" | "show";
  sheet: string;
  newName?: string;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(args.sheet);
    switch (args.action) {
      case "rename":
        if (!args.newName) return "no new name given";
        ws.name = args.newName;
        break;
      case "delete":
        ws.delete();
        break;
      case "activate":
        ws.activate();
        break;
      case "hide":
        ws.visibility = Excel.SheetVisibility.hidden;
        break;
      case "show":
        ws.visibility = Excel.SheetVisibility.visible;
        break;
    }
    await ctx.sync();
    return args.action + " done on sheet " + args.sheet;
  });
}

export async function replaceInRange(args: {
  sheet?: string;
  address?: string;
  find: string;
  replace: string;
  matchCase?: boolean;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const range = args.address ? ws.getRange(args.address) : ws.getUsedRange(true);
    const count = range.replaceAll(args.find, args.replace, {
      completeMatch: false,
      matchCase: args.matchCase ?? false,
    });
    await ctx.sync();
    return "replaced " + count.value + ' occurrence(s) of "' + args.find + '" on ' + ws.name;
  });
}

export async function describeRange(args: { sheet?: string; address: string }): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const range = ws.getRange(args.address);
    range.load(["values", "rowCount", "columnCount", "address"]);
    await ctx.sync();

    const nums = (range.values as unknown[][])
      .flat()
      .filter((v) => typeof v === "number") as number[];
    if (!nums.length) return "No numbers in " + ws.name + "!" + args.address + ".";

    const sum = nums.reduce((a, b) => a + b, 0);
    const sorted = [...nums].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    const median = sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;

    return [
      ws.name + "!" + range.address.split("!").pop(),
      "count " + nums.length,
      "sum " + sum,
      "average " + (sum / nums.length).toFixed(4),
      "median " + median,
      "min " + sorted[0],
      "max " + sorted[sorted.length - 1],
    ].join(" | ");
  });
}
