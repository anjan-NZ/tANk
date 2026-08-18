/** Workbook furniture: validation, names, notes, grouping, protection, printing, pivots, charts. */

function sheetOf(ctx: Excel.RequestContext, name?: string): Excel.Worksheet {
  return name ? ctx.workbook.worksheets.getItem(name) : ctx.workbook.worksheets.getActiveWorksheet();
}

export async function setValidation(args: {
  sheet?: string;
  address: string;
  type: "list" | "wholeNumber" | "decimal" | "date" | "textLength" | "clear";
  values?: string[];
  operator?: string;
  min?: number | string;
  max?: number | string;
  message?: string;
  errorMessage?: string;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const range = ws.getRange(args.address);

    if (args.type === "clear") {
      range.dataValidation.clear();
      await ctx.sync();
      return "cleared validation on " + ws.name + "!" + args.address;
    }

    const rule: Excel.DataValidationRule =
      args.type === "list"
        ? { list: { inCellDropDown: true, source: (args.values ?? []).join(",") } }
        : ({
            [args.type]: {
              formula1: String(args.min ?? 0),
              formula2: args.max !== undefined ? String(args.max) : undefined,
              operator: (args.operator ?? (args.max !== undefined ? "Between" : "GreaterThanOrEqualTo")) as any,
            },
          } as unknown as Excel.DataValidationRule);

    range.dataValidation.rule = rule;
    if (args.message)
      range.dataValidation.prompt = { message: args.message, showPrompt: true, title: "" };
    range.dataValidation.errorAlert = {
      message: args.errorMessage ?? "That value is not allowed here.",
      showAlert: true,
      style: Excel.DataValidationAlertStyle.stop,
      title: "Not allowed",
    };
    await ctx.sync();
    return "validation set on " + ws.name + "!" + args.address;
  });
}

export async function manageName(args: {
  action: "add" | "list" | "delete";
  name?: string;
  sheet?: string;
  address?: string;
  comment?: string;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    if (args.action === "list") {
      const names = ctx.workbook.names;
      names.load("items/name,items/formula,items/comment");
      await ctx.sync();
      if (!names.items.length) return "no named ranges";
      return names.items.map((n) => n.name + " -> " + n.formula + (n.comment ? "  // " + n.comment : "")).join("\n");
    }

    if (!args.name) return "no name given";

    if (args.action === "delete") {
      ctx.workbook.names.getItem(args.name).delete();
      await ctx.sync();
      return "deleted the name " + args.name;
    }

    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    await ctx.sync();
    const added = ctx.workbook.names.add(args.name, ws.getRange(args.address ?? "A1"));
    if (args.comment) added.comment = args.comment;
    await ctx.sync();
    return "named " + ws.name + "!" + args.address + " as " + args.name;
  });
}

export async function addNote(args: { sheet?: string; cell: string; text: string }): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const target = ws.getRange(args.cell);
    // Comments need ExcelApi 1.10, and a second comment on the same cell throws,
    // so find any existing one by its location and remove it first.
    const comments = ctx.workbook.comments;
    comments.load("items/id");
    await ctx.sync();

    const located = comments.items.map((c) => {
      const cell = c.getLocation();
      cell.load("address");
      return { c, cell };
    });
    await ctx.sync();

    const wanted = (ws.name + "!" + args.cell).replace(/\$/g, "").toUpperCase();
    const hit = located.find((x) => x.cell.address.replace(/\$/g, "").toUpperCase() === wanted);
    if (hit) hit.c.delete();

    ctx.workbook.comments.add(target, args.text);
    await ctx.sync();
    return "left a note on " + ws.name + "!" + args.cell;
  });
}

export async function listNotes(args: { sheet?: string }): Promise<string> {
  return Excel.run(async (ctx) => {
    const comments = ctx.workbook.comments;
    comments.load("items/content,items/authorName");
    await ctx.sync();
    if (!comments.items.length) return "no notes in this workbook";

    const located = comments.items.map((c) => {
      const cell = c.getLocation();
      cell.load("address");
      return { c, cell };
    });
    await ctx.sync();

    return located
      .filter((x) => !args.sheet || x.cell.address.startsWith(args.sheet + "!"))
      .map((x) => x.cell.address + " (" + x.c.authorName + "): " + x.c.content)
      .join("\n");
  });
}

export async function groupRange(args: {
  sheet?: string;
  address: string;
  what: "rows" | "columns";
  ungroup?: boolean;
  collapse?: boolean;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const range = ws.getRange(args.address);
    const option =
      args.what === "rows" ? Excel.GroupOption.byRows : Excel.GroupOption.byColumns;
    if (args.ungroup) range.ungroup(option);
    else {
      range.group(option);
      if (args.collapse) range.hideGroupDetails(option);
    }
    await ctx.sync();
    return (args.ungroup ? "ungrouped " : "grouped ") + args.what + " " + ws.name + "!" + args.address;
  });
}

export async function protectSheet(args: {
  sheet?: string;
  unprotect?: boolean;
  password?: string;
  allowFormatting?: boolean;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    if (args.unprotect) {
      ws.protection.unprotect(args.password);
      await ctx.sync();
      return "unprotected " + ws.name;
    }
    ws.protection.protect(
      {
        allowFormatCells: args.allowFormatting ?? false,
        allowFormatColumns: args.allowFormatting ?? false,
        allowFormatRows: args.allowFormatting ?? false,
        allowSort: true,
        allowAutoFilter: true,
      },
      args.password
    );
    await ctx.sync();
    return "protected " + ws.name + (args.password ? " with a password" : "");
  });
}

export async function lockCells(args: { sheet?: string; address: string; locked: boolean }): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    ws.getRange(args.address).format.protection.locked = args.locked;
    await ctx.sync();
    return (args.locked ? "locked " : "unlocked ") + ws.name + "!" + args.address + " (takes effect once the sheet is protected)";
  });
}

export async function pageSetup(args: {
  sheet?: string;
  orientation?: "Portrait" | "Landscape";
  printArea?: string;
  fitToPagesWide?: number;
  fitToPagesTall?: number;
  repeatTopRows?: string;
  headerText?: string;
  footerText?: string;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const ps = ws.pageLayout;

    if (args.orientation) ps.orientation = args.orientation as unknown as Excel.PageOrientation;
    if (args.printArea) ps.setPrintArea(args.printArea);
    if (args.repeatTopRows) ps.setPrintTitleRows(args.repeatTopRows);
    if (args.fitToPagesWide || args.fitToPagesTall) {
      ps.zoom = {
        horizontalFitToPages: args.fitToPagesWide ?? 1,
        verticalFitToPages: args.fitToPagesTall ?? 0,
      };
    }
    if (args.headerText) ps.headersFooters.defaultForAllPages.centerHeader = args.headerText;
    if (args.footerText) ps.headersFooters.defaultForAllPages.centerFooter = args.footerText;

    await ctx.sync();
    return "page setup updated on " + ws.name;
  });
}

export async function createPivot(args: {
  sourceSheet?: string;
  sourceAddress: string;
  destinationSheet: string;
  destinationCell?: string;
  rows?: string[];
  columns?: string[];
  values?: Array<{ field: string; summarizeBy?: string }>;
  name?: string;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const src = sheetOf(ctx, args.sourceSheet);
    src.load("name");

    const existing = ctx.workbook.worksheets.getItemOrNullObject(args.destinationSheet);
    existing.load("isNullObject");
    await ctx.sync();
    const dest = existing.isNullObject
      ? ctx.workbook.worksheets.add(args.destinationSheet)
      : existing;

    const pivot = ctx.workbook.pivotTables.add(
      args.name ?? "Pivot" + Date.now().toString().slice(-4),
      src.getRange(args.sourceAddress),
      dest.getRange(args.destinationCell ?? "A1")
    );
    await ctx.sync();

    for (const r of args.rows ?? []) pivot.rowHierarchies.add(pivot.hierarchies.getItem(r));
    for (const c of args.columns ?? []) pivot.columnHierarchies.add(pivot.hierarchies.getItem(c));
    for (const v of args.values ?? []) {
      const dh = pivot.dataHierarchies.add(pivot.hierarchies.getItem(v.field));
      if (v.summarizeBy)
        dh.summarizeBy = v.summarizeBy as unknown as Excel.AggregationFunction;
    }
    dest.activate();
    await ctx.sync();
    return "pivot built on " + args.destinationSheet + " from " + src.name + "!" + args.sourceAddress;
  });
}

export async function formatChart(args: {
  sheet?: string;
  chartName?: string;
  index?: number;
  title?: string;
  xAxisTitle?: string;
  yAxisTitle?: string;
  legendPosition?: "Top" | "Bottom" | "Left" | "Right" | "None";
  showDataLabels?: boolean;
  seriesColors?: string[];
  height?: number;
  width?: number;
}): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const charts = ws.charts;
    charts.load("items/name");
    await ctx.sync();
    if (!charts.items.length) return "no charts on " + ws.name;

    const chart = args.chartName
      ? charts.items.find((c) => c.name === args.chartName) ?? charts.items[0]
      : charts.items[args.index ?? 0] ?? charts.items[0];

    if (args.title) {
      chart.title.text = args.title;
      chart.title.visible = true;
    }
    if (args.xAxisTitle) {
      chart.axes.categoryAxis.title.text = args.xAxisTitle;
      chart.axes.categoryAxis.title.visible = true;
    }
    if (args.yAxisTitle) {
      chart.axes.valueAxis.title.text = args.yAxisTitle;
      chart.axes.valueAxis.title.visible = true;
    }
    if (args.legendPosition) {
      chart.legend.visible = args.legendPosition !== "None";
      if (args.legendPosition !== "None")
        chart.legend.position = args.legendPosition as unknown as Excel.ChartLegendPosition;
    }
    if (args.showDataLabels !== undefined) chart.dataLabels.showValue = args.showDataLabels;
    if (args.height) chart.height = args.height;
    if (args.width) chart.width = args.width;

    if (args.seriesColors?.length) {
      const series = chart.series;
      series.load("items/name");
      await ctx.sync();
      series.items.forEach((s, i) => {
        const colour = args.seriesColors![i % args.seriesColors!.length];
        s.format.fill.setSolidColor(colour);
      });
    }

    await ctx.sync();
    return "chart formatted on " + ws.name;
  });
}
