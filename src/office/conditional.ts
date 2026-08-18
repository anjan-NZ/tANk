/** Conditional formatting rules, as opposed to a static fill. */

function sheetOf(ctx: Excel.RequestContext, name?: string): Excel.Worksheet {
  return name ? ctx.workbook.worksheets.getItem(name) : ctx.workbook.worksheets.getActiveWorksheet();
}

function stripSheet(address: string): string {
  return address.includes("!") ? address.split("!")[1] : address;
}

export interface CfStyle {
  fillColor?: string;
  fontColor?: string;
  bold?: boolean;
  italic?: boolean;
}

export interface CfArgs {
  sheet?: string;
  address: string;
  type:
    | "cellValue"
    | "formula"
    | "textContains"
    | "topBottom"
    | "duplicates"
    | "colorScale"
    | "dataBar"
    | "iconSet";
  operator?: string;
  value1?: string | number;
  value2?: string | number;
  text?: string;
  formula?: string;
  rank?: number;
  rankType?: "topItems" | "bottomItems" | "topPercent" | "bottomPercent";
  iconStyle?: string;
  style?: CfStyle;
}

function applyStyle(format: Excel.ConditionalRangeFormat, style?: CfStyle): void {
  const s = style ?? { fillColor: "#FFEB9C", fontColor: "#9C5700" };
  if (s.fillColor) format.fill.color = s.fillColor;
  if (s.fontColor) format.font.color = s.fontColor;
  if (s.bold !== undefined) format.font.bold = s.bold;
  if (s.italic !== undefined) format.font.italic = s.italic;
}

/** Excel wants rule values as formula strings: 10 -> "=10". */
function asFormula(v: string | number | undefined): string {
  if (v === undefined || v === null || v === "") return "=0";
  const s = String(v).trim();
  return s.startsWith("=") ? s : "=" + s;
}

export async function addConditionalFormat(args: CfArgs): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const range = ws.getRange(args.address);
    range.load("address");
    await ctx.sync();

    const T = Excel.ConditionalFormatType;
    let describe = args.type;

    switch (args.type) {
      case "cellValue": {
        const cf = range.conditionalFormats.add(T.cellValue);
        applyStyle(cf.cellValue.format, args.style);
        cf.cellValue.rule = {
          formula1: asFormula(args.value1),
          formula2: args.value2 !== undefined ? asFormula(args.value2) : undefined,
          operator: (args.operator ??
            "GreaterThan") as unknown as Excel.ConditionalCellValueOperator,
        };
        describe = ("cell value " + (args.operator ?? "GreaterThan") + " " + args.value1) as CfArgs["type"];
        break;
      }

      case "formula": {
        const cf = range.conditionalFormats.add(T.custom);
        applyStyle(cf.custom.format, args.style);
        // The formula is written for the top-left cell; Excel shifts it across the range.
        cf.custom.rule.formula = args.formula ?? "=TRUE";
        break;
      }

      case "textContains": {
        const cf = range.conditionalFormats.add(T.containsText);
        applyStyle(cf.textComparison.format, args.style);
        cf.textComparison.rule = {
          operator: (args.operator ?? "Contains") as unknown as Excel.ConditionalTextOperator,
          text: args.text ?? "",
        };
        break;
      }

      case "topBottom": {
        const cf = range.conditionalFormats.add(T.topBottom);
        applyStyle(cf.topBottom.format, args.style);
        cf.topBottom.rule = {
          rank: args.rank ?? 10,
          type: (args.rankType ??
            "TopItems") as unknown as Excel.ConditionalTopBottomCriterionType,
        };
        break;
      }

      case "duplicates": {
        const cf = range.conditionalFormats.add(T.presetCriteria);
        applyStyle(cf.preset.format, args.style);
        cf.preset.rule = {
          criterion: "DuplicateValues" as unknown as Excel.ConditionalFormatPresetCriterion,
        };
        break;
      }

      case "colorScale": {
        const cf = range.conditionalFormats.add(T.colorScale);
        cf.colorScale.criteria = {
          minimum: { formula: null as unknown as string, type: "LowestValue" as unknown as Excel.ConditionalFormatColorCriterionType, color: "#F8696B" },
          midpoint: { formula: "=50", type: "Percentile" as unknown as Excel.ConditionalFormatColorCriterionType, color: "#FFEB84" },
          maximum: { formula: null as unknown as string, type: "HighestValue" as unknown as Excel.ConditionalFormatColorCriterionType, color: "#63BE7B" },
        };
        break;
      }

      case "dataBar": {
        const cf = range.conditionalFormats.add(T.dataBar);
        cf.dataBar.positiveFormat.fillColor = args.style?.fillColor ?? "#638EC6";
        break;
      }

      case "iconSet": {
        const cf = range.conditionalFormats.add(T.iconSet);
        cf.iconSet.style = (args.iconStyle ??
          "ThreeTrafficLights1") as unknown as Excel.IconSet;
        break;
      }

      default:
        return "Unknown conditional format type: " + args.type;
    }

    await ctx.sync();
    return "added a " + describe + " rule on " + ws.name + "!" + stripSheet(range.address);
  });
}

export async function listConditionalFormats(args: { sheet?: string; address: string }): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const cfs = ws.getRange(args.address).conditionalFormats;
    cfs.load("items/type,items/id");
    await ctx.sync();
    if (!cfs.items.length) return "no conditional formatting rules on " + args.address;
    return cfs.items.map((c, i) => i + 1 + ". " + c.type).join("\n");
  });
}

export async function clearConditionalFormats(args: { sheet?: string; address: string }): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, args.sheet);
    ws.load("name");
    const range = ws.getRange(args.address);
    range.conditionalFormats.clearAll();
    await ctx.sync();
    return "cleared conditional formatting on " + ws.name + "!" + args.address;
  });
}
