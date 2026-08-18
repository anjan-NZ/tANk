import type { Scope } from "../types";
import { scopeLabel } from "../office/scope";
import {
  addConditionalFormat,
  clearConditionalFormats,
  listConditionalFormats,
} from "../office/conditional";
import {
  addSheet,
  probeTarget,
  findText,
  formatRange,
  insertChart,
  listSheets,
  readRange,
  setFormula,
  sortRange,
  writeRange,
} from "../office/excel";

export interface JsonSchema {
  type: "object";
  properties: Record<string, unknown>;
  required?: string[];
}

export interface ToolDef {
  name: string;
  description: string;
  parameters: JsonSchema;
}

const sheetProp = {
  sheet: { type: "string", description: "Sheet name. Omit to use the active sheet." },
};

const EXCEL_TOOLS: ToolDef[] = [
  {
    name: "ask_user",
    description:
      "Ask the user a short question and wait for the answer. Use this when the working scope is unknown or ambiguous, or before a wide or destructive write. Offer options when there is a small set of sensible answers.",
    parameters: {
      type: "object",
      properties: {
        question: { type: "string" },
        options: {
          type: "array",
          items: { type: "string" },
          description: "Optional quick-reply choices shown as buttons.",
        },
      },
      required: ["question"],
    },
  },
  {
    name: "set_scope",
    description:
      "Set the working scope for the rest of the conversation, after the target is agreed. Accepts several ranges so work can span multiple sheets.",
    parameters: {
      type: "object",
      properties: {
        ranges: {
          type: "array",
          items: {
            type: "object",
            properties: {
              sheet: { type: "string" },
              address: { type: "string", description: "A1 style, e.g. A1:F120" },
            },
            required: ["sheet", "address"],
          },
        },
      },
      required: ["ranges"],
    },
  },
  {
    name: "list_sheets",
    description:
      "List every sheet in the workbook with its used range, size and first-row headers. Cheap; call it when you need to find where data lives.",
    parameters: { type: "object", properties: {} },
  },
  {
    name: "read_range",
    description:
      "Read cell contents. Returns tab separated rows. Omit address to read the whole used range of the sheet. Long ranges are truncated; read in slices if needed.",
    parameters: {
      type: "object",
      properties: {
        ...sheetProp,
        address: { type: "string", description: "A1 style range. Omit for the whole used range." },
        mode: { type: "string", enum: ["values", "formulas", "both"] },
      },
    },
  },
  {
    name: "write_range",
    description:
      "Write a block of values. The address is the top-left anchor; the block is sized from the values you pass. Snapshotted so the user can undo.",
    parameters: {
      type: "object",
      properties: {
        ...sheetProp,
        address: { type: "string", description: "Top-left anchor cell, e.g. H1" },
        values: {
          type: "array",
          description: "Rows of cell values.",
          items: { type: "array", items: {} },
        },
      },
      required: ["address", "values"],
    },
  },
  {
    name: "set_formula",
    description:
      "Put a formula in a range. Write the formula as it would be typed in the top-left cell; relative references are adjusted down the range automatically.",
    parameters: {
      type: "object",
      properties: {
        ...sheetProp,
        address: { type: "string", description: "Target range, e.g. G2:G120" },
        formula: { type: "string", description: "e.g. =IF(F2>0,F2*0.13,0)" },
        fill: {
          type: "boolean",
          description: "Default true: adjust references per row. False writes the identical formula everywhere.",
        },
      },
      required: ["address", "formula"],
    },
  },
  {
    name: "format_range",
    description: "Apply formatting: bold, italic, colours, number format, wrap, alignment, borders, autofit.",
    parameters: {
      type: "object",
      properties: {
        ...sheetProp,
        address: { type: "string" },
        bold: { type: "boolean" },
        italic: { type: "boolean" },
        fillColor: { type: "string", description: "Hex, e.g. #FFF2CC" },
        fontColor: { type: "string", description: "Hex, e.g. #C00000" },
        numberFormat: { type: "string", description: 'e.g. "#,##0.00" or "0.0%"' },
        wrapText: { type: "boolean" },
        horizontalAlignment: { type: "string", enum: ["Left", "Center", "Right"] },
        borders: { type: "string", enum: ["all", "bottom", "none"] },
        autofitColumns: { type: "boolean" },
      },
      required: ["address"],
    },
  },
  {
    name: "add_conditional_format",
    description:
      "Add a real Excel conditional formatting rule to a range, so it keeps reacting as the numbers change. Use this whenever the user asks to highlight or flag cells by a condition. Do NOT use format_range for that: a plain fill never updates.",
    parameters: {
      type: "object",
      properties: {
        ...sheetProp,
        address: { type: "string", description: "Range the rule covers, e.g. G5:G100" },
        type: {
          type: "string",
          enum: [
            "cellValue",
            "formula",
            "textContains",
            "topBottom",
            "duplicates",
            "colorScale",
            "dataBar",
            "iconSet",
          ],
          description:
            "cellValue for numeric comparisons, formula for anything custom, textContains for words, topBottom for the top or bottom N, duplicates, colorScale, dataBar, iconSet.",
        },
        operator: {
          type: "string",
          description:
            "For cellValue: GreaterThan, GreaterThanOrEqual, LessThan, LessThanOrEqual, EqualTo, NotEqualTo, Between, NotBetween. For textContains: Contains, NotContains, BeginsWith, EndsWith.",
        },
        value1: { type: "string", description: "Threshold, e.g. 10. Also the low end of Between." },
        value2: { type: "string", description: "High end of Between." },
        text: { type: "string", description: "The word to look for with textContains." },
        formula: {
          type: "string",
          description:
            'Custom rule written for the top-left cell of the range, e.g. "=AND($G5>10,$H5<>\"\")". Excel shifts it across the rest.',
        },
        rank: { type: "number", description: "N for topBottom." },
        rankType: {
          type: "string",
          enum: ["TopItems", "BottomItems", "TopPercent", "BottomPercent"],
        },
        iconStyle: {
          type: "string",
          description: "e.g. ThreeTrafficLights1, ThreeArrows, FiveRatings.",
        },
        style: {
          type: "object",
          description: "How matching cells should look. Defaults to the usual amber highlight.",
          properties: {
            fillColor: { type: "string", description: "Hex, e.g. #FFEB9C" },
            fontColor: { type: "string", description: "Hex, e.g. #9C5700" },
            bold: { type: "boolean" },
            italic: { type: "boolean" },
          },
        },
      },
      required: ["address", "type"],
    },
  },
  {
    name: "list_conditional_formats",
    description: "List the conditional formatting rules already on a range.",
    parameters: {
      type: "object",
      properties: { ...sheetProp, address: { type: "string" } },
      required: ["address"],
    },
  },
  {
    name: "clear_conditional_formats",
    description: "Remove every conditional formatting rule from a range.",
    parameters: {
      type: "object",
      properties: { ...sheetProp, address: { type: "string" } },
      required: ["address"],
    },
  },
  {
    name: "add_sheet",
    description: "Create a new worksheet.",
    parameters: {
      type: "object",
      properties: { name: { type: "string" }, activate: { type: "boolean" } },
      required: ["name"],
    },
  },
  {
    name: "find",
    description: "Search text across one sheet or the whole workbook. Returns matching addresses.",
    parameters: {
      type: "object",
      properties: { query: { type: "string" }, ...sheetProp },
      required: ["query"],
    },
  },
  {
    name: "insert_chart",
    description: "Insert a chart built from a data range.",
    parameters: {
      type: "object",
      properties: {
        ...sheetProp,
        dataRange: { type: "string" },
        chartType: {
          type: "string",
          description: "ColumnClustered, Line, Pie, BarClustered, XYScatter ...",
        },
        title: { type: "string" },
        anchorCell: { type: "string", description: "Top-left cell to place the chart, e.g. J2" },
      },
      required: ["dataRange"],
    },
  },
  {
    name: "sort_range",
    description: "Sort a range by one column index (0 based, relative to the range).",
    parameters: {
      type: "object",
      properties: {
        ...sheetProp,
        address: { type: "string" },
        columnIndex: { type: "number" },
        ascending: { type: "boolean" },
        hasHeaders: { type: "boolean" },
      },
      required: ["address", "columnIndex"],
    },
  },
];

import { detectHost, type HostApp } from "../office/host";
import { EXCEL_EXTRA_TOOLS, EXCEL_EXTRA_WRITING_TOOLS, execExcelExtraTool } from "./toolsExcelExtra";
import { WORD_TOOLS, WORD_WRITING_TOOLS, execWordTool } from "./toolsWord";
import { PPT_TOOLS, PPT_WRITING_TOOLS, execPptTool } from "./toolsPpt";
import { SCRIPT_WRITING_TOOLS, execScriptTool, scriptTools } from "./toolsScript";

/** Tool list depends on the host app. */
export function toolsFor(host: HostApp = detectHost(), lean = false): ToolDef[] {
  const ask = EXCEL_TOOLS.filter((t) => t.name === "ask_user");
  if (host === "word") return [...ask, ...WORD_TOOLS, ...scriptTools(host)];
  if (host === "powerpoint") return [...ask, ...PPT_TOOLS, ...scriptTools(host)];
  // Lean mode drops the convenience wrappers; run_office_script still reaches all of it.
  return lean
    ? [...EXCEL_TOOLS, ...scriptTools(host)]
    : [...EXCEL_TOOLS, ...EXCEL_EXTRA_TOOLS, ...scriptTools(host)];
}

export interface ToolDeps {
  maxRows: number;
  ask: (question: string, options: string[]) => Promise<string>;
  setScope: (scope: Scope) => void;
  /** "ask" pauses before anything already written gets replaced */
  editMode: "ask" | "auto";
  confirmEdit: (summary: string) => Promise<boolean>;
}

type Args = Record<string, any>;

/** Wording for the confirmation prompt. */
function describeChange(name: string, args: Args): string {
  const where = args.address
    ? (args.sheet ? args.sheet + "!" : "") + args.address
    : args.sheet ?? "";
  switch (name) {
    case "clear_range":
      return "Clear " + (args.what ?? "contents") + " in " + where + "?";
    case "delete_rows_columns":
      return "Delete " + args.what + " " + where + "? Everything below shifts up, so undo cannot put it back.";
    case "remove_duplicates":
      return "Remove duplicate rows from " + where + "? Undo cannot put the rows back.";
    case "copy_range":
      return "Copy " + args.from + " over " + (args.toSheet ? args.toSheet + "!" : "") + args.to + "?";
    case "replace_in_range":
      return 'Replace "' + args.find + '" with "' + args.replace + '" in ' + (where || "this sheet") + "?";
    case "manage_sheet":
      return args.action === "delete"
        ? "Delete the sheet " + args.sheet + " and everything on it? Undo does not cover this one."
        : "Rename the sheet " + args.sheet + " to " + args.newName + "?";
    case "audit_check":
      return (
        "Write the " + args.action + " result onto the sheet " + args.writeToSheet +
        "? Anything already there from A1 down is replaced, and undo does not cover it."
      );
    case "merge_cells":
      return (args.unmerge ? "Unmerge " : "Merge ") + where + "?";
    case "replace_selection":
      return "Replace the selected text?";
    case "find_replace":
      return 'Replace "' + args.find + '" with "' + args.replace + '" through the document?';
    case "insert_text":
      return "Insert a paragraph into the document?";
    case "insert_list":
      return "Insert a list of " + (args.items?.length ?? 0) + " items?";
    case "insert_table":
      return "Insert a table of " + (args.rows?.length ?? 0) + " rows?";
    case "delete_slide":
      return "Delete slide " + args.slide + "?";
    case "delete_shape":
      return "Delete that shape from slide " + args.slide + "?";
    case "set_shape_text":
      return "Replace the text of a shape on slide " + args.slide + "?";
    case "add_bullet_slide":
      return 'Add a slide titled "' + args.title + '"?';
    case "add_text_box":
      return "Add a text box to slide " + args.slide + "?";
    case "run_office_script":
      return (
        (args.purpose ? String(args.purpose) : "Run a script on this file") +
        "? Undo does not cover this one, so read it before saying yes.\n" +
        String(args.code ?? "").slice(0, 220) +
        (String(args.code ?? "").length > 220 ? " …" : "")
      );
    case "workbook_setup":
      return "Apply " + args.action + " to " + (args.sheet ? args.sheet + "!" : "") + (args.address ?? "this sheet") + "?";
    default:
      return "Go ahead with " + name + "?";
  }
}

/** Returns a refusal string when the user declines, or null to go ahead. */
async function guardOverwrite(
  deps: ToolDeps,
  sheet: string | undefined,
  address: string,
  rows: number | undefined,
  cols: number | undefined,
  verb: string
): Promise<string | null> {
  if (deps.editMode !== "ask") return null;
  const probe = await probeTarget({ sheet, address, rows, cols });
  if (!probe.occupied) return null; // empty cells, nothing to lose
  const ok = await deps.confirmEdit(
    verb +
      " " +
      probe.sheet +
      "!" +
      probe.address +
      "? " +
      probe.occupied +
      " of " +
      probe.cells +
      " cells there already have something in them."
  );
  return ok ? null : "The user said no, so nothing was changed.";
}

/**
 * Models often hand back numbers as strings (Gemini always does, its schema has no "any").
 * Convert only when the round trip is lossless, so account codes like "0012" stay text.
 */
function coerceCell(v: unknown): unknown {
  if (typeof v !== "string") return v;
  const t = v.trim();
  if (!t || !/^-?\d+(\.\d+)?$/.test(t)) return v;
  const n = Number(t);
  return String(n) === t ? n : v;
}

const NEEDS_PERMISSION = new Set([
  ...EXCEL_EXTRA_WRITING_TOOLS,
  ...WORD_WRITING_TOOLS,
  ...PPT_WRITING_TOOLS,
  ...SCRIPT_WRITING_TOOLS,
]);

/** Confirmed whatever editMode says: Auto waives overwrites, not running model-written code. */
const ALWAYS_CONFIRM = new Set(["run_office_script"]);

function needsPermission(name: string, args: Args): boolean {
  if (ALWAYS_CONFIRM.has(name)) return true;
  // audit_check only writes when asked to
  if (name === "audit_check") return typeof args.writeToSheet === "string" && !!args.writeToSheet;
  if (name === "manage_sheet") return args.action === "delete" || args.action === "rename";
  return NEEDS_PERMISSION.has(name);
}

/** Returns a refusal string when the user declines, or null to go ahead. */
async function gate(name: string, args: Args, deps: ToolDeps): Promise<string | null> {
  if (!ALWAYS_CONFIRM.has(name) && deps.editMode !== "ask") return null;
  if (!needsPermission(name, args)) return null;
  const ok = await deps.confirmEdit(describeChange(name, args));
  return ok ? null : "The user said no, so nothing was changed.";
}

/** These fetch a URL or call out of Excel, so a written formula never carries one. */
const UNSAFE_FUNCTIONS = /\b(WEBSERVICE|RTD|CALL|REGISTER(?:\.ID)?|EXEC|IMAGE)\s*\(/i;

export function unsafeFormula(v: unknown): string | null {
  const m = String(v ?? "").match(UNSAFE_FUNCTIONS);
  if (!m) return null;
  return (
    "Refused: " +
    m[1].toUpperCase() +
    "() can send the contents of this file out over the network, so tANk will not write it."
  );
}

export async function execTool(name: string, args: Args, deps: ToolDeps): Promise<string> {
  const host = detectHost();

  if (host === "word" || host === "powerpoint") {
    if (name === "ask_user")
      return deps.ask(String(args.question ?? ""), Array.isArray(args.options) ? args.options : []);

    const refused = await gate(name, args, deps);
    if (refused) return refused;

    const scripted = await execScriptTool(name, args);
    if (scripted !== null) return scripted;

    const result =
      host === "word" ? await execWordTool(name, args) : await execPptTool(name, args);
    return result ?? "That tool is not available in " + host + ".";
  }

  const badFormula =
    unsafeFormula(args.formula) ?? (name === "replace_in_range" ? unsafeFormula(args.replace) : null);
  if (badFormula) return badFormula;

  const refused = await gate(name, args, deps);
  if (refused) return refused;

  const handled = (await execExcelExtraTool(name, args)) ?? (await execScriptTool(name, args));
  if (handled !== null && handled !== undefined) return handled;

  switch (name) {
    case "ask_user":
      return deps.ask(String(args.question ?? ""), Array.isArray(args.options) ? args.options : []);

    case "set_scope": {
      const ranges = (args.ranges ?? []) as Array<{ sheet: string; address: string }>;
      if (!ranges.length) return "no ranges given";
      const scope: Scope = { ranges, label: scopeLabel(ranges), source: "user" };
      deps.setScope(scope);
      return "scope set to " + scope.label;
    }

    case "list_sheets": {
      const sheets = await listSheets();
      return sheets
        .map((s) =>
          s.usedRange
            ? `${s.name} | used ${s.usedRange} (${s.rows}r x ${s.cols}c) | headers: ${s.headers
                .slice(0, 12)
                .join(", ")}`
            : `${s.name} | empty`
        )
        .join("\n");
    }

    case "read_range": {
      const res = await readRange({
        sheet: args.sheet,
        address: args.address,
        mode: args.mode,
        maxRows: deps.maxRows,
      });
      const head = `${res.sheet}!${res.address} ${res.rows}r x ${res.cols}c${
        res.truncated ? ` (TRUNCATED to first ${deps.maxRows} rows)` : ""
      }`;
      const body = [
        res.values ? "VALUES\n" + res.values : "",
        res.formulas ? "FORMULAS\n" + res.formulas : "",
      ]
        .filter(Boolean)
        .join("\n\n");
      return head + "\n" + body;
    }

    case "write_range": {
      const values = (args.values as unknown[][]).map((row) => row.map(coerceCell));
      for (const row of values) {
        for (const cell of row) {
          const bad = unsafeFormula(cell);
          if (bad) return bad;
        }
      }
      const rows = values.length;
      const cols = rows ? Math.max(...values.map((r) => r.length)) : 0;
      const blocked = await guardOverwrite(deps, args.sheet, args.address, rows, cols, "Write");
      if (blocked) return blocked;
      return writeRange({ sheet: args.sheet, address: args.address, values });
    }

    case "set_formula": {
      const blocked = await guardOverwrite(deps, args.sheet, args.address, undefined, undefined, "Put a formula in");
      if (blocked) return blocked;
      return setFormula({
        sheet: args.sheet,
        address: args.address,
        formula: args.formula,
        fill: args.fill,
      });
    }

    case "format_range":
      return formatRange(args as any);

    case "add_conditional_format":
      return addConditionalFormat(args as any);

    case "list_conditional_formats":
      return listConditionalFormats({ sheet: args.sheet, address: args.address });

    case "clear_conditional_formats":
      return clearConditionalFormats({ sheet: args.sheet, address: args.address });

    case "add_sheet":
      return addSheet({ name: args.name, activate: args.activate });

    case "find":
      return findText({ query: args.query, sheet: args.sheet });

    case "insert_chart":
      return insertChart(args as any);

    case "sort_range": {
      if (deps.editMode === "ask") {
        const ok = await deps.confirmEdit(
          "Reorder the rows in " + (args.sheet ? args.sheet + "!" : "") + args.address + "?"
        );
        if (!ok) return "The user said no, nothing was sorted.";
      }
      return sortRange(args as any);
    }

    default:
      return "unknown tool: " + name;
  }
}
