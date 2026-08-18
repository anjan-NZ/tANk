import type { ToolDef } from "./tools";
import * as X from "../office/excelExtra";

const sheetProp = {
  sheet: { type: "string", description: "Sheet name. Omit to use the active sheet." },
};

export const EXCEL_EXTRA_TOOLS: ToolDef[] = [
  {
    name: "clear_range",
    description: "Clear cell contents, formatting, or both.",
    parameters: {
      type: "object",
      properties: {
        ...sheetProp,
        address: { type: "string" },
        what: { type: "string", enum: ["contents", "formats", "all"] },
      },
      required: ["address"],
    },
  },
  {
    name: "insert_rows_columns",
    description: "Insert blank rows or columns, shifting the existing cells down or right.",
    parameters: {
      type: "object",
      properties: {
        ...sheetProp,
        address: { type: "string", description: 'e.g. "5:7" for three rows, "C:C" for one column' },
        what: { type: "string", enum: ["rows", "columns"] },
      },
      required: ["address", "what"],
    },
  },
  {
    name: "delete_rows_columns",
    description: "Delete rows or columns and close the gap.",
    parameters: {
      type: "object",
      properties: {
        ...sheetProp,
        address: { type: "string", description: 'e.g. "5:7" or "C:D"' },
        what: { type: "string", enum: ["rows", "columns"] },
      },
      required: ["address", "what"],
    },
  },
  {
    name: "merge_cells",
    description: "Merge a range into one cell, or unmerge it again.",
    parameters: {
      type: "object",
      properties: {
        ...sheetProp,
        address: { type: "string" },
        unmerge: { type: "boolean" },
        across: { type: "boolean", description: "Merge each row separately rather than the whole block." },
      },
      required: ["address"],
    },
  },
  {
    name: "freeze_panes",
    description: "Freeze header rows or leading columns so they stay put when scrolling. Pass nothing to unfreeze.",
    parameters: {
      type: "object",
      properties: { ...sheetProp, rows: { type: "number" }, columns: { type: "number" } },
    },
  },
  {
    name: "create_table",
    description: "Turn a range into a real Excel table, which brings filter buttons and banded rows.",
    parameters: {
      type: "object",
      properties: {
        ...sheetProp,
        address: { type: "string" },
        hasHeaders: { type: "boolean" },
        name: { type: "string" },
        style: { type: "string", description: 'e.g. "TableStyleMedium2"' },
      },
      required: ["address"],
    },
  },
  {
    name: "list_tables",
    description: "List the tables in the workbook, or on one sheet, with their ranges.",
    parameters: { type: "object", properties: { ...sheetProp } },
  },
  {
    name: "filter_range",
    description:
      "Apply an autofilter to a range, optionally keeping only certain values in one column. Pass clear to remove the filter.",
    parameters: {
      type: "object",
      properties: {
        ...sheetProp,
        address: { type: "string" },
        columnIndex: { type: "number", description: "0 based column inside the range." },
        values: { type: "array", items: { type: "string" } },
        clear: { type: "boolean" },
      },
      required: ["address"],
    },
  },
  {
    name: "remove_duplicates",
    description: "Remove duplicate rows from a range, judged on the columns you name.",
    parameters: {
      type: "object",
      properties: {
        ...sheetProp,
        address: { type: "string" },
        columns: { type: "array", items: { type: "number" }, description: "0 based column indexes inside the range." },
        hasHeaders: { type: "boolean" },
      },
      required: ["address", "columns"],
    },
  },
  {
    name: "copy_range",
    description: "Copy cells somewhere else, keeping everything or only values, formulas or formats.",
    parameters: {
      type: "object",
      properties: {
        ...sheetProp,
        from: { type: "string" },
        to: { type: "string", description: "Top-left cell of the destination." },
        toSheet: { type: "string" },
        what: { type: "string", enum: ["all", "formulas", "values", "formats"] },
      },
      required: ["from", "to"],
    },
  },
  {
    name: "set_sizes",
    description: "Set column width or row height in points, or autofit to the contents.",
    parameters: {
      type: "object",
      properties: {
        ...sheetProp,
        address: { type: "string" },
        columnWidth: { type: "number" },
        rowHeight: { type: "number" },
        autofit: { type: "boolean" },
      },
      required: ["address"],
    },
  },
  {
    name: "manage_sheet",
    description: "Rename, delete, activate, hide or show a worksheet.",
    parameters: {
      type: "object",
      properties: {
        action: { type: "string", enum: ["rename", "delete", "activate", "hide", "show"] },
        sheet: { type: "string" },
        newName: { type: "string" },
      },
      required: ["action", "sheet"],
    },
  },
  {
    name: "replace_in_range",
    description: "Find and replace text across a range, or the whole used range of a sheet.",
    parameters: {
      type: "object",
      properties: {
        ...sheetProp,
        address: { type: "string" },
        find: { type: "string" },
        replace: { type: "string" },
        matchCase: { type: "boolean" },
      },
      required: ["find", "replace"],
    },
  },
  {
    name: "describe_range",
    description:
      "Count, sum, average, median, min and max of the numbers in a range. Use this instead of reading thousands of rows just to total them.",
    parameters: {
      type: "object",
      properties: { ...sheetProp, address: { type: "string" } },
      required: ["address"],
    },
  },
];

export async function execExcelExtraTool(name: string, args: any): Promise<string | null> {
  switch (name) {
    case "clear_range":
      return X.clearRange(args);
    case "insert_rows_columns":
      return X.insertRowsColumns(args);
    case "delete_rows_columns":
      return X.deleteRowsColumns(args);
    case "merge_cells":
      return X.mergeCells(args);
    case "freeze_panes":
      return X.freezePanes(args);
    case "create_table":
      return X.createTable(args);
    case "list_tables":
      return X.listTables({ sheet: args.sheet });
    case "filter_range":
      return X.autoFilter(args);
    case "remove_duplicates":
      return X.removeDuplicates(args);
    case "copy_range":
      return X.copyRange(args);
    case "set_sizes":
      return X.setSizes(args);
    case "manage_sheet":
      return X.manageSheet(args);
    case "replace_in_range":
      return X.replaceInRange(args);
    case "describe_range":
      return X.describeRange(args);
    default:
      return null;
  }
}

/** Extra Excel tools that destroy or move data, so the pane asks first. */
export const EXCEL_EXTRA_WRITING_TOOLS = new Set([
  "clear_range",
  "delete_rows_columns",
  "remove_duplicates",
  "copy_range",
  "replace_in_range",
  "merge_cells",
]);
