import type { ToolDef } from "./tools";
import type { HostApp } from "../office/host";
import { evaluateFormula, runOfficeScript } from "../office/script";
import * as A from "../office/excelAudit";
import * as P from "../office/excelPro";

const EXAMPLE: Record<string, string> = {
  excel:
    'const s = context.workbook.worksheets.getActiveWorksheet(); const r = s.getRange("A1:C9"); ' +
    'r.load("values"); await context.sync(); return "read " + r.values.length + " rows";',
  word:
    'const b = context.document.body; b.load("text"); await context.sync(); ' +
    'return "the document has " + b.text.length + " characters";',
  powerpoint:
    'const s = context.presentation.slides; s.load("items/id"); await context.sync(); ' +
    'return s.items.length + " slides";',
};

/** Gives the model the whole Office JS API instead of a tool per feature. */
export function scriptTools(host: HostApp): ToolDef[] {
  const api =
    host === "word"
      ? "Word.run context: context.document (body, getSelection, sections, contentControls, properties, changeTrackingMode), paragraphs, tables, lists, comments, insertOoxml, insertFootnote, headers and footers, styles."
      : host === "powerpoint"
        ? "PowerPoint.run context: context.presentation (slides, slideMasters, layouts, tags), shapes, textFrame, textRange, fill, lineFormat."
        : "Excel.run context: context.workbook (worksheets, tables, names, comments, pivotTables, dataValidation, conditionalFormats, charts, protection, pageLayout, autoFilter, slicers, custom XML).";

  const tools: ToolDef[] = [
    {
      name: "run_office_script",
      description:
        "Run a short Office JavaScript snippet against the open file. This reaches anything the Office API can do, so use it whenever no named tool fits: pivot tables, data validation, named ranges, cell notes, grouping and outlining, sheet protection, page setup and print areas, chart formatting, track changes, tables of contents, footnotes, headers and footers, document properties, and so on. " +
        api +
        " Rules: the code is the body of an async function with `context` in scope; call `await context.sync()` after loading or writing; finish with `return` and a short plain-English string saying what happened; keep it under about 40 lines; never loop over thousands of cells one at a time, read or write whole ranges at once. Example: " +
        (EXAMPLE[host] ?? EXAMPLE.excel),
      parameters: {
        type: "object",
        properties: {
          code: { type: "string", description: "The function body to run." },
          purpose: { type: "string", description: "One short line on what this does, shown to the user." },
        },
        required: ["code"],
      },
    },
  ];

  if (host !== "word" && host !== "powerpoint") {
    tools.push(
      {
        name: "evaluate_formula",
        description:
          "Work out an Excel formula without leaving anything in the sheet. Gives you every built-in worksheet function as a calculator: XLOOKUP, SUMIFS, SUMPRODUCT, IRR, NPV, TEXTJOIN, FILTER, UNIQUE and the rest. Prefer this over reading a range and adding it up yourself.",
        parameters: {
          type: "object",
          properties: {
            formula: { type: "string", description: 'e.g. =SUMIFS(TB!D:D,TB!B:B,"Revenue")' },
            sheet: { type: "string", description: "Sheet whose context the formula is worked out in." },
            spill: { type: "boolean", description: "true when the formula returns an array, e.g. FILTER or UNIQUE." },
          },
          required: ["formula"],
        },
      },
      {
        name: "audit_check",
        description:
          "Ready-made audit routines over a range: reconcile two ranges by a key column, tie out debits against credits, pick a sample, hunt duplicates and missing numbers in a sequence, review formulas for errors and hardcoded constants, or age balances into buckets.",
        parameters: {
          type: "object",
          properties: {
            action: {
              type: "string",
              enum: ["reconcile", "tie_out", "sample", "duplicates_gaps", "scan_formulas", "ageing"],
            },
            sheet: { type: "string" },
            address: { type: "string", description: "Range to work on. Omit for the whole used range." },
            sheetB: { type: "string", description: "reconcile: the second sheet." },
            addressB: { type: "string", description: "reconcile: the second range." },
            keyColumn: { type: "number", description: "reconcile: 0 based column holding the key." },
            amountColumn: { type: "number", description: "0 based column holding the amount." },
            debitColumn: { type: "number", description: "tie_out: 0 based debit column." },
            creditColumn: { type: "number", description: "tie_out: 0 based credit column." },
            column: { type: "number", description: "duplicates_gaps: 0 based column to examine." },
            dateColumn: { type: "number", description: "ageing: 0 based date column." },
            nameColumn: { type: "number", description: "ageing: 0 based column to group by, e.g. customer." },
            asOf: { type: "string", description: "ageing: date to age against, e.g. 2026-03-31." },
            buckets: { type: "array", items: { type: "number" }, description: "ageing: day cut-offs, default 30,60,90,180." },
            size: { type: "number", description: "sample: how many items." },
            method: { type: "string", enum: ["random", "systematic", "highValue", "monetaryUnit"] },
            seed: { type: "number", description: "sample: keeps the selection reproducible." },
            tolerance: { type: "number" },
            hasHeaders: { type: "boolean", description: "Default true." },
            writeToSheet: { type: "string", description: "sample and ageing: write the result to this sheet." },
          },
          required: ["action"],
        },
      },
      {
        name: "workbook_setup",
        description:
          "Workbook furniture in one call: data validation and dropdowns, named ranges, cell notes, row and column grouping, sheet protection and cell locking, page setup and print area, pivot tables, and chart formatting.",
        parameters: {
          type: "object",
          properties: {
            action: {
              type: "string",
              enum: [
                "validation",
                "name",
                "add_note",
                "list_notes",
                "group",
                "protect",
                "lock_cells",
                "page_setup",
                "pivot",
                "format_chart",
              ],
            },
            sheet: { type: "string" },
            address: { type: "string" },
            cell: { type: "string", description: "add_note: the single cell." },
            text: { type: "string", description: "add_note: the note." },
            type: {
              type: "string",
              enum: ["list", "wholeNumber", "decimal", "date", "textLength", "clear"],
              description: "validation: which rule.",
            },
            values: { type: "array", items: { type: "string" }, description: "validation list entries." },
            min: { type: "string" },
            max: { type: "string" },
            message: { type: "string" },
            nameAction: { type: "string", enum: ["add", "list", "delete"], description: "name: what to do." },
            name: { type: "string" },
            what: { type: "string", enum: ["rows", "columns"], description: "group: which axis." },
            ungroup: { type: "boolean" },
            collapse: { type: "boolean" },
            unprotect: { type: "boolean" },
            password: { type: "string" },
            locked: { type: "boolean", description: "lock_cells: true to lock." },
            orientation: { type: "string", enum: ["Portrait", "Landscape"] },
            printArea: { type: "string" },
            fitToPagesWide: { type: "number" },
            repeatTopRows: { type: "string", description: 'e.g. "$1:$1"' },
            headerText: { type: "string" },
            footerText: { type: "string" },
            sourceAddress: { type: "string", description: "pivot: the source range." },
            destinationSheet: { type: "string", description: "pivot: sheet to build it on." },
            rows: { type: "array", items: { type: "string" }, description: "pivot: row fields, by header name." },
            columns: { type: "array", items: { type: "string" }, description: "pivot: column fields." },
            valueFields: {
              type: "array",
              items: { type: "string" },
              description: "pivot: fields to total.",
            },
            title: { type: "string", description: "format_chart: chart title." },
            xAxisTitle: { type: "string" },
            yAxisTitle: { type: "string" },
            legendPosition: { type: "string", enum: ["Top", "Bottom", "Left", "Right", "None"] },
            showDataLabels: { type: "boolean" },
          },
          required: ["action"],
        },
      }
    );
  }

  return tools;
}

export async function execScriptTool(name: string, args: any): Promise<string | null> {
  switch (name) {
    case "run_office_script":
      return runOfficeScript({ code: String(args.code ?? "") });

    case "evaluate_formula":
      return evaluateFormula({ formula: String(args.formula ?? ""), sheet: args.sheet, spill: args.spill });

    case "audit_check":
      switch (args.action) {
        case "reconcile":
          return A.reconcile({
            sheetA: args.sheet,
            addressA: args.address,
            sheetB: args.sheetB,
            addressB: args.addressB,
            keyColumn: args.keyColumn,
            amountColumn: args.amountColumn,
            hasHeaders: args.hasHeaders,
            tolerance: args.tolerance,
          });
        case "tie_out":
          return A.tieOut({
            sheet: args.sheet,
            address: args.address,
            debitColumn: args.debitColumn ?? 0,
            creditColumn: args.creditColumn ?? 1,
            hasHeaders: args.hasHeaders,
            tolerance: args.tolerance,
          });
        case "sample":
          return A.sampleRows({
            sheet: args.sheet,
            address: args.address,
            size: args.size ?? 25,
            method: args.method,
            amountColumn: args.amountColumn,
            seed: args.seed,
            hasHeaders: args.hasHeaders,
            writeToSheet: args.writeToSheet,
          });
        case "duplicates_gaps":
          return A.duplicatesAndGaps({
            sheet: args.sheet,
            address: args.address,
            column: args.column ?? 0,
            hasHeaders: args.hasHeaders,
          });
        case "scan_formulas":
          return A.scanFormulas({ sheet: args.sheet, address: args.address });
        case "ageing":
          return A.ageing({
            sheet: args.sheet,
            address: args.address,
            dateColumn: args.dateColumn ?? 0,
            amountColumn: args.amountColumn ?? 1,
            nameColumn: args.nameColumn,
            asOf: args.asOf,
            buckets: args.buckets,
            hasHeaders: args.hasHeaders,
            writeToSheet: args.writeToSheet,
          });
        default:
          return "Unknown audit action: " + args.action;
      }

    case "workbook_setup":
      switch (args.action) {
        case "validation":
          return P.setValidation({
            sheet: args.sheet,
            address: args.address,
            type: args.type ?? "list",
            values: args.values,
            min: args.min,
            max: args.max,
            message: args.message,
          });
        case "name":
          return P.manageName({
            action: args.nameAction ?? "list",
            name: args.name,
            sheet: args.sheet,
            address: args.address,
          });
        case "add_note":
          return P.addNote({ sheet: args.sheet, cell: args.cell ?? args.address, text: String(args.text ?? "") });
        case "list_notes":
          return P.listNotes({ sheet: args.sheet });
        case "group":
          return P.groupRange({
            sheet: args.sheet,
            address: args.address,
            what: args.what ?? "rows",
            ungroup: args.ungroup,
            collapse: args.collapse,
          });
        case "protect":
          return P.protectSheet({ sheet: args.sheet, unprotect: args.unprotect, password: args.password });
        case "lock_cells":
          return P.lockCells({ sheet: args.sheet, address: args.address, locked: args.locked ?? true });
        case "page_setup":
          return P.pageSetup(args);
        case "pivot":
          return P.createPivot({
            sourceSheet: args.sheet,
            sourceAddress: args.sourceAddress ?? args.address,
            destinationSheet: args.destinationSheet ?? "Pivot",
            rows: args.rows,
            columns: args.columns,
            values: (args.valueFields ?? []).map((f: string) => ({ field: f, summarizeBy: "Sum" })),
          });
        case "format_chart":
          return P.formatChart(args);
        default:
          return "Unknown setup action: " + args.action;
      }

    default:
      return null;
  }
}

/** These change the file, so the pane asks first. */
export const SCRIPT_WRITING_TOOLS = new Set(["run_office_script", "workbook_setup"]);
