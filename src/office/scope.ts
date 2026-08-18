import type { Scope } from "../types";

export interface SelectionInfo {
  sheet: string;
  address: string;
  rows: number;
  cols: number;
  /** true when the selection is a single cell (or a single empty cell) */
  single: boolean;
  empty: boolean;
}

function stripSheet(address: string): string {
  return address.includes("!") ? address.split("!")[1] : address;
}

export async function readSelection(): Promise<SelectionInfo> {
  return Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    const ws = range.worksheet;
    range.load(["address", "rowCount", "columnCount", "valueTypes", "values"]);
    ws.load("name");
    await ctx.sync();

    const single = range.rowCount === 1 && range.columnCount === 1;
    const flat = (range.values as unknown[][]).flat();
    const empty = flat.every((v) => v === "" || v === null || v === undefined);

    return {
      sheet: ws.name,
      address: stripSheet(range.address),
      rows: range.rowCount,
      cols: range.columnCount,
      single,
      empty,
    };
  });
}

/** Grow a single cell out to the contiguous block around it (Excel's Ctrl+* region). */
export async function surroundingRegion(sheet: string, address: string): Promise<SelectionInfo | null> {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheet);
    const region = ws.getRange(address).getSurroundingRegion();
    region.load(["address", "rowCount", "columnCount"]);
    ws.load("name");
    try {
      await ctx.sync();
    } catch {
      return null;
    }
    return {
      sheet: ws.name,
      address: stripSheet(region.address),
      rows: region.rowCount,
      cols: region.columnCount,
      single: region.rowCount === 1 && region.columnCount === 1,
      empty: false,
    };
  });
}

export function scopeFromSelection(sel: SelectionInfo): Scope {
  return {
    ranges: [{ sheet: sel.sheet, address: sel.address }],
    label: sel.sheet + "!" + sel.address,
    source: "selection",
  };
}

export function scopeLabel(ranges: Array<{ sheet: string; address: string }>): string {
  return ranges.map((r) => r.sheet + "!" + r.address).join(" + ");
}

/**
 * Decide what to work on before a turn starts.
 * - real multi-cell selection -> use it, no questions
 * - single cell inside a block -> propose the surrounding block for confirmation
 * - nothing useful selected    -> no scope; the model must ask
 */
export async function proposeScope(): Promise<{ scope: Scope | null; needsConfirm: boolean; proposal?: Scope }> {
  const sel = await readSelection();

  if (!sel.single && !(sel.rows === 1 && sel.cols === 1)) {
    return { scope: scopeFromSelection(sel), needsConfirm: false };
  }

  const region = await surroundingRegion(sel.sheet, sel.address);
  if (region && !(region.rows === 1 && region.cols === 1)) {
    return { scope: null, needsConfirm: true, proposal: scopeFromSelection(region) };
  }

  // A lone cell still counts. "put it here" needs an address, not a question.
  return { scope: scopeFromSelection(sel), needsConfirm: false };
}
