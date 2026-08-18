import type { UndoEntry } from "../types";

const stack: UndoEntry[] = [];
const MAX = 20;

/** Snapshot formulas + number formats of a range so an AI write can be reverted. */
export async function snapshot(
  ctx: Excel.RequestContext,
  sheet: string,
  address: string,
  label: string
): Promise<void> {
  const range = ctx.workbook.worksheets.getItem(sheet).getRange(address);
  range.load(["formulas", "numberFormat", "address"]);
  await ctx.sync();
  const entry: UndoEntry = {
    label,
    at: Date.now(),
    cells: [
      {
        sheet,
        address: range.address.includes("!") ? range.address.split("!")[1] : range.address,
        formulas: range.formulas as unknown[][],
        numberFormat: range.numberFormat as unknown as string[][],
      },
    ],
  };
  stack.push(entry);
  if (stack.length > MAX) stack.shift();
}

export function peekUndo(): UndoEntry | null {
  return stack.length ? stack[stack.length - 1] : null;
}

export function undoDepth(): number {
  return stack.length;
}

/** Revert the most recent AI write. Returns a human label, or null when nothing to undo. */
export async function undoLast(): Promise<string | null> {
  const entry = stack.pop();
  if (!entry) return null;
  await Excel.run(async (ctx) => {
    for (const cell of entry.cells) {
      const range = ctx.workbook.worksheets.getItem(cell.sheet).getRange(cell.address);
      range.formulas = cell.formulas as (string | number | boolean)[][];
      range.numberFormat = cell.numberFormat as unknown as string[][];
    }
    await ctx.sync();
  });
  return entry.label;
}
