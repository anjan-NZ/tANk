/** Word side of tANk: read the document, edit the selection, restructure text. */

const MAX_CHARS = 12000;

function clip(text: string): { text: string; truncated: boolean } {
  if (text.length <= MAX_CHARS) return { text, truncated: false };
  return { text: text.slice(0, MAX_CHARS), truncated: true };
}

export async function getSelection(): Promise<string> {
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.load(["text", "style"]);
    await ctx.sync();
    if (!sel.text.trim()) return "Nothing is selected.";
    const { text, truncated } = clip(sel.text);
    return (
      "Selected text" +
      (sel.style ? " (style: " + sel.style + ")" : "") +
      (truncated ? ", clipped" : "") +
      ":\n" +
      text
    );
  });
}

export async function getDocumentText(args: { fromParagraph?: number; count?: number }): Promise<string> {
  return Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs;
    paras.load("items/text,items/style,items/styleBuiltIn");
    await ctx.sync();

    const start = args.fromParagraph ?? 0;
    const count = args.count ?? 80;
    const slice = paras.items.slice(start, start + count);
    const body = slice
      .map((p, i) => start + i + ": " + (p.style ? "[" + p.style + "] " : "") + p.text)
      .join("\n");
    const { text, truncated } = clip(body);

    return (
      "Paragraphs " +
      start +
      " to " +
      Math.min(start + count, paras.items.length) +
      " of " +
      paras.items.length +
      (truncated ? " (clipped)" : "") +
      ":\n" +
      text
    );
  });
}

export async function getOutline(): Promise<string> {
  return Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs;
    paras.load("items/text,items/style");
    await ctx.sync();
    const heads = paras.items
      .map((p, i) => ({ i, text: p.text, style: p.style }))
      .filter((p) => /heading|title/i.test(p.style ?? ""));
    if (!heads.length) return "No headings in this document.";
    return heads.map((h) => h.i + ": [" + h.style + "] " + h.text).join("\n");
  });
}

export async function replaceSelection(args: { text: string }): Promise<string> {
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.insertText(args.text, Word.InsertLocation.replace);
    await ctx.sync();
    return "replaced the selection with " + args.text.length + " characters";
  });
}

export async function insertText(args: {
  text: string;
  where?: "afterSelection" | "beforeSelection" | "endOfDocument" | "startOfDocument";
  style?: string;
}): Promise<string> {
  return Word.run(async (ctx) => {
    const where = args.where ?? "afterSelection";
    let para: Word.Paragraph;

    if (where === "endOfDocument") {
      para = ctx.document.body.insertParagraph(args.text, Word.InsertLocation.end);
    } else if (where === "startOfDocument") {
      para = ctx.document.body.insertParagraph(args.text, Word.InsertLocation.start);
    } else {
      const sel = ctx.document.getSelection();
      para = sel.insertParagraph(
        args.text,
        where === "beforeSelection" ? Word.InsertLocation.before : Word.InsertLocation.after
      );
    }

    if (args.style) para.style = args.style;
    await ctx.sync();
    return "inserted a paragraph at " + where;
  });
}

export async function formatSelection(args: {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  highlightColor?: string;
  size?: number;
  font?: string;
  alignment?: "Left" | "Centered" | "Right" | "Justified";
  style?: string;
}): Promise<string> {
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    if (args.style) sel.style = args.style;
    if (args.bold !== undefined) sel.font.bold = args.bold;
    if (args.italic !== undefined) sel.font.italic = args.italic;
    if (args.underline !== undefined)
      sel.font.underline = args.underline
        ? Word.UnderlineType.single
        : Word.UnderlineType.none;
    if (args.color) sel.font.color = args.color;
    if (args.highlightColor) sel.font.highlightColor = args.highlightColor;
    if (args.size) sel.font.size = args.size;
    if (args.font) sel.font.name = args.font;
    if (args.alignment)
      sel.paragraphs.getFirst().alignment = args.alignment as unknown as Word.Alignment;
    await ctx.sync();
    return "formatted the selection";
  });
}

export async function findReplace(args: {
  find: string;
  replace: string;
  matchCase?: boolean;
  wholeWord?: boolean;
}): Promise<string> {
  return Word.run(async (ctx) => {
    const results = ctx.document.body.search(args.find, {
      matchCase: args.matchCase ?? false,
      matchWholeWord: args.wholeWord ?? false,
    });
    results.load("items/text");
    await ctx.sync();

    results.items.forEach((r) => r.insertText(args.replace, Word.InsertLocation.replace));
    await ctx.sync();
    return "replaced " + results.items.length + ' occurrence(s) of "' + args.find + '"';
  });
}

export async function insertList(args: { items: string[]; ordered?: boolean }): Promise<string> {
  return Word.run(async (ctx) => {
    if (!args.items?.length) return "no items given";
    const sel = ctx.document.getSelection();
    const first = sel.insertParagraph(args.items[0], Word.InsertLocation.after);
    const list = first.startNewList();
    if (args.ordered) list.setLevelNumbering(0, Word.ListNumbering.arabic);
    else list.setLevelBullet(0, Word.ListBullet.solid);
    await ctx.sync();

    let previous = first;
    for (const item of args.items.slice(1)) {
      previous = previous.insertParagraph(item, Word.InsertLocation.after);
      previous.attachToList(list.id, 0);
    }
    await ctx.sync();
    return "inserted a list of " + args.items.length + " items";
  });
}

export async function insertTable(args: { rows: string[][]; header?: boolean }): Promise<string> {
  return Word.run(async (ctx) => {
    const rows = args.rows ?? [];
    if (!rows.length) return "no rows given";
    const cols = Math.max(...rows.map((r) => r.length));
    const grid = rows.map((r) => Array.from({ length: cols }, (_, i) => r[i] ?? ""));

    const sel = ctx.document.getSelection();
    const table = sel.insertTable(grid.length, cols, Word.InsertLocation.after, grid);
    table.headerRowCount = args.header === false ? 0 : 1;
    // The enum was renamed Style -> BuiltInStyleName in WordApi 1.5; the string is what goes over the wire.
    table.styleBuiltIn = "GridTable4_Accent1" as unknown as Word.Table["styleBuiltIn"];
    await ctx.sync();
    return "inserted a " + grid.length + " by " + cols + " table";
  });
}

export async function insertPageBreak(): Promise<string> {
  return Word.run(async (ctx) => {
    ctx.document.getSelection().insertBreak(Word.BreakType.page, Word.InsertLocation.after);
    await ctx.sync();
    return "inserted a page break";
  });
}

export async function addComment(args: { text: string }): Promise<string> {
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    // Comments need WordApi 1.4; older builds throw and the caller reports it.
    sel.insertComment(args.text);
    await ctx.sync();
    return "left a comment on the selection";
  });
}

export async function documentStats(): Promise<string> {
  return Word.run(async (ctx) => {
    const body = ctx.document.body;
    body.load("text");
    const paras = body.paragraphs;
    paras.load("items/text");
    await ctx.sync();
    const text = body.text ?? "";
    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    return (
      paras.items.length + " paragraphs, " + words + " words, " + text.length + " characters"
    );
  });
}
