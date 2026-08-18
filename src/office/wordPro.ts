/** Word's own machinery: revisions, comments, tables, footnotes, headers, styles, properties. */

export async function setTrackChanges(args: { mode: "off" | "trackAll" | "trackMineOnly" }): Promise<string> {
  return Word.run(async (ctx) => {
    // ChangeTrackingMode needs WordApi 1.4.
    ctx.document.changeTrackingMode = args.mode as unknown as Word.ChangeTrackingMode;
    await ctx.sync();
    return args.mode === "off" ? "track changes turned off" : "track changes turned on (" + args.mode + ")";
  });
}

export async function listComments(): Promise<string> {
  return Word.run(async (ctx) => {
    const comments = ctx.document.body.getComments();
    comments.load("items/id,items/content,items/authorName,items/resolved,items/creationDate");
    await ctx.sync();
    if (!comments.items.length) return "no comments in this document";

    const ranges = comments.items.map((c) => {
      const r = c.getRange();
      r.load("text");
      return { c, r };
    });
    await ctx.sync();

    return ranges
      .map(
        (x, i) =>
          i +
          1 +
          ". " +
          x.c.authorName +
          (x.c.resolved ? " [resolved]" : "") +
          ': "' +
          x.c.content +
          '" on: ' +
          (x.r.text ?? "").slice(0, 80)
      )
      .join("\n");
  });
}

export async function resolveComment(args: { index: number; resolved?: boolean }): Promise<string> {
  return Word.run(async (ctx) => {
    const comments = ctx.document.body.getComments();
    comments.load("items/id,items/resolved");
    await ctx.sync();
    const c = comments.items[(args.index ?? 1) - 1];
    if (!c) return "there is no comment " + args.index;
    c.resolved = args.resolved ?? true;
    await ctx.sync();
    return "comment " + args.index + (args.resolved === false ? " reopened" : " marked resolved");
  });
}

export async function insertFootnote(args: { text: string }): Promise<string> {
  return Word.run(async (ctx) => {
    // Footnotes need WordApi 1.5.
    ctx.document.getSelection().insertFootnote(args.text);
    await ctx.sync();
    return "footnote added at the selection";
  });
}

export async function headerFooter(args: {
  which: "header" | "footer";
  text: string;
  type?: "Primary" | "FirstPage" | "EvenPages";
}): Promise<string> {
  return Word.run(async (ctx) => {
    const sections = ctx.document.sections;
    sections.load("items");
    await ctx.sync();
    const section = sections.items[0];
    const type = (args.type ?? "Primary") as unknown as Word.HeaderFooterType;
    const target = args.which === "footer" ? section.getFooter(type) : section.getHeader(type);
    target.clear();
    target.insertText(args.text, Word.InsertLocation.start);
    await ctx.sync();
    return args.which + " set to: " + args.text;
  });
}

export async function documentProperties(args: {
  title?: string;
  subject?: string;
  author?: string;
  keywords?: string;
  read?: boolean;
}): Promise<string> {
  return Word.run(async (ctx) => {
    const props = ctx.document.properties;
    if (args.read) {
      props.load(["title", "subject", "author", "keywords", "lastAuthor", "revisionNumber"]);
      await ctx.sync();
      return [
        "title: " + (props.title || "(none)"),
        "subject: " + (props.subject || "(none)"),
        "author: " + (props.author || "(none)"),
        "keywords: " + (props.keywords || "(none)"),
        "last saved by: " + (props.lastAuthor || "(unknown)"),
        "revision: " + props.revisionNumber,
      ].join("\n");
    }
    if (args.title) props.title = args.title;
    if (args.subject) props.subject = args.subject;
    if (args.author) props.author = args.author;
    if (args.keywords) props.keywords = args.keywords;
    await ctx.sync();
    return "document properties updated";
  });
}

export async function listTables(): Promise<string> {
  return Word.run(async (ctx) => {
    const tables = ctx.document.body.tables;
    tables.load("items/rowCount,items/headerRowCount,items/values");
    await ctx.sync();
    if (!tables.items.length) return "no tables in this document";
    return tables.items
      .map((t, i) => {
        const rows = (t.values as string[][]) ?? [];
        const preview = rows
          .slice(0, 3)
          .map((r) => r.join(" | "))
          .join("\n    ");
        return "Table " + (i + 1) + " (" + t.rowCount + " rows):\n    " + preview;
      })
      .join("\n");
  });
}

export async function setTableCell(args: {
  table: number;
  row: number;
  column: number;
  text: string;
}): Promise<string> {
  return Word.run(async (ctx) => {
    const tables = ctx.document.body.tables;
    tables.load("items/rowCount");
    await ctx.sync();
    const t = tables.items[(args.table ?? 1) - 1];
    if (!t) return "there is no table " + args.table;
    const cell = t.getCell(args.row, args.column);
    cell.value = args.text;
    await ctx.sync();
    return "table " + args.table + " cell (" + args.row + "," + args.column + ") set";
  });
}

export async function addTableRow(args: { table: number; values: string[] }): Promise<string> {
  return Word.run(async (ctx) => {
    const tables = ctx.document.body.tables;
    tables.load("items/rowCount");
    await ctx.sync();
    const t = tables.items[(args.table ?? 1) - 1];
    if (!t) return "there is no table " + args.table;
    t.addRows(Word.InsertLocation.end, 1, [args.values ?? []]);
    await ctx.sync();
    return "row added to table " + args.table;
  });
}

/** Word has no direct "insert TOC", so drop in the field as OOXML. */
export async function insertTableOfContents(): Promise<string> {
  return Word.run(async (ctx) => {
    const ooxml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships></pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr>
      <w:sdtContent><w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>
      <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>Right-click here and choose Update Field to build the table of contents.</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent></w:sdt>
    </w:body></w:document></pkg:xmlData>
  </pkg:part>
</pkg:package>`;
    ctx.document.getSelection().insertOoxml(ooxml, Word.InsertLocation.after);
    await ctx.sync();
    return "table of contents field inserted; Word fills it in on update";
  });
}

export async function replaceParagraph(args: { index: number; text: string; style?: string }): Promise<string> {
  return Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs;
    paras.load("items/text");
    await ctx.sync();
    const p = paras.items[args.index];
    if (!p) return "there is no paragraph " + args.index;
    p.insertText(args.text, Word.InsertLocation.replace);
    if (args.style) p.style = args.style;
    await ctx.sync();
    return "paragraph " + args.index + " replaced";
  });
}

export async function listStyles(): Promise<string> {
  return Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs;
    paras.load("items/style");
    await ctx.sync();
    const counts = new Map<string, number>();
    paras.items.forEach((p) => counts.set(p.style, (counts.get(p.style) ?? 0) + 1));
    return [...counts.entries()]
      .sort((a, b) => b[1] - a[1])
      .map(([style, n]) => style + " x" + n)
      .join("\n");
  });
}
