import type { ToolDef } from "./tools";
import * as W from "../office/word";
import * as WP from "../office/wordPro";

export const WORD_TOOLS: ToolDef[] = [
  {
    name: "get_selection",
    description: "Read whatever the user has selected in the document, with its paragraph style.",
    parameters: { type: "object", properties: {} },
  },
  {
    name: "get_document_text",
    description:
      "Read the document as numbered paragraphs, with their styles. Long documents come back in slices, so ask for the next slice when you need it.",
    parameters: {
      type: "object",
      properties: {
        fromParagraph: { type: "number", description: "0 based paragraph to start at." },
        count: { type: "number", description: "How many paragraphs to return. Default 80." },
      },
    },
  },
  {
    name: "get_outline",
    description: "List only the headings, with their paragraph numbers. Cheap way to see the shape of a document.",
    parameters: { type: "object", properties: {} },
  },
  {
    name: "document_stats",
    description: "Paragraph, word and character counts for the whole document.",
    parameters: { type: "object", properties: {} },
  },
  {
    name: "replace_selection",
    description: "Replace the selected text. Use this for rewrites, translations and corrections.",
    parameters: {
      type: "object",
      properties: { text: { type: "string" } },
      required: ["text"],
    },
  },
  {
    name: "insert_text",
    description: "Insert a new paragraph, optionally with a style such as Heading 1 or Quote.",
    parameters: {
      type: "object",
      properties: {
        text: { type: "string" },
        where: {
          type: "string",
          enum: ["afterSelection", "beforeSelection", "endOfDocument", "startOfDocument"],
        },
        style: { type: "string", description: 'e.g. "Heading 1", "Heading 2", "Normal", "Quote"' },
      },
      required: ["text"],
    },
  },
  {
    name: "format_selection",
    description: "Change how the selected text looks, or apply a named Word style to it.",
    parameters: {
      type: "object",
      properties: {
        bold: { type: "boolean" },
        italic: { type: "boolean" },
        underline: { type: "boolean" },
        color: { type: "string", description: "Hex, e.g. #C00000" },
        highlightColor: { type: "string", description: "Hex, e.g. #FFFF00" },
        size: { type: "number", description: "Points" },
        font: { type: "string" },
        alignment: { type: "string", enum: ["Left", "Centered", "Right", "Justified"] },
        style: { type: "string" },
      },
    },
  },
  {
    name: "find_replace",
    description: "Find and replace text through the whole document.",
    parameters: {
      type: "object",
      properties: {
        find: { type: "string" },
        replace: { type: "string" },
        matchCase: { type: "boolean" },
        wholeWord: { type: "boolean" },
      },
      required: ["find", "replace"],
    },
  },
  {
    name: "insert_list",
    description: "Insert a bulleted or numbered list after the selection.",
    parameters: {
      type: "object",
      properties: {
        items: { type: "array", items: { type: "string" } },
        ordered: { type: "boolean", description: "true for 1. 2. 3., false for bullets" },
      },
      required: ["items"],
    },
  },
  {
    name: "insert_table",
    description: "Insert a table after the selection. First row is treated as the header.",
    parameters: {
      type: "object",
      properties: {
        rows: { type: "array", items: { type: "array", items: { type: "string" } } },
        header: { type: "boolean" },
      },
      required: ["rows"],
    },
  },
  {
    name: "insert_page_break",
    description: "Insert a page break after the selection.",
    parameters: { type: "object", properties: {} },
  },
  {
    name: "word_setup",
    description:
      "Word's own machinery in one call: turn track changes on or off, read the comments in the document and resolve them, insert a footnote, set a header or footer, insert a table of contents field, read or set document properties, list the styles in use, read the tables, edit a table cell, add a table row, or replace a numbered paragraph.",
    parameters: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: [
            "track_changes",
            "list_comments",
            "resolve_comment",
            "footnote",
            "header",
            "footer",
            "insert_toc",
            "properties",
            "list_styles",
            "list_tables",
            "set_table_cell",
            "add_table_row",
            "replace_paragraph",
          ],
        },
        mode: { type: "string", enum: ["off", "trackAll", "trackMineOnly"], description: "track_changes." },
        text: { type: "string", description: "Footnote text, header or footer text, or the replacement text." },
        index: { type: "number", description: "Comment number, or paragraph number for replace_paragraph." },
        resolved: { type: "boolean", description: "resolve_comment: false reopens it." },
        title: { type: "string", description: "properties." },
        subject: { type: "string" },
        author: { type: "string" },
        read: { type: "boolean", description: "properties: true reads them instead of setting." },
        table: { type: "number", description: "Which table, 1 based." },
        row: { type: "number", description: "0 based row." },
        column: { type: "number", description: "0 based column." },
        values: { type: "array", items: { type: "string" }, description: "add_table_row: the cells." },
        style: { type: "string", description: "replace_paragraph: style to apply." },
      },
      required: ["action"],
    },
  },
  {
    name: "add_comment",
    description: "Leave a Word comment on the selected text. Good for review notes instead of editing the text.",
    parameters: {
      type: "object",
      properties: { text: { type: "string" } },
      required: ["text"],
    },
  },
];

export async function execWordTool(name: string, args: any): Promise<string | null> {
  switch (name) {
    case "get_selection":
      return W.getSelection();
    case "get_document_text":
      return W.getDocumentText({ fromParagraph: args.fromParagraph, count: args.count });
    case "get_outline":
      return W.getOutline();
    case "document_stats":
      return W.documentStats();
    case "replace_selection":
      return W.replaceSelection({ text: String(args.text ?? "") });
    case "insert_text":
      return W.insertText(args);
    case "format_selection":
      return W.formatSelection(args);
    case "find_replace":
      return W.findReplace(args);
    case "insert_list":
      return W.insertList(args);
    case "insert_table":
      return W.insertTable(args);
    case "insert_page_break":
      return W.insertPageBreak();
    case "add_comment":
      return W.addComment({ text: String(args.text ?? "") });

    case "word_setup":
      switch (args.action) {
        case "track_changes":
          return WP.setTrackChanges({ mode: args.mode ?? "trackAll" });
        case "list_comments":
          return WP.listComments();
        case "resolve_comment":
          return WP.resolveComment({ index: args.index ?? 1, resolved: args.resolved });
        case "footnote":
          return WP.insertFootnote({ text: String(args.text ?? "") });
        case "header":
          return WP.headerFooter({ which: "header", text: String(args.text ?? "") });
        case "footer":
          return WP.headerFooter({ which: "footer", text: String(args.text ?? "") });
        case "insert_toc":
          return WP.insertTableOfContents();
        case "properties":
          return WP.documentProperties(args);
        case "list_styles":
          return WP.listStyles();
        case "list_tables":
          return WP.listTables();
        case "set_table_cell":
          return WP.setTableCell({
            table: args.table ?? 1,
            row: args.row ?? 0,
            column: args.column ?? 0,
            text: String(args.text ?? ""),
          });
        case "add_table_row":
          return WP.addTableRow({ table: args.table ?? 1, values: args.values ?? [] });
        case "replace_paragraph":
          return WP.replaceParagraph({
            index: args.index ?? 0,
            text: String(args.text ?? ""),
            style: args.style,
          });
        default:
          return "Unknown Word action: " + args.action;
      }
    default:
      return null;
  }
}

/** Word tools that change the document, so the pane can ask first. */
export const WORD_WRITING_TOOLS = new Set([
  "word_setup",
  "replace_selection",
  "insert_text",
  "find_replace",
  "insert_list",
  "insert_table",
  "insert_page_break",
]);
