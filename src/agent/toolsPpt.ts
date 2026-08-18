import type { ToolDef } from "./tools";
import * as P from "../office/ppt";
import * as PP from "../office/pptPro";

export const PPT_TOOLS: ToolDef[] = [
  {
    name: "list_slides",
    description: "List every slide with a preview of the text on it. Start here to find your way around a deck.",
    parameters: { type: "object", properties: {} },
  },
  {
    name: "get_slide",
    description: "List the shapes on one slide with their ids, positions and text.",
    parameters: {
      type: "object",
      properties: { slide: { type: "number", description: "1 based slide number" } },
      required: ["slide"],
    },
  },
  {
    name: "list_layouts",
    description: "List the slide layouts available in this deck's master.",
    parameters: { type: "object", properties: {} },
  },
  {
    name: "add_bullet_slide",
    description:
      "Add a slide with a title and bullet points. This is the quickest way to build out a deck.",
    parameters: {
      type: "object",
      properties: {
        title: { type: "string" },
        bullets: { type: "array", items: { type: "string" } },
      },
      required: ["title", "bullets"],
    },
  },
  {
    name: "add_slide",
    description: "Add an empty slide, optionally using a named layout from list_layouts.",
    parameters: {
      type: "object",
      properties: { layoutName: { type: "string" } },
    },
  },
  {
    name: "delete_slide",
    description: "Delete a slide by its 1 based number.",
    parameters: {
      type: "object",
      properties: { slide: { type: "number" } },
      required: ["slide"],
    },
  },
  {
    name: "add_text_box",
    description: "Put a text box on a slide. Positions are in points, a slide is about 720 by 540.",
    parameters: {
      type: "object",
      properties: {
        slide: { type: "number" },
        text: { type: "string" },
        left: { type: "number" },
        top: { type: "number" },
        width: { type: "number" },
        height: { type: "number" },
        fontSize: { type: "number" },
        bold: { type: "boolean" },
        color: { type: "string", description: "Hex, e.g. #1F3864" },
      },
      required: ["slide", "text"],
    },
  },
  {
    name: "set_shape_text",
    description: "Replace the text inside a shape you found with get_slide.",
    parameters: {
      type: "object",
      properties: {
        slide: { type: "number" },
        shapeId: { type: "string" },
        shapeName: { type: "string" },
        text: { type: "string" },
      },
      required: ["slide", "text"],
    },
  },
  {
    name: "deck_setup",
    description:
      "Shapes and looks: add a geometric shape (Rectangle, Ellipse, RoundRectangle, Chevron, Arrow…), restyle or move an existing shape, place an image from base64 data, or set the speaker notes on a slide. Positions are in points on a 720 by 540 slide.",
    parameters: {
      type: "object",
      properties: {
        action: { type: "string", enum: ["add_shape", "format_shape", "add_image", "notes"] },
        slide: { type: "number", description: "1 based slide number." },
        shape: { type: "string", description: "add_shape: the geometry, default Rectangle." },
        shapeId: { type: "string", description: "format_shape: id from get_slide." },
        shapeName: { type: "string" },
        text: { type: "string", description: "Shape text, or the speaker notes." },
        base64: { type: "string", description: "add_image: raw base64 image data." },
        left: { type: "number" },
        top: { type: "number" },
        width: { type: "number" },
        height: { type: "number" },
        fillColor: { type: "string", description: "Hex, e.g. #14724A" },
        lineColor: { type: "string" },
        fontColor: { type: "string" },
        fontSize: { type: "number" },
        bold: { type: "boolean" },
      },
      required: ["action", "slide"],
    },
  },
  {
    name: "delete_shape",
    description: "Delete a shape from a slide by its id.",
    parameters: {
      type: "object",
      properties: { slide: { type: "number" }, shapeId: { type: "string" } },
      required: ["slide", "shapeId"],
    },
  },
];

export async function execPptTool(name: string, args: any): Promise<string | null> {
  switch (name) {
    case "list_slides":
      return P.listSlides();
    case "get_slide":
      return P.getSlide({ slide: Number(args.slide) });
    case "list_layouts":
      return P.listLayouts();
    case "add_bullet_slide":
      return P.addBulletSlide({ title: String(args.title ?? ""), bullets: args.bullets ?? [] });
    case "add_slide":
      return P.addSlide({ layoutName: args.layoutName });
    case "delete_slide":
      return P.deleteSlide({ slide: Number(args.slide) });
    case "add_text_box":
      return P.addTextBox(args);
    case "set_shape_text":
      return P.setShapeText(args);
    case "delete_shape":
      return P.deleteShape({ slide: Number(args.slide), shapeId: String(args.shapeId) });

    case "deck_setup":
      switch (args.action) {
        case "add_shape":
          return PP.addShape(args);
        case "format_shape":
          return PP.formatShape(args);
        case "add_image":
          return PP.addImage(args);
        case "notes":
          return PP.setSlideNotes({ slide: Number(args.slide), text: String(args.text ?? "") });
        default:
          return "Unknown deck action: " + args.action;
      }
    default:
      return null;
  }
}

/** Deck changing tools, so the pane can ask before they run. */
export const PPT_WRITING_TOOLS = new Set([
  "deck_setup",
  "add_bullet_slide",
  "add_slide",
  "delete_slide",
  "add_text_box",
  "set_shape_text",
  "delete_shape",
]);
