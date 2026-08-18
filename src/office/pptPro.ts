/** PowerPoint beyond text: shapes, images, colours, sizing, and slide notes where the API allows. */

async function slideAt(ctx: PowerPoint.RequestContext, index: number): Promise<PowerPoint.Slide | null> {
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  const idx = (index ?? 1) - 1;
  return idx >= 0 && idx < slides.items.length ? slides.items[idx] : null;
}

export async function addShape(args: {
  slide: number;
  shape?: string;
  text?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  fillColor?: string;
  lineColor?: string;
  fontColor?: string;
  fontSize?: number;
}): Promise<string> {
  return PowerPoint.run(async (ctx) => {
    const slide = await slideAt(ctx, args.slide);
    if (!slide) return "there is no slide " + args.slide;

    const shape = slide.shapes.addGeometricShape(
      (args.shape ?? "Rectangle") as unknown as PowerPoint.GeometricShapeType,
      {
        left: args.left ?? 60,
        top: args.top ?? 60,
        width: args.width ?? 200,
        height: args.height ?? 100,
      }
    );
    if (args.fillColor) shape.fill.setSolidColor(args.fillColor);
    if (args.lineColor) shape.lineFormat.color = args.lineColor;
    if (args.text) {
      shape.textFrame.textRange.text = args.text;
      if (args.fontColor) shape.textFrame.textRange.font.color = args.fontColor;
      if (args.fontSize) shape.textFrame.textRange.font.size = args.fontSize;
    }
    await ctx.sync();
    return "added a " + (args.shape ?? "Rectangle") + " to slide " + args.slide;
  });
}

export async function formatShape(args: {
  slide: number;
  shapeId?: string;
  shapeName?: string;
  fillColor?: string;
  lineColor?: string;
  fontColor?: string;
  fontSize?: number;
  bold?: boolean;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}): Promise<string> {
  return PowerPoint.run(async (ctx) => {
    const slide = await slideAt(ctx, args.slide);
    if (!slide) return "there is no slide " + args.slide;

    const shapes = slide.shapes;
    shapes.load("items/id,items/name");
    await ctx.sync();

    const shape = args.shapeId
      ? shapes.items.find((s) => s.id === args.shapeId)
      : args.shapeName
        ? shapes.items.find((s) => s.name === args.shapeName)
        : shapes.items[0];
    if (!shape) return "no shape matched on slide " + args.slide;

    if (args.fillColor) shape.fill.setSolidColor(args.fillColor);
    if (args.lineColor) shape.lineFormat.color = args.lineColor;
    if (args.left !== undefined) shape.left = args.left;
    if (args.top !== undefined) shape.top = args.top;
    if (args.width !== undefined) shape.width = args.width;
    if (args.height !== undefined) shape.height = args.height;

    const font = shape.textFrame.textRange.font;
    if (args.fontColor) font.color = args.fontColor;
    if (args.fontSize) font.size = args.fontSize;
    if (args.bold !== undefined) font.bold = args.bold;

    await ctx.sync();
    return "formatted " + shape.name + " on slide " + args.slide;
  });
}

export async function addImage(args: {
  slide: number;
  base64: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}): Promise<string> {
  return PowerPoint.run(async (ctx) => {
    const slide = await slideAt(ctx, args.slide);
    if (!slide) return "there is no slide " + args.slide;
    const clean = (args.base64 ?? "").replace(/^data:image\/\w+;base64,/, "");
    if (!clean) return "no image data given";

    // addImage is newer than the typings bundled here; it throws on builds that lack it.
    const shapes = slide.shapes as unknown as {
      addImage?: (data: string, opts: Record<string, unknown>) => unknown;
    };
    if (!shapes.addImage) return "this build of PowerPoint cannot place images from the add-in";
    shapes.addImage(clean, {
      left: args.left ?? 60,
      top: args.top ?? 60,
      width: args.width,
      height: args.height,
    });
    await ctx.sync();
    return "image placed on slide " + args.slide;
  });
}

export async function setSlideNotes(args: { slide: number; text: string }): Promise<string> {
  return PowerPoint.run(async (ctx) => {
    const slide = await slideAt(ctx, args.slide);
    if (!slide) return "there is no slide " + args.slide;
    // Speaker notes arrived in PowerPointApi 1.8; older builds throw and the model is told.
    const notes = (slide as unknown as { notes?: { body?: { text: string } } }).notes;
    if (!notes?.body) return "this build of PowerPoint does not expose speaker notes";
    notes.body.text = args.text;
    await ctx.sync();
    return "speaker notes set on slide " + args.slide;
  });
}
