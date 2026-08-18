/** PowerPoint side of tANk. The API is thinner than Excel's: slides, shapes and text frames. */

function textOfShape(shape: PowerPoint.Shape): string {
  try {
    return shape.textFrame?.textRange?.text ?? "";
  } catch {
    return "";
  }
}

export async function listSlides(): Promise<string> {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();

    const shapeSets = slides.items.map((slide) => {
      const shapes = slide.shapes;
      shapes.load("items/id,items/name,items/textFrame/textRange/text");
      return shapes;
    });
    await ctx.sync();

    return slides.items
      .map((_slide, i) => {
        const texts = shapeSets[i].items
          .map((sh) => textOfShape(sh).trim())
          .filter(Boolean)
          .join(" | ");
        return (
          "Slide " +
          (i + 1) +
          " (" +
          shapeSets[i].items.length +
          " shapes): " +
          (texts.length > 160 ? texts.slice(0, 160) + "…" : texts || "empty")
        );
      })
      .join("\n");
  });
}

export async function getSlide(args: { slide: number }): Promise<string> {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();

    const idx = (args.slide ?? 1) - 1;
    if (idx < 0 || idx >= slides.items.length) return "There is no slide " + args.slide + ".";

    const shapes = slides.items[idx].shapes;
    shapes.load("items/id,items/name,items/left,items/top,items/width,items/height,items/textFrame/textRange/text");
    await ctx.sync();

    if (!shapes.items.length) return "Slide " + args.slide + " is empty.";
    return shapes.items
      .map(
        (sh) =>
          sh.name +
          " (id " +
          sh.id +
          ", " +
          Math.round(sh.left) +
          "," +
          Math.round(sh.top) +
          " " +
          Math.round(sh.width) +
          "x" +
          Math.round(sh.height) +
          "): " +
          textOfShape(sh)
      )
      .join("\n");
  });
}

export async function addSlide(args: { layoutName?: string }): Promise<string> {
  return PowerPoint.run(async (ctx) => {
    if (args.layoutName) {
      const layouts = ctx.presentation.slideMasters.getItemAt(0).layouts;
      layouts.load("items/id,items/name");
      await ctx.sync();
      const match = layouts.items.find(
        (l) => l.name.toLowerCase() === args.layoutName!.toLowerCase()
      );
      if (match) {
        ctx.presentation.slides.add({ layoutId: match.id });
        await ctx.sync();
        return 'added a slide using the "' + match.name + '" layout';
      }
    }
    ctx.presentation.slides.add();
    await ctx.sync();
    return "added a slide";
  });
}

export async function listLayouts(): Promise<string> {
  return PowerPoint.run(async (ctx) => {
    const layouts = ctx.presentation.slideMasters.getItemAt(0).layouts;
    layouts.load("items/name");
    await ctx.sync();
    return layouts.items.map((l) => l.name).join("\n");
  });
}

export async function deleteSlide(args: { slide: number }): Promise<string> {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();
    const idx = (args.slide ?? 0) - 1;
    if (idx < 0 || idx >= slides.items.length) return "There is no slide " + args.slide + ".";
    slides.items[idx].delete();
    await ctx.sync();
    return "deleted slide " + args.slide;
  });
}

export async function addTextBox(args: {
  slide: number;
  text: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  fontSize?: number;
  bold?: boolean;
  color?: string;
}): Promise<string> {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();

    const idx = (args.slide ?? 1) - 1;
    if (idx < 0 || idx >= slides.items.length) return "There is no slide " + args.slide + ".";

    const box = slides.items[idx].shapes.addTextBox(args.text, {
      left: args.left ?? 60,
      top: args.top ?? 60,
      width: args.width ?? 560,
      height: args.height ?? 120,
    });
    const range = box.textFrame.textRange;
    if (args.fontSize) range.font.size = args.fontSize;
    if (args.bold !== undefined) range.font.bold = args.bold;
    if (args.color) range.font.color = args.color;
    await ctx.sync();
    return "added a text box to slide " + args.slide;
  });
}

export async function setShapeText(args: {
  slide: number;
  shapeId?: string;
  shapeName?: string;
  text: string;
}): Promise<string> {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();

    const idx = (args.slide ?? 1) - 1;
    if (idx < 0 || idx >= slides.items.length) return "There is no slide " + args.slide + ".";

    const shapes = slides.items[idx].shapes;
    shapes.load("items/id,items/name");
    await ctx.sync();

    const shape = args.shapeId
      ? shapes.items.find((s) => s.id === args.shapeId)
      : shapes.items.find((s) => s.name === args.shapeName);
    if (!shape) return "No shape matched on slide " + args.slide + ".";

    shape.textFrame.textRange.text = args.text;
    await ctx.sync();
    return "set the text of " + shape.name + " on slide " + args.slide;
  });
}

export async function deleteShape(args: { slide: number; shapeId: string }): Promise<string> {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();
    const idx = (args.slide ?? 1) - 1;
    if (idx < 0 || idx >= slides.items.length) return "There is no slide " + args.slide + ".";

    const shapes = slides.items[idx].shapes;
    shapes.load("items/id");
    await ctx.sync();
    const shape = shapes.items.find((s) => s.id === args.shapeId);
    if (!shape) return "No shape with that id on slide " + args.slide + ".";
    shape.delete();
    await ctx.sync();
    return "deleted the shape";
  });
}

/** Title plus bullet lines, the shape of slide most decks are made of. */
export async function addBulletSlide(args: {
  title: string;
  bullets: string[];
  position?: number;
}): Promise<string> {
  return PowerPoint.run(async (ctx) => {
    ctx.presentation.slides.add();
    await ctx.sync();

    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();

    const slide = slides.items[slides.items.length - 1];
    const title = slide.shapes.addTextBox(args.title, {
      left: 60,
      top: 40,
      width: 600,
      height: 60,
    });
    title.textFrame.textRange.font.size = 30;
    title.textFrame.textRange.font.bold = true;

    const body = slide.shapes.addTextBox((args.bullets ?? []).map((b) => "• " + b).join("\n"), {
      left: 60,
      top: 120,
      width: 600,
      height: 320,
    });
    body.textFrame.textRange.font.size = 18;
    await ctx.sync();
    return 'added a slide titled "' + args.title + '" with ' + (args.bullets?.length ?? 0) + " bullets";
  });
}
