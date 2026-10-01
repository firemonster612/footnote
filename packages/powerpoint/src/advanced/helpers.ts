/// <reference types="office-js" />
// `footnote` helper library for code mode (execute_office_js). Runs in the Office realm next to the model's code.

const POINTS_PER_INCH = 72;
const CM_PER_INCH = 2.54;

export interface TextStyle {
  font?: string;
  size?: number;
  /** "#RRGGBB". */
  color?: string;
  bold?: boolean;
  italic?: boolean;
  align?: "Left" | "Center" | "Right" | "Justify";
}

export interface ShapeSummary {
  id: string;
  name: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  text?: string;
}

const round = (n: number) => Math.round(n * 10) / 10;

/** Shape types that have a text frame before PowerPointApi 1.10 added getTextFrameOrNullObject. */
const TEXT_SHAPE_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder"]);

export interface TextShape {
  shape: PowerPoint.Shape;
  id: string;
  name: string;
  type: string;
  text: string;
}

export function createFootnoteHelpers(context: PowerPoint.RequestContext) {
  const slides = context.presentation.slides;

  function slide(idOrNumber: string | number): PowerPoint.Slide {
    return typeof idOrNumber === "string"
      ? slides.getItem(idOrNumber)
      : slides.getItemAt(idOrNumber - 1);
  }

  return {
    slide,

    shape(slideIdOrNumber: string | number, shapeId: string): PowerPoint.Shape {
      return slide(slideIdOrNumber).shapes.getItem(shapeId);
    },

    async allSlides(): Promise<PowerPoint.Slide[]> {
      slides.load("items/id");
      await context.sync();
      return slides.items;
    },

    pt: (points: number) => points,
    inches: (inches: number) => inches * POINTS_PER_INCH,
    cm: (cm: number) => (cm / CM_PER_INCH) * POINTS_PER_INCH,

    setText(shape: PowerPoint.Shape, text: string, style: TextStyle = {}): void {
      const range = shape.textFrame.textRange;
      range.text = text;
      const { align, font, size, color, bold, italic } = style;
      if (font !== undefined) range.font.name = font;
      if (size !== undefined) range.font.size = size;
      if (color !== undefined) range.font.color = color;
      if (bold !== undefined) range.font.bold = bold;
      if (italic !== undefined) range.font.italic = italic;
      if (align) range.paragraphFormat.horizontalAlignment = align;
    },

    fit(shape: PowerPoint.Shape): void {
      shape.textFrame.wordWrap = true;
      shape.textFrame.autoSizeSetting = "AutoSizeTextToFitShape";
    },

    /** Loads `props` on every object in one sync. */
    async load<T extends { load(props: string): unknown }>(
      objects: T[],
      props: string,
    ): Promise<T[]> {
      for (const object of objects) object.load(props);
      await context.sync();
      return objects;
    },

    async readShape(shape: PowerPoint.Shape): Promise<ShapeSummary> {
      shape.load("id,name,type,left,top,width,height");
      const text = await readText(context, shape);
      const summary: ShapeSummary = {
        id: shape.id,
        name: shape.name,
        type: shape.type,
        left: round(shape.left),
        top: round(shape.top),
        width: round(shape.width),
        height: round(shape.height),
      };
      return text === undefined ? summary : { ...summary, text };
    },

    /**
     * Shapes that have a text frame, with their text loaded, on one slide or several (two syncs in total however many
     * slides). Loading text on lines, pictures or groups fails the whole sync.
     */
    async textShapes(targets: PowerPoint.Slide | PowerPoint.Slide[]): Promise<TextShape[]> {
      const lists = (Array.isArray(targets) ? targets : [targets]).map((target) =>
        target.shapes.load("items/id,items/name,items/type"),
      );
      await context.sync();
      const candidates = lists.flatMap((list) =>
        list.items.filter((shape) => TEXT_SHAPE_TYPES.has(shape.type)),
      );
      const ranges = candidates.map((shape) => shape.textFrame.textRange.load("text"));
      await context.sync();
      return candidates.map((shape, i) => ({
        shape,
        id: shape.id,
        name: shape.name,
        type: shape.type,
        text: ranges[i]!.text,
      }));
    },

    /**
     * Queues deletion of the shapes whose name starts with `namePrefix`, on one slide or several, and returns how many.
     * Syncs once to find them; the deletes go out with your next sync, ahead of anything you queue after this call.
     */
    async removeShapes(
      targets: PowerPoint.Slide | PowerPoint.Slide[],
      namePrefix: string,
    ): Promise<number> {
      const lists = (Array.isArray(targets) ? targets : [targets]).map((target) =>
        target.shapes.load("items/name"),
      );
      await context.sync();
      const matches = lists.flatMap((list) =>
        list.items.filter((shape) => shape.name.startsWith(namePrefix)),
      );
      for (const shape of matches) shape.delete();
      return matches.length;
    },

    /** Draws on a canvas and returns base64 PNG for \`shape.fill.setImage(base64)\`: the way to get blur, gradients, glows and shadows, which Office.js can't set. */
    async canvasImage(
      width: number,
      height: number,
      draw: (ctx: OffscreenCanvasRenderingContext2D) => void,
    ): Promise<string> {
      const canvas = new OffscreenCanvas(width, height);
      const ctx = canvas.getContext("2d");
      if (!ctx) throw new Error("Canvas 2D isn't available here.");
      draw(ctx);
      const blob = await canvas.convertToBlob({ type: "image/png" });
      const bytes = new Uint8Array(await blob.arrayBuffer());
      let binary = "";
      for (let i = 0; i < bytes.length; i += 0x8000)
        binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
      return btoa(binary);
    },

    /** JSON-safe copy (Office objects serialize their loaded properties). */
    json(value: unknown): unknown {
      return value === undefined ? null : JSON.parse(JSON.stringify(value));
    },
  };
}

/**
 * Syncs whatever the caller queued on `shape` (which must include `type`) together with its text; undefined when it
 * has no text frame. One sync on PowerPointApi 1.10; before it, a second for text, since loading text on a shape
 * without a frame fails the sync.
 */
async function readText(
  context: PowerPoint.RequestContext,
  shape: PowerPoint.Shape,
): Promise<string | undefined> {
  if (Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
    const frame = shape.getTextFrameOrNullObject();
    frame.load("hasText");
    frame.textRange.load("text");
    await context.sync();
    return frame.isNullObject ? undefined : frame.textRange.text;
  }
  await context.sync();
  if (!TEXT_SHAPE_TYPES.has(shape.type)) return undefined;
  const range = shape.textFrame.textRange.load("text");
  await context.sync();
  return range.text;
}

/** Markdown reference for the execute_office_js tool description. */
export const codeHelpersReference = `\`footnote\` helpers (units are points; slide numbers are 1-based like the deck outline):
- \`footnote.slide(idOrNumber)\` → PowerPoint.Slide by ID ("267#…") or slide number.
- \`footnote.shape(slideIdOrNumber, shapeId)\` → PowerPoint.Shape.
- \`await footnote.allSlides()\` → loaded Slide[] (ids).
- \`await footnote.load(objects, "prop,prop")\` → loads the properties on every object in one sync.
- \`footnote.pt(n)\`, \`footnote.inches(n)\`, \`footnote.cm(n)\` → points.
- \`footnote.setText(shape, text, { font, size, color: "#RRGGBB", bold, italic, align: "Left"|"Center"|"Right"|"Justify" }?)\` queues text + formatting.
- \`footnote.fit(shape)\` queues word wrap + shrink-text-on-overflow.
- \`await footnote.readShape(shape)\` → { id, name, type, left, top, width, height, text? } (syncs).
- \`await footnote.textShapes(slideOrSlides)\` → [{ shape, id, name, type, text }] for shapes that have text, on one slide or an array of slides, in two syncs total. Use it instead of loading \`textFrame\` on every shape.
- \`await footnote.removeShapes(slideOrSlides, "fn-glow")\` → count; queues the deletes (they go out with your next sync, before shapes you add after it). Name generated shapes with a prefix and remove them first so reruns don't duplicate.
- \`await footnote.canvasImage(w, h, ctx => { ctx.filter = "blur(40px)"; … })\` → base64 PNG; apply with \`shape.fill.setImage(png)\` on a rectangle. Use it for blur, gradients, glows and shadows.
- \`footnote.json(value)\` → JSON-safe copy of loaded Office objects for the return value.`;
