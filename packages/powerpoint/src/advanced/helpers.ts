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

    async readShape(shape: PowerPoint.Shape): Promise<ShapeSummary> {
      shape.load("id,name,type,left,top,width,height");
      await context.sync();
      const summary: ShapeSummary = {
        id: shape.id,
        name: shape.name,
        type: shape.type,
        left: round(shape.left),
        top: round(shape.top),
        width: round(shape.width),
        height: round(shape.height),
      };
      const text = await readText(context, shape);
      return text === undefined ? summary : { ...summary, text };
    },

    /** JSON-safe copy (Office objects serialize their loaded properties). */
    json(value: unknown): unknown {
      return value === undefined ? null : JSON.parse(JSON.stringify(value));
    },
  };
}

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
  // Before 1.10, loading text on a shape without a text frame (pictures, charts) fails the sync.
  try {
    shape.textFrame.textRange.load("text");
    await context.sync();
    return shape.textFrame.textRange.text;
  } catch {
    return undefined;
  }
}

/** Markdown reference for the execute_office_js tool description. */
export const codeHelpersReference = `\`footnote\` helpers (units are points; slide numbers are 1-based like the deck outline):
- \`footnote.slide(idOrNumber)\` → PowerPoint.Slide by ID ("267#…") or slide number.
- \`footnote.shape(slideIdOrNumber, shapeId)\` → PowerPoint.Shape.
- \`await footnote.allSlides()\` → loaded Slide[] (ids).
- \`footnote.pt(n)\`, \`footnote.inches(n)\`, \`footnote.cm(n)\` → points.
- \`footnote.setText(shape, text, { font, size, color: "#RRGGBB", bold, italic, align: "Left"|"Center"|"Right"|"Justify" }?)\` queues text + formatting; call \`await context.sync()\` after.
- \`footnote.fit(shape)\` queues word wrap + shrink-text-on-overflow.
- \`await footnote.readShape(shape)\` → { id, name, type, left, top, width, height, text? } (syncs).
- \`footnote.json(value)\` → JSON-safe copy of loaded Office objects for the return value.`;
