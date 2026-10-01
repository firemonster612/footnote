import { fingerprintShapes } from "./fingerprint.ts";
import { findShape, findSlide, requireApi } from "./presentation.ts";
import { readShapeLists } from "./shapeReader.ts";
import type {
  AddShapeArgs,
  FontStyle,
  ShapeInfo,
  ShapeStyle,
  ShapeUpdate,
  WriteReceipt,
} from "./types.ts";

const MAX_VERIFIED_TEXT = 300;

export function applyFont(font: PowerPoint.ShapeFont, style: FontStyle): void {
  if (style.name !== undefined) font.name = style.name;
  if (style.size !== undefined) font.size = style.size;
  if (style.color !== undefined) font.color = style.color;
  if (style.bold !== undefined) font.bold = style.bold;
  if (style.italic !== undefined) font.italic = style.italic;
  if (style.underline !== undefined) font.underline = style.underline ? "Single" : "None";
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/** Queues every style change on `shape` and syncs. */
export async function applyStyle(
  context: PowerPoint.RequestContext,
  shape: PowerPoint.Shape,
  style: ShapeStyle,
): Promise<void> {
  if (style.text !== undefined && style.runs)
    throw new Error("Pass either text or runs, not both.");
  if (style.name !== undefined) shape.name = style.name;
  if (style.left !== undefined) shape.left = style.left;
  if (style.top !== undefined) shape.top = style.top;
  if (style.width !== undefined) shape.width = style.width;
  if (style.height !== undefined) shape.height = style.height;
  if (style.rotation !== undefined) {
    requireApi("1.10", "Rotation");
    shape.rotation = style.rotation;
  }

  const range = shape.textFrame.textRange;
  if (style.text !== undefined) range.text = style.text;
  if (style.runs) {
    range.text = style.runs.map((run) => run.text).join("");
    let start = 0;
    for (const { text, ...font } of style.runs) {
      if (Object.keys(font).length > 0 && text.length > 0)
        applyFont(range.getSubstring(start, text.length).font, font);
      start += text.length;
    }
  }
  if (style.font) applyFont(range.font, style.font);
  if (style.paragraph) {
    const { align, bullets, indentLevel } = style.paragraph;
    if (align) range.paragraphFormat.horizontalAlignment = align;
    if (bullets !== undefined) range.paragraphFormat.bulletFormat.visible = bullets;
    if (indentLevel !== undefined) {
      requireApi("1.10", "Indent level");
      range.paragraphFormat.indentLevel = indentLevel;
    }
  }
  if (style.textFrame) {
    const { autoSize, verticalAlign, wordWrap } = style.textFrame;
    if (autoSize) shape.textFrame.autoSizeSetting = autoSize;
    if (verticalAlign) shape.textFrame.verticalAlignment = verticalAlign;
    if (wordWrap !== undefined) shape.textFrame.wordWrap = wordWrap;
  }
  if (style.fill) {
    if (style.fill.color === "none") shape.fill.clear();
    else shape.fill.setSolidColor(style.fill.color);
    if (style.fill.transparency !== undefined) shape.fill.transparency = style.fill.transparency;
  }
  if (style.line) {
    const { color, weight, dash, visible } = style.line;
    if (visible !== undefined) shape.lineFormat.visible = visible;
    if (color) shape.lineFormat.color = color;
    if (weight !== undefined) shape.lineFormat.weight = weight;
    if (dash) shape.lineFormat.dashStyle = dash;
  }
  await context.sync();
  if (style.hyperlink) await setHyperlink(context, shape, style.hyperlink);
}

async function setHyperlink(
  context: PowerPoint.RequestContext,
  shape: PowerPoint.Shape,
  { text, ...link }: NonNullable<ShapeStyle["hyperlink"]>,
): Promise<void> {
  requireApi("1.10", "Hyperlinks");
  if (text === undefined) {
    shape.setHyperlink(link);
    await context.sync();
    return;
  }
  const range = shape.textFrame.textRange.load("text");
  await context.sync();
  const start = range.text.indexOf(text);
  if (start < 0) throw new Error(`Hyperlink text "${text}" not found in shape ${shape.id}.`);
  range.getSubstring(start, text.length).setHyperlink(link);
  await context.sync();
}

function compactShape(shape: ShapeInfo): ShapeInfo {
  const { paragraphs: _, children, text, ...rest } = shape;
  return {
    ...rest,
    ...(text !== undefined && {
      text: text.length > MAX_VERIFIED_TEXT ? `${text.slice(0, MAX_VERIFIED_TEXT)}…` : text,
    }),
    ...(children && { children: children.map(compactShape) }),
  };
}

function flattenShapes(shapes: ShapeInfo[]): ShapeInfo[] {
  return shapes.flatMap((shape) => [shape, ...flattenShapes(shape.children ?? [])]);
}

/** Receipt for shape-level writes: reads the slide back and reports the changed shapes as they are now. */
export async function shapeReceipt(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  changed: string[],
  warnings: string[],
  deleted: string[] = [],
): Promise<WriteReceipt> {
  const [shapes = []] = await readShapeLists(context, [slide.shapes], true);
  const byId = new Map(flattenShapes(shapes).map((shape) => [shape.id, shape]));
  const verified: Record<string, unknown> = {};
  for (const id of changed) {
    const shape = byId.get(id);
    if (deleted.includes(id)) {
      if (shape) warnings.push(`Shape ${id} still exists after delete.`);
      else verified[id] = "deleted";
    } else if (shape) verified[id] = compactShape(shape);
    else warnings.push(`Shape ${id} not found after the write.`);
  }
  return { changed, verified, warnings, fingerprints: { [slide.id]: fingerprintShapes(shapes) } };
}

function createShape(shapes: PowerPoint.ShapeCollection, args: AddShapeArgs): PowerPoint.Shape {
  const bounds = { left: args.left, top: args.top, width: args.width, height: args.height };
  switch (args.kind) {
    case "textbox":
      return shapes.addTextBox(args.text ?? "", bounds);
    case "line":
      return shapes.addLine(args.connector ?? "Straight", bounds);
    case "image":
      return shapes.addGeometricShape("Rectangle", bounds);
    case "geometric":
      // Office validates the name and fails the sync with InvalidArgument for unknown geometry.
      return shapes.addGeometricShape(
        (args.geometry ?? "Rectangle") as PowerPoint.GeometricShapeType,
        bounds,
      );
  }
}

export const shapeOps = {
  add_shape: (args: AddShapeArgs) =>
    PowerPoint.run(async (context) => {
      const { slideId, kind, geometry: _, connector: __, imageBase64, ...style } = args;
      const { slide } = await findSlide(context, slideId);
      if (kind === "image" && !imageBase64) throw new Error("kind image needs an image.");
      const shape = createShape(slide.shapes, args).load("id");
      await context.sync();

      const warnings: string[] = [];
      try {
        if (imageBase64) {
          requireApi("1.8", "Images");
          shape.fill.setImage(imageBase64);
          shape.lineFormat.visible = false;
        }
        await applyStyle(context, shape, style);
      } catch (error) {
        warnings.push(`Shape ${shape.id} was created but styling failed: ${errorMessage(error)}`);
      }
      return shapeReceipt(context, slide, [shape.id], warnings);
    }),

  update_shapes: ({ slideId, updates }: { slideId: string; updates: ShapeUpdate[] }) =>
    PowerPoint.run(async (context) => {
      const { slide } = await findSlide(context, slideId);
      const warnings: string[] = [];
      const changed: string[] = [];
      const deleted: string[] = [];
      // One sync per update so a failure names its shape and the rest still apply.
      for (const { shapeId, delete: remove, zOrder, ...style } of updates) {
        try {
          const shape = await findShape(slide, shapeId);
          if (remove) {
            shape.delete();
            await context.sync();
            deleted.push(shapeId);
          } else {
            if (zOrder) {
              requireApi("1.8", "Z-order");
              shape.setZOrder(zOrder);
            }
            await applyStyle(context, shape, style);
          }
          changed.push(shapeId);
        } catch (error) {
          warnings.push(`Shape ${shapeId}: ${errorMessage(error)}`);
        }
      }
      return shapeReceipt(context, slide, changed, warnings, deleted);
    }),

  group_shapes: ({
    slideId,
    shapeIds,
    name,
  }: {
    slideId: string;
    shapeIds: string[];
    name?: string;
  }) =>
    PowerPoint.run(async (context) => {
      requireApi("1.8", "Grouping");
      const { slide } = await findSlide(context, slideId);
      const group = slide.shapes.addGroup(shapeIds).load("id");
      if (name) group.name = name;
      await context.sync();
      return shapeReceipt(context, slide, [group.id], []);
    }),

  ungroup_shape: ({ slideId, shapeId }: { slideId: string; shapeId: string }) =>
    PowerPoint.run(async (context) => {
      requireApi("1.8", "Ungrouping");
      const { slide } = await findSlide(context, slideId);
      const shape = await findShape(slide, shapeId);
      const children = shape.group.shapes.load("items/id");
      await context.sync();
      const childIds = children.items.map((child) => child.id);
      shape.group.ungroup();
      await context.sync();
      return shapeReceipt(context, slide, childIds, []);
    }),
};
