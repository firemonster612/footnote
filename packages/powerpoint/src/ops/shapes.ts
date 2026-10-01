import { fingerprintShapes } from "./fingerprint.ts";
import { type CheckedWrite, guardWrite, textShapeIds } from "./guard.ts";
import {
  batchFailureWarning,
  describeError,
  readBackFailureWarning,
  requireApi,
  shapeNotFound,
  supportsApi,
} from "./presentation.ts";
import {
  type PendingShapeLists,
  queueShapeLists,
  readShapeLists,
  TEXT_TYPES,
} from "./shapeReader.ts";
import type {
  AddShapeArgs,
  FontStyle,
  ShapeInfo,
  ShapeStyle,
  ShapeUpdate,
  SlideShapeUpdates,
  WriteGuardArgs,
  WriteReceipt,
} from "./types.ts";

const MAX_VERIFIED_TEXT = 300;

type ShapeList = PowerPoint.ShapeCollection;

export function applyFont(font: PowerPoint.ShapeFont, style: FontStyle): void {
  if (style.name !== undefined) font.name = style.name;
  if (style.size !== undefined) font.size = style.size;
  if (style.color !== undefined) font.color = style.color;
  if (style.bold !== undefined) font.bold = style.bold;
  if (style.italic !== undefined) font.italic = style.italic;
  if (style.underline !== undefined) font.underline = style.underline ? "Single" : "None";
}

const touchesText = (style: ShapeStyle): boolean =>
  style.text !== undefined ||
  style.runs !== undefined ||
  style.font !== undefined ||
  style.paragraph !== undefined ||
  style.textFrame !== undefined ||
  style.hyperlink !== undefined;

const replacedText = (style: ShapeStyle): string | undefined =>
  style.runs ? style.runs.map((run) => run.text).join("") : style.text;

/**
 * Checks that `style` can apply, before anything is queued, so a bad field fails alone instead of failing the batch.
 * `currentText` is the shape's text, needed only to link part of text the style doesn't replace.
 * Returns where the linked text starts, when the style links part of the text.
 */
function checkStyle(style: ShapeStyle, currentText?: string): number | undefined {
  if (style.text !== undefined && style.runs)
    throw new Error("Pass either text or runs, not both.");
  if (style.rotation !== undefined) requireApi("1.10", "Rotation");
  if (style.paragraph?.indentLevel !== undefined) requireApi("1.10", "Indent level");
  if (!style.hyperlink) return undefined;
  requireApi("1.10", "Hyperlinks");
  const linkText = style.hyperlink.text;
  if (linkText === undefined) return undefined;
  const start = (replacedText(style) ?? currentText ?? "").indexOf(linkText);
  if (start < 0) throw new Error(`Hyperlink text "${linkText}" not found in the shape's text.`);
  return start;
}

/** Queues every style change on `shape`; the caller syncs. Call checkStyle first. */
export function queueStyle(shape: PowerPoint.Shape, style: ShapeStyle, linkStart?: number): void {
  if (style.name !== undefined) shape.name = style.name;
  if (style.left !== undefined) shape.left = style.left;
  if (style.top !== undefined) shape.top = style.top;
  if (style.width !== undefined) shape.width = style.width;
  if (style.height !== undefined) shape.height = style.height;
  if (style.rotation !== undefined) shape.rotation = style.rotation;

  const range = shape.textFrame.textRange;
  if (style.text !== undefined) range.text = style.text;
  if (style.runs) {
    range.text = replacedText(style) ?? "";
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
    if (indentLevel !== undefined) range.paragraphFormat.indentLevel = indentLevel;
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
  if (style.hyperlink) {
    const { text, ...link } = style.hyperlink;
    if (text === undefined || linkStart === undefined) shape.setHyperlink(link);
    else range.getSubstring(linkStart, text.length).setHyperlink(link);
  }
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

/** Read-back of one slide after a shape-level write, or the error that stopped it. */
type ReadBack = ShapeInfo[] | { error: unknown };

/**
 * Finishes the read-back the write's own sync carried. When that sync failed, the queued loads went with it,
 * so the slides are read again.
 */
export async function readBack(
  context: PowerPoint.RequestContext,
  lists: ShapeList[],
  pending: PendingShapeLists | undefined,
): Promise<ReadBack[]> {
  try {
    const shapes = await (pending ? pending.read() : readShapeLists(context, lists, "styles"));
    return lists.map((_, i) => shapes[i] ?? []);
  } catch (error) {
    return lists.map(() => ({ error }));
  }
}

/** Receipt for shape-level writes on one slide: the changed shapes as they read back. */
export function shapeReceipt(
  slideId: string,
  read: ReadBack,
  changed: string[],
  warnings: string[],
  deleted: string[] = [],
): WriteReceipt {
  if (!Array.isArray(read))
    return {
      changed,
      verified: {},
      warnings: [...warnings, readBackFailureWarning(read.error)],
      fingerprints: {},
    };
  const byId = new Map(flattenShapes(read).map((shape) => [shape.id, shape]));
  const verified: Record<string, unknown> = {};
  for (const id of changed) {
    const shape = byId.get(id);
    if (deleted.includes(id)) {
      if (shape) warnings.push(`Shape ${id} still exists after delete.`);
      else verified[id] = "deleted";
    } else if (shape) verified[id] = compactShape(shape);
    else warnings.push(`Shape ${id} not found after the write.`);
  }
  return { changed, verified, warnings, fingerprints: { [slideId]: fingerprintShapes(read) } };
}

function createShape(shapes: ShapeList, args: AddShapeArgs): PowerPoint.Shape {
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

/** One update with the lookups its checks need, all loaded in the guard's sync. */
interface ShapeTarget {
  slideId: string;
  update: ShapeUpdate;
  shape: PowerPoint.Shape;
  /** Loaded when the update touches text and getTextFrameOrNullObject exists (1.10). */
  frame?: PowerPoint.TextFrame;
}

function lookUpTarget(shapes: ShapeList, slideId: string, update: ShapeUpdate): ShapeTarget {
  const shape = shapes.getItemOrNullObject(update.shapeId).load("id,type");
  const frame =
    touchesText(update) && supportsApi("1.10")
      ? shape.getTextFrameOrNullObject().load("hasText")
      : undefined;
  return { slideId, update, shape, ...(frame && { frame }) };
}

/** Whether the target can take text edits: a non-null frame (1.10), or a text-bearing type before 1.10. */
function hasTextFrame({ shape, frame }: ShapeTarget): boolean {
  return frame ? !frame.isNullObject : TEXT_TYPES.has(shape.type);
}

/** The shape's text before the write: from the freshness check's read when there was one. */
function knownText(checked: CheckedWrite, { slideId, update }: ShapeTarget): string | undefined {
  return flattenShapes(checked.shapes.get(slideId) ?? []).find(
    (shape) => shape.id === update.shapeId,
  )?.text;
}

/**
 * Loads the current text of targets that link part of text they don't replace and whose text the freshness check
 * didn't read. Rare, so it costs its own read-only sync instead of slowing every write.
 */
async function loadLinkTexts(
  context: PowerPoint.RequestContext,
  checked: CheckedWrite,
  targets: ShapeTarget[],
): Promise<Map<ShapeTarget, string>> {
  const texts = new Map<ShapeTarget, string>();
  const pending: [ShapeTarget, PowerPoint.TextRange][] = [];
  for (const target of targets) {
    const { update, shape } = target;
    if (update.hyperlink?.text === undefined || replacedText(update) !== undefined) continue;
    if (shape.isNullObject || !hasTextFrame(target)) continue;
    const text = knownText(checked, target);
    if (text !== undefined) texts.set(target, text);
    else pending.push([target, shape.textFrame.textRange.load("text")]);
  }
  if (pending.length === 0) return texts;
  await context.sync();
  for (const [target, range] of pending) texts.set(target, range.text);
  return texts;
}

/** Merges per-slide receipts into one; update_shapes keys `verified` by slide because shape IDs repeat across slides. */
function mergeReceipts(
  receipts: { slideId: string; receipt: WriteReceipt }[],
  warnings: string[],
  snapshots: CheckedWrite["snapshots"],
): WriteReceipt {
  return {
    changed: receipts.flatMap(({ receipt }) => receipt.changed),
    verified: Object.fromEntries(
      receipts.map(({ slideId, receipt }) => [slideId, receipt.verified]),
    ),
    warnings: [...warnings, ...receipts.flatMap(({ receipt }) => receipt.warnings)],
    fingerprints: Object.assign({}, ...receipts.map(({ receipt }) => receipt.fingerprints)),
    snapshots,
  };
}

/** Adds the snapshots to a receipt; only when there are any, so unguarded receipts stay unchanged. */
export const withSnapshots = (receipt: WriteReceipt, { snapshots }: CheckedWrite): WriteReceipt =>
  Object.keys(snapshots).length > 0 ? { ...receipt, snapshots } : receipt;

export const shapeOps = {
  add_shape: ({ expectedFingerprints, snapshotSlideIds, ...args }: AddShapeArgs & WriteGuardArgs) =>
    PowerPoint.run(async (context) => {
      const { slideId, kind, geometry: _, connector: __, imageBase64, ...style } = args;
      if (kind === "image" && !imageBase64) throw new Error("kind image needs an image.");
      if (imageBase64) requireApi("1.8", "Images");
      const linkStart = checkStyle(style, args.text ?? "");
      const guard = guardWrite(context, [slideId], { expectedFingerprints, snapshotSlideIds });
      const shapes = context.presentation.slides.getItem(slideId).shapes.load("items/id");
      const checked = await guard.check();
      const before = new Set(shapes.items.map((shape) => shape.id));

      const shape = createShape(shapes, args).load("id");
      if (imageBase64) {
        shape.fill.setImage(imageBase64);
        shape.lineFormat.visible = false;
      }
      queueStyle(shape, style, linkStart);
      let pending: PendingShapeLists | undefined = queueShapeLists(context, [shapes], "styles", [
        textShapeIds(checked, slideId),
      ]);
      const warnings: string[] = [];
      let shapeId: string;
      try {
        await context.sync();
        shapeId = shape.id;
      } catch (error) {
        // Creation and styling share one batch; a styling failure leaves the new shape in place.
        const after = shapes.load("items/id");
        await context.sync();
        const created = after.items.find((item) => !before.has(item.id));
        if (!created) throw error;
        shapeId = created.id;
        warnings.push(`Shape ${shapeId} was created but styling failed: ${describeError(error)}`);
        pending = undefined;
      }
      const [read = []] = await readBack(context, [shapes], pending);
      return withSnapshots(shapeReceipt(slideId, read, [shapeId], warnings), checked);
    }),

  /** Applies every update on every listed slide in one batch, so the edits appear together. */
  update_shapes: ({
    slides: edits,
    ...guardArgs
  }: { slides: SlideShapeUpdates[] } & WriteGuardArgs) =>
    PowerPoint.run(async (context) => {
      const guard = guardWrite(
        context,
        edits.map((edit) => edit.slideId),
        guardArgs,
      );
      const lists = edits.map(({ slideId }) => context.presentation.slides.getItem(slideId).shapes);
      const targets = edits.flatMap(({ slideId, updates }, i) =>
        updates.map((update) => lookUpTarget(lists[i]!, slideId, update)),
      );
      const checked = await guard.check();
      const linkTexts = await loadLinkTexts(context, checked, targets);

      const warnings: string[] = [];
      const perSlide = new Map(
        edits.map(({ slideId }) => [
          slideId,
          {
            changed: [] as string[],
            deleted: [] as string[],
            textIds: textShapeIds(checked, slideId),
          },
        ]),
      );
      for (const target of targets) {
        const { slideId, update, shape } = target;
        const { shapeId, delete: remove, zOrder, ...style } = update;
        const slide = perSlide.get(slideId)!;
        if (shape.isNullObject) {
          warnings.push(shapeNotFound(shapeId, slideId).message);
          continue;
        }
        try {
          // Touching a deleted shape would fail the whole batch, not just this update.
          if (slide.deleted.includes(shapeId))
            throw new Error("It is deleted earlier in this call.");
          if (remove) {
            shape.delete();
            slide.deleted.push(shapeId);
          } else {
            if (touchesText(style) && !hasTextFrame(target))
              throw new Error(`A ${shape.type} has no text, so text and font fields don't apply.`);
            const linkStart = checkStyle(style, linkTexts.get(target));
            if (zOrder) {
              requireApi("1.8", "Z-order");
              shape.setZOrder(zOrder);
            }
            queueStyle(shape, style, linkStart);
            if (touchesText(style)) slide.textIds.add(shapeId);
          }
          slide.changed.push(shapeId);
        } catch (error) {
          warnings.push(`Shape ${shapeId} on slide ${slideId}: ${describeError(error)}`);
        }
      }

      const slideOf = (i: number) => perSlide.get(edits[i]!.slideId)!;
      let pending: PendingShapeLists | undefined = queueShapeLists(
        context,
        lists,
        "styles",
        lists.map((_, i) => slideOf(i).textIds),
      );
      try {
        await context.sync();
      } catch (error) {
        warnings.push(batchFailureWarning(error));
        pending = undefined;
      }
      const reads = await readBack(context, lists, pending);
      const receipts = edits.map(({ slideId }, i) => ({
        slideId,
        receipt: shapeReceipt(slideId, reads[i]!, slideOf(i).changed, [], slideOf(i).deleted),
      }));
      return mergeReceipts(receipts, warnings, checked.snapshots);
    }),

  group_shapes: ({
    slideId,
    shapeIds,
    name,
    ...guardArgs
  }: { slideId: string; shapeIds: string[]; name?: string } & WriteGuardArgs) =>
    PowerPoint.run(async (context) => {
      requireApi("1.8", "Grouping");
      const guard = guardWrite(context, [slideId], guardArgs);
      const shapes = context.presentation.slides.getItem(slideId).shapes;
      const checked = await guard.check();
      const group = shapes.addGroup(shapeIds).load("id");
      if (name) group.name = name;
      const pending = queueShapeLists(context, [shapes], "styles", [
        textShapeIds(checked, slideId),
      ]);
      await context.sync();
      const [read = []] = await readBack(context, [shapes], pending);
      return withSnapshots(shapeReceipt(slideId, read, [group.id], []), checked);
    }),

  ungroup_shape: ({
    slideId,
    shapeId,
    ...guardArgs
  }: { slideId: string; shapeId: string } & WriteGuardArgs) =>
    PowerPoint.run(async (context) => {
      requireApi("1.8", "Ungrouping");
      const guard = guardWrite(context, [slideId], guardArgs);
      const shapes = context.presentation.slides.getItem(slideId).shapes;
      const shape = shapes.getItemOrNullObject(shapeId).load("id");
      const checked = await guard.check();
      if (shape.isNullObject) throw shapeNotFound(shapeId, slideId);
      const children = shape.group.shapes.load("items/id");
      await context.sync();
      const childIds = children.items.map((child) => child.id);
      shape.group.ungroup();
      const pending = queueShapeLists(context, [shapes], "styles", [
        textShapeIds(checked, slideId),
      ]);
      await context.sync();
      const [read = []] = await readBack(context, [shapes], pending);
      return withSnapshots(shapeReceipt(slideId, read, childIds, []), checked);
    }),
};
