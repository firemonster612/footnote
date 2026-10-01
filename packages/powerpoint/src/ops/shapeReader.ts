import { supportsApi } from "./presentation.ts";
import type { FontInfo, ShapeInfo } from "./types.ts";

type ShapeList = PowerPoint.ShapeCollection | PowerPoint.ShapeScopedCollection;

// Shape types that carry a text frame and a fill when getTextFrameOrNullObject (1.10) isn't available.
export const TEXT_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder", "Callout"]);
const LINE_TYPES = new Set([...TEXT_TYPES, "Line", "Image"]);
/** Types that always have a text frame, so their text can load before the frame is checked. */
const CERTAIN_TEXT_TYPES = new Set(["GeometricShape", "TextBox"]);
const MAX_GROUP_DEPTH = 2;
const MAX_PARAGRAPHS = 30;
const FONT_FIELDS = "name,size,color,bold,italic,underline";

/**
 * How much of each shape a read loads.
 * summary: IDs, bounds, text, placeholder types and tables; enough for outlines and fingerprints.
 * styles: adds z-order, rotation, fonts, paragraph formatting, fill and line.
 * paragraphs: adds per-paragraph fonts for text whose font is mixed (one more sync when there is any).
 */
export type ShapeReadLevel = "summary" | "styles" | "paragraphs";

/** Finishes a queued read. Call it after a sync has carried the queued loads. */
export interface PendingShapeLists {
  read(): Promise<ShapeInfo[][]>;
}

interface Node {
  shape: PowerPoint.Shape;
  /** IDs of shapes on this node's slide known to have a text frame. */
  knownText?: ReadonlySet<string>;
  frame?: PowerPoint.TextFrame;
  range?: PowerPoint.TextRange;
  placeholder?: PowerPoint.PlaceholderFormat;
  table?: PowerPoint.Table;
  fill?: PowerPoint.ShapeFill;
  line?: PowerPoint.ShapeLineFormat;
  childList?: PowerPoint.ShapeScopedCollection;
  children?: Node[];
  paragraphs?: { text: string; range: PowerPoint.TextRange }[];
}

/** Reads the shape trees of several shape lists in a fixed number of syncs per group level. */
export async function readShapeLists(
  context: PowerPoint.RequestContext,
  lists: ShapeList[],
  level: ShapeReadLevel,
): Promise<ShapeInfo[][]> {
  const pending = queueShapeLists(context, lists, level);
  await context.sync();
  return pending.read();
}

/**
 * Queues the first load of a shape-tree read, so it can ride on the caller's next sync (a write's own sync reads back
 * its result). `textShapeIds[i]` names shapes in `lists[i]` known to have a text frame, from an earlier read: their
 * text loads with their details instead of in a separate pass. Shape IDs are per slide, so the sets are too.
 */
export function queueShapeLists(
  context: PowerPoint.RequestContext,
  lists: ShapeList[],
  level: ShapeReadLevel,
  textShapeIds: readonly ReadonlySet<string>[] = [],
): PendingShapeLists {
  const detail = level !== "summary";
  for (const list of lists) loadBasics(list, detail);
  return {
    async read() {
      const roots = lists.map((list, i) =>
        list.items.map((shape): Node => ({ shape, knownText: textShapeIds[i] })),
      );
      let nodes = roots.flat();
      if (nodes.length === 0) return roots.map(() => []);
      const hasKnownFrame = (node: Node) =>
        CERTAIN_TEXT_TYPES.has(node.shape.type) || node.knownText?.has(node.shape.id) === true;
      const queueLevel = (batch: Node[], depth: number) => {
        for (const node of batch) {
          queueDetails(node, detail, depth);
          if (hasKnownFrame(node)) queueText(node, detail);
        }
      };
      queueLevel(nodes, 0);
      await context.sync();

      // Each pass loads the text this level still needs together with the next group level's details.
      for (let depth = 1; nodes.length > 0; depth++) {
        let queued = false;
        const next: Node[] = [];
        for (const node of nodes) {
          if (!node.range && node.frame && !node.frame.isNullObject && node.frame.hasText) {
            queueText(node, detail);
            queued = true;
          }
          if (!node.childList) continue;
          node.children = node.childList.items.map((shape): Node => ({
            shape,
            knownText: node.knownText,
          }));
          next.push(...node.children);
        }
        queueLevel(next, depth);
        if (queued || next.length > 0) await context.sync();

        if (level === "paragraphs" && nodes.map(queueParagraphs).includes(true))
          await context.sync();
        nodes = next;
      }
      return roots.map((shapes) => shapes.map((node) => toShapeInfo(node, detail)));
    },
  };
}

function loadBasics(list: ShapeList, detail: boolean): void {
  const fields = ["id", "name", "type", "left", "top", "width", "height"];
  if (detail) fields.push("zOrderPosition");
  if (detail && supportsApi("1.10")) fields.push("rotation");
  list.load(fields.map((field) => `items/${field}`).join(","));
}

function queueDetails(node: Node, detail: boolean, depth: number): void {
  const { shape } = node;
  const type = shape.type;
  if (supportsApi("1.10")) node.frame = shape.getTextFrameOrNullObject();
  else if (TEXT_TYPES.has(type)) node.frame = shape.textFrame;
  node.frame?.load(detail ? "hasText,autoSizeSetting" : "hasText");
  if (type === "Placeholder") node.placeholder = shape.placeholderFormat.load("type");
  if (type === "Table") node.table = shape.getTable().load("rowCount,columnCount,values");
  if (type === "Group" && depth < MAX_GROUP_DEPTH) {
    node.childList = shape.group.shapes;
    loadBasics(node.childList, detail);
  }
  if (detail && TEXT_TYPES.has(type))
    node.fill = shape.fill.load("type,foregroundColor,transparency");
  if (detail && LINE_TYPES.has(type))
    node.line = shape.lineFormat.load("visible,color,weight,dashStyle");
}

function queueText(node: Node, detail: boolean): void {
  if (!node.frame) return;
  node.range = node.frame.textRange.load("text");
  if (!detail) return;
  node.range.font.load(FONT_FIELDS);
  node.range.paragraphFormat.load("horizontalAlignment");
  node.range.paragraphFormat.bulletFormat.load("visible");
}

/** The shape's text range, when it has text. Shapes with a known frame load their range before hasText is known. */
function textOf(node: Node): PowerPoint.TextRange | undefined {
  return node.range && node.frame?.hasText ? node.range : undefined;
}

/** For text whose font is mixed, loads the font of each paragraph. Returns whether anything was queued. */
function queueParagraphs(node: Node): boolean {
  const range = textOf(node);
  if (!range || !isMixed(fontInfo(range.font))) return false;
  node.paragraphs = paragraphSpans(range.text)
    .slice(0, MAX_PARAGRAPHS)
    .map((span) => ({ text: span.text, range: range.getSubstring(span.start, span.text.length) }));
  for (const paragraph of node.paragraphs) paragraph.range.font.load(FONT_FIELDS);
  return node.paragraphs.length > 0;
}

/** Non-empty paragraphs of a text range with their offsets. PowerPoint separates paragraphs with \r or \n (\v is a line break). */
export function paragraphSpans(text: string): { start: number; text: string }[] {
  const spans: { start: number; text: string }[] = [];
  let start = 0;
  for (const part of text.split(/\r\n|\r|\n/)) {
    if (part.trim()) spans.push({ start, text: part });
    start += part.length + (text.startsWith("\r\n", start + part.length) ? 2 : 1);
  }
  return spans;
}

function isMixed(font: FontInfo): boolean {
  return Object.values(font).some((value) => value === null);
}

function fontInfo(font: PowerPoint.ShapeFont): FontInfo {
  return {
    name: font.name,
    size: font.size,
    color: font.color,
    bold: font.bold,
    italic: font.italic,
    underline: font.underline,
  };
}

const round = (value: number): number => Math.round(value * 10) / 10;

function toShapeInfo(node: Node, detail: boolean): ShapeInfo {
  const { shape } = node;
  const info: ShapeInfo = {
    id: shape.id,
    name: shape.name,
    type: shape.type,
    left: round(shape.left),
    top: round(shape.top),
    width: round(shape.width),
    height: round(shape.height),
  };
  if (detail) {
    info.z = shape.zOrderPosition;
    if (supportsApi("1.10") && shape.rotation) info.rotation = shape.rotation;
  }
  if (node.placeholder) info.placeholder = node.placeholder.type;
  const range = textOf(node);
  if (range) info.text = range.text;
  if (detail && node.frame && !node.frame.isNullObject) info.autoSize = node.frame.autoSizeSetting;
  if (detail && range) {
    info.font = fontInfo(range.font);
    info.align = range.paragraphFormat.horizontalAlignment;
    info.bullets = range.paragraphFormat.bulletFormat.visible;
  }
  if (node.paragraphs)
    info.paragraphs = node.paragraphs.map((p) => ({ text: p.text, font: fontInfo(p.range.font) }));
  if (node.fill) {
    info.fill =
      node.fill.type === "Solid"
        ? {
            type: node.fill.type,
            color: node.fill.foregroundColor,
            transparency: node.fill.transparency,
          }
        : { type: node.fill.type };
  }
  if (node.line?.visible)
    info.line = { color: node.line.color, weight: node.line.weight, dash: node.line.dashStyle };
  if (node.table)
    info.table = {
      rows: node.table.rowCount,
      columns: node.table.columnCount,
      values: node.table.values,
    };
  if (node.children) info.children = node.children.map((child) => toShapeInfo(child, detail));
  return info;
}
