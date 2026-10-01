import type { FootnoteTool, ToolEnv } from "@footnote/core/contracts";
import { Type } from "typebox";
import type {
  DeckInfo,
  LayoutsInfo,
  SelectionInfo,
  ShapeInfo,
  SlideDetail,
  ThemeInfo,
} from "../ops/types.ts";
import type { WriteGuard } from "../writeGuard.ts";
import { jsonResult } from "./results.ts";
import { defineTool, SlideId, type SlidePositions, slideLabel } from "./schemas.ts";

const MAX_SHAPE_TEXT = 2_000;
const MAX_TABLE_ROWS = 30;
const SLIDE_BUDGET_CHARS = 12_000;
const DEFAULT_SHAPE_LIMIT = 40;

function truncateShape(shape: ShapeInfo): ShapeInfo {
  const { text, table, children } = shape;
  return {
    ...shape,
    ...(text &&
      text.length > MAX_SHAPE_TEXT && {
        text: `${text.slice(0, MAX_SHAPE_TEXT)}… (${text.length} chars)`,
      }),
    ...(table &&
      table.rows > MAX_TABLE_ROWS && {
        table: {
          ...table,
          values: [
            ...table.values.slice(0, MAX_TABLE_ROWS),
            [`… ${table.rows - MAX_TABLE_ROWS} more rows`],
          ],
        },
      }),
    ...(children && { children: children.map(truncateShape) }),
  };
}

/** Takes shapes from `offset` until the limit or the character budget, whichever comes first (always at least one). */
export function pageShapes(
  shapes: ShapeInfo[],
  offset: number,
  limit: number,
): { page: ShapeInfo[]; nextOffset?: number } {
  const page: ShapeInfo[] = [];
  let chars = 0;
  for (const shape of shapes.slice(offset, offset + limit).map(truncateShape)) {
    chars += JSON.stringify(shape).length;
    if (page.length > 0 && chars > SLIDE_BUDGET_CHARS) break;
    page.push(shape);
  }
  const next = offset + page.length;
  return next < shapes.length ? { page, nextOffset: next } : { page };
}

export function createReadTools(
  env: ToolEnv,
  guard: WriteGuard,
  positions: SlidePositions,
): FootnoteTool[] {
  const { host } = env;

  const getDeck = defineTool({
    name: "get_deck",
    label: "Read deck",
    access: "read",
    description:
      "Deck overview: slide size, every slide (position, ID, layout, title, shape count), theme fonts and colors, and each master's layouts with IDs. " +
      "<deck_state> already gives the outline each turn; call this when you need layout IDs (add_slide, apply_layout), the theme, or the full outline of a long deck.",
    parameters: Type.Object({}),
    describeCall: () => "Read deck outline, theme and layouts",
    async execute() {
      const deck = await host.call<DeckInfo>("get_deck");
      const theme = await host.call<ThemeInfo>("get_theme");
      const layouts = await host.call<LayoutsInfo>("get_layouts");
      return jsonResult({
        slideWidth: deck.slideWidth,
        slideHeight: deck.slideHeight,
        slides: deck.slides.map(({ id, index, layout, title, shapeCount }) => ({
          position: index + 1,
          id,
          layout,
          title,
          shapes: shapeCount,
        })),
        theme,
        masters: layouts.masters,
      });
    },
  });

  const getSlide = defineTool({
    name: "get_slide",
    label: "Read slide",
    access: "read",
    description:
      "Shape tree of one slide: shape IDs, names, types, bounds and z-order (points), placeholder type, text with font/size/color/bold " +
      "(per paragraph when mixed), alignment, bullets, fill, line, table cells and group children. " +
      "Call it before editing a slide: write tools fail when the slide changed since you last read it. Large slides are paged; follow nextOffset.",
    parameters: Type.Object({
      slideId: SlideId,
      offset: Type.Optional(
        Type.Integer({ minimum: 0, description: "Shape offset for paging. Default 0." }),
      ),
      limit: Type.Optional(
        Type.Integer({
          minimum: 1,
          maximum: 100,
          description: `Max shapes. Default ${DEFAULT_SHAPE_LIMIT}.`,
        }),
      ),
    }),
    describeCall: (args: { slideId: string }) => `Read ${slideLabel(positions, args.slideId)}`,
    async execute(_id, { slideId, offset = 0, limit = DEFAULT_SHAPE_LIMIT }) {
      const slide = await host.call<SlideDetail>("get_slide", { slideId });
      guard.observe(slide.id, slide.fingerprint);
      const { page, nextOffset } = pageShapes(slide.shapes, offset, limit);
      return jsonResult({
        id: slide.id,
        position: slide.index + 1,
        layout: slide.layout,
        totalShapes: slide.shapes.length,
        shapes: page,
        ...(nextOffset !== undefined && { nextOffset }),
      });
    },
  });

  const getSelection = defineTool({
    name: "get_selection",
    label: "Read selection",
    access: "read",
    description:
      "What the user has selected right now: slides, shapes (with full detail) and selected text. " +
      'Use when the user says "this slide", "these boxes" or "the selected text" and <deck_state> may be out of date.',
    parameters: Type.Object({}),
    describeCall: () => "Read the current selection",
    async execute() {
      const selection = await host.call<SelectionInfo>("get_selection");
      const [slideId] = selection.slideIds;
      if (!slideId) return jsonResult({ slides: [], shapes: [] });
      const slide = await host.call<SlideDetail>("get_slide", { slideId });
      guard.observe(slide.id, slide.fingerprint);
      return jsonResult({
        slides: selection.slideIds,
        position: slide.index + 1,
        shapes: slide.shapes
          .filter((shape) => selection.shapeIds.includes(shape.id))
          .map(truncateShape),
        ...(selection.text && { text: selection.text }),
      });
    },
  });

  const renderSlide = defineTool({
    name: "render_slide",
    label: "Render slide",
    access: "read",
    description:
      "Image of a slide as it looks now. Use it to check layout work: overflow, overlap, alignment, contrast, empty placeholders. " +
      "Render after visual changes rather than assuming they look right; skip it for pure text edits you already verified in the receipt.",
    parameters: Type.Object({
      slideId: SlideId,
      width: Type.Optional(
        Type.Integer({ minimum: 320, maximum: 1280, description: "Pixels. Default 1280." }),
      ),
    }),
    describeCall: (args: { slideId: string }) => `Render ${slideLabel(positions, args.slideId)}`,
    async execute(_id, { slideId, width }) {
      const data = await host.call<unknown>("render_slide", { slideId, ...(width && { width }) });
      // An image block with non-string data poisons every later request (Anthropic rejects the whole transcript).
      if (typeof data !== "string" || data.length === 0)
        throw new Error(
          `PowerPoint returned no image for slide ${slideId}. Try again, or check the slide with get_slide.`,
        );
      return {
        content: [
          { type: "text", text: `Slide ${slideId}:` },
          { type: "image", data, mimeType: "image/png" },
        ],
        details: { slideId },
      };
    },
  });

  return [getDeck, getSlide, getSelection, renderSlide];
}
