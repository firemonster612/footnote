import type { FootnoteTool } from "@footnote/core/contracts";
import { type TSchema, Type } from "typebox";

// Pi resolves Static through its own typebox copy, so take the params type from the tool itself.
type ToolParams<P extends TSchema> = Parameters<FootnoteTool<P>["execute"]>[1];

/** Lets execute see params typed by its schema while the tool joins a list of tools with other schemas. */
export function defineTool<P extends TSchema>(tool: FootnoteTool<P>): FootnoteTool {
  return {
    ...tool,
    // Pi validates arguments against `parameters` before execute runs, so params already match P.
    execute: (toolCallId, params, signal, onUpdate) =>
      tool.execute(toolCallId, params as ToolParams<P>, signal, onUpdate),
  };
}

/** A slide's 1-based position for tool-call labels, when known. */
export type SlidePositions = (slideId: string) => number | undefined;

/** "slide 3" for labels, or the raw ID when the position isn't known. */
export const slideLabel = (positions: SlidePositions, slideId: string): string =>
  `slide ${positions(slideId) ?? slideId}`;

/** "slide 3" or "slides 2, 4, 5". */
export function slidesLabel(positions: SlidePositions, slideIds: string[]): string {
  const [first, ...rest] = slideIds;
  if (first !== undefined && rest.length === 0) return slideLabel(positions, first);
  return `slides ${slideIds.map((id) => positions(id) ?? id).join(", ")}`;
}

export const SlideId = Type.String({
  description: "Slide ID from <deck_state> or get_deck (not the position).",
});
export const ShapeId = Type.String({ description: "Shape ID from get_slide." });
export const Position = Type.Integer({
  minimum: 1,
  description: "1-based slide position in the deck.",
});
export const Color = Type.String({ pattern: "^#[0-9A-Fa-f]{6}$", description: '"#RRGGBB"' });
const Points = (description: string) => Type.Number({ description: `${description} in points.` });
const Align = Type.Enum(["Left", "Center", "Right", "Justify"]);

export const FontStyle = Type.Object({
  name: Type.Optional(Type.String({ description: "Font family. Prefer the theme fonts." })),
  size: Type.Optional(Type.Number({ minimum: 1, description: "Points." })),
  color: Type.Optional(Color),
  bold: Type.Optional(Type.Boolean()),
  italic: Type.Optional(Type.Boolean()),
  underline: Type.Optional(Type.Boolean()),
});

const TextRun = Type.Object({ text: Type.String(), ...FontStyle.properties });

/** Fields shared by add_shape and update_shapes. */
export const shapeStyleProperties = {
  name: Type.Optional(Type.String({ description: "Shape name shown in the selection pane." })),
  rotation: Type.Optional(Type.Number({ description: "Degrees clockwise." })),
  text: Type.Optional(
    Type.String({ description: "Replaces all text. Separate paragraphs with \\n." }),
  ),
  runs: Type.Optional(
    Type.Array(TextRun, {
      minItems: 1,
      description:
        "Replaces all text with runs that carry their own formatting (e.g. a bold lead-in). Use instead of text.",
    }),
  ),
  font: Type.Optional(FontStyle),
  paragraph: Type.Optional(
    Type.Object({
      align: Type.Optional(Align),
      bullets: Type.Optional(Type.Boolean()),
      indentLevel: Type.Optional(Type.Integer({ minimum: 0, maximum: 8 })),
    }),
  ),
  textFrame: Type.Optional(
    Type.Object({
      autoSize: Type.Optional(
        Type.Enum(["AutoSizeNone", "AutoSizeTextToFitShape", "AutoSizeShapeToFitText"]),
      ),
      verticalAlign: Type.Optional(Type.Enum(["Top", "Middle", "Bottom"])),
      wordWrap: Type.Optional(Type.Boolean()),
    }),
  ),
  fill: Type.Optional(
    Type.Object({
      color: Type.String({
        pattern: "^(#[0-9A-Fa-f]{6}|none)$",
        description: '"#RRGGBB", or "none" to remove the fill.',
      }),
      transparency: Type.Optional(Type.Number({ minimum: 0, maximum: 1 })),
    }),
  ),
  line: Type.Optional(
    Type.Object({
      color: Type.Optional(Color),
      weight: Type.Optional(Points("Line width")),
      dash: Type.Optional(Type.Enum(["Solid", "Dash", "RoundDot", "LongDash", "DashDot"])),
      visible: Type.Optional(Type.Boolean({ description: "false removes the outline." })),
    }),
  ),
  hyperlink: Type.Optional(
    Type.Object({
      address: Type.String({ description: "URL." }),
      screenTip: Type.Optional(Type.String()),
      text: Type.Optional(
        Type.String({
          description: "Link only the first occurrence of this text; omit to link the whole shape.",
        }),
      ),
    }),
  ),
};

export const boundsProperties = {
  left: Points("Distance from the slide's left edge"),
  top: Points("Distance from the slide's top edge"),
  width: Points("Width"),
  height: Points("Height"),
};
