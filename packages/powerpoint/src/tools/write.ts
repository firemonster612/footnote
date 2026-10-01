import type { FootnoteTool, ToolEnv } from "@footnote/core/contracts";
import { Type } from "typebox";
import type {
  AddShapeArgs,
  EditTableArgs,
  ShapeUpdate,
  SlideDetail,
  WriteReceipt,
} from "../ops/types.ts";
import type { WriteGuard } from "../writeGuard.ts";
import { plural, receiptResult } from "./results.ts";
import {
  boundsProperties,
  Color,
  defineTool,
  FontStyle,
  Position,
  ShapeId,
  SlideId,
  shapeStyleProperties,
} from "./schemas.ts";

const MAX_UPDATES = 50;
const sequential = "sequential" as const;

export function createWriteTools(env: ToolEnv, guard: WriteGuard): FootnoteTool[] {
  const { host } = env;

  /** Runs a core write op, then records its receipt (undo bookkeeping, fresh fingerprints). */
  async function write(op: string, args: object): Promise<WriteReceipt> {
    const receipt = await host.call<WriteReceipt>(op, args);
    guard.record(receipt);
    return receipt;
  }

  function imageFromAttachment(attachmentId: string): string {
    const attachment = env.attachments.get(attachmentId);
    if (!attachment)
      throw new Error(
        `Attachment ${attachmentId} not found. Attachment IDs are listed with the user's message.`,
      );
    if (attachment.kind !== "image")
      throw new Error(`Attachment ${attachment.name} is a ${attachment.kind}, not an image.`);
    return attachment.base64;
  }

  const addSlide = defineTool({
    name: "add_slide",
    label: "Add slide",
    access: "write",
    executionMode: sequential,
    description:
      "Adds a slide built on one of the deck's layouts (layout IDs from get_deck; prefer a layout whose placeholders match the content). " +
      "Returns the new slide's placeholder shapes so you can fill them with update_shapes in one follow-up call.",
    parameters: Type.Object({
      layoutId: Type.Optional(
        Type.String({
          description: "Layout ID from get_deck. Default: the master's first layout.",
        }),
      ),
      slideMasterId: Type.Optional(
        Type.String({ description: "Only needed when the deck has several masters." }),
      ),
      position: Type.Optional(Position),
    }),
    describeCall: (args: { position?: number }) =>
      args.position ? `Add a slide at position ${args.position}` : "Add a slide at the end",
    async execute(_id, { layoutId, slideMasterId, position }) {
      const receipt = await write("add_slide", {
        layoutId,
        slideMasterId,
        ...(position && { index: position - 1 }),
      });
      const [slideId] = receipt.createdSlideIds ?? [];
      if (!slideId) return receiptResult(receipt);
      const slide = await host.call<SlideDetail>("get_slide", { slideId });
      guard.observe(slide.id, slide.fingerprint);
      const placeholders = slide.shapes.map(
        ({ id, name, placeholder, left, top, width, height }) => ({
          id,
          name,
          placeholder,
          left,
          top,
          width,
          height,
        }),
      );
      return receiptResult(receipt, { shapes: placeholders });
    },
  });

  const deleteSlides = defineTool({
    name: "delete_slides",
    label: "Delete slides",
    access: "write",
    executionMode: sequential,
    description:
      "Deletes slides by ID. Positions of later slides shift; use IDs, not positions, in follow-up calls.",
    parameters: Type.Object({ slideIds: Type.Array(SlideId, { minItems: 1 }) }),
    describeCall: (args: { slideIds: string[] }) =>
      `Delete ${plural(args.slideIds.length, "slide")}`,
    async execute(_id, { slideIds }) {
      await guard.beforeWrite(slideIds);
      return receiptResult(await write("delete_slides", { slideIds }));
    },
  });

  const moveSlide = defineTool({
    name: "move_slide",
    label: "Move slide",
    access: "write",
    executionMode: sequential,
    description: "Moves a slide to a new 1-based position. The slide keeps its ID.",
    parameters: Type.Object({ slideId: SlideId, position: Position }),
    describeCall: (args: { slideId: string; position: number }) =>
      `Move slide ${args.slideId} to position ${args.position}`,
    async execute(_id, { slideId, position }) {
      await guard.checkpoint([slideId]);
      return receiptResult(await write("move_slide", { slideId, toIndex: position - 1 }));
    },
  });

  const duplicateSlide = defineTool({
    name: "duplicate_slide",
    label: "Duplicate slide",
    access: "write",
    executionMode: sequential,
    description:
      "Copies a slide with all its content. The copy lands right after the original unless you give a position. " +
      "Good for building a series of slides that share a structure: duplicate, then edit the copy.",
    parameters: Type.Object({ slideId: SlideId, position: Type.Optional(Position) }),
    describeCall: (args: { slideId: string }) => `Duplicate slide ${args.slideId}`,
    async execute(_id, { slideId, position }) {
      return receiptResult(
        await write("duplicate_slide", { slideId, ...(position && { toIndex: position - 1 }) }),
      );
    },
  });

  const applyLayout = defineTool({
    name: "apply_layout",
    label: "Apply layout",
    access: "write",
    executionMode: sequential,
    description:
      "Switches a slide to another layout of its master (layout IDs from get_deck). Content in placeholders the new layout lacks may move or disappear; render to check.",
    parameters: Type.Object({ slideId: SlideId, layoutId: Type.String() }),
    describeCall: (args: { slideId: string }) => `Change the layout of slide ${args.slideId}`,
    async execute(_id, { slideId, layoutId }) {
      await guard.beforeWrite([slideId]);
      return receiptResult(await write("apply_layout", { slideId, layoutId }));
    },
  });

  const addShape = defineTool({
    name: "add_shape",
    label: "Add shape",
    access: "write",
    executionMode: sequential,
    description:
      "Adds one shape to a slide: a text box, a geometric shape (geometry: Rectangle, RoundRectangle, Ellipse, Triangle, RightArrow, Chevron, " +
      "Pentagon, Star5, …), a line (from left,top to left+width,top+height), or an image from an attachment (keep its aspect ratio). " +
      "Style it in the same call. Prefer filling an existing placeholder with update_shapes over adding text boxes. Units are points.",
    parameters: Type.Object({
      slideId: SlideId,
      kind: Type.Enum(["textbox", "geometric", "line", "image"]),
      geometry: Type.Optional(
        Type.String({ description: "GeometricShapeType for kind geometric. Default Rectangle." }),
      ),
      connector: Type.Optional(
        Type.Enum(["Straight", "Elbow", "Curve"], {
          description: "For kind line. Default Straight.",
        }),
      ),
      attachmentId: Type.Optional(Type.String({ description: "Image attachment for kind image." })),
      ...boundsProperties,
      ...shapeStyleProperties,
    }),
    describeCall: (args: { kind: string; slideId: string }) =>
      `Add a ${args.kind} to slide ${args.slideId}`,
    async execute(_id, { attachmentId, ...params }) {
      await guard.beforeWrite([params.slideId]);
      const args: AddShapeArgs = {
        ...params,
        ...(attachmentId && { imageBase64: imageFromAttachment(attachmentId) }),
      };
      return receiptResult(await write("add_shape", args));
    },
  });

  const updateShapes = defineTool({
    name: "update_shapes",
    label: "Update shapes",
    access: "write",
    executionMode: sequential,
    description:
      "Batch-edits shapes on one slide: text (plain or formatted runs), font, paragraph alignment/bullets, text frame, fill, line, " +
      "bounds, rotation, z-order, name, hyperlink, or delete. Put every change for the slide in one call. " +
      "Only listed fields change. Use theme fonts and colors unless asked otherwise; keep text inside its shape.",
    parameters: Type.Object({
      slideId: SlideId,
      updates: Type.Array(
        Type.Object({
          shapeId: ShapeId,
          delete: Type.Optional(
            Type.Boolean({ description: "Removes the shape; other fields are ignored." }),
          ),
          zOrder: Type.Optional(
            Type.Enum(["BringToFront", "SendToBack", "BringForward", "SendBackward"]),
          ),
          left: Type.Optional(boundsProperties.left),
          top: Type.Optional(boundsProperties.top),
          width: Type.Optional(boundsProperties.width),
          height: Type.Optional(boundsProperties.height),
          ...shapeStyleProperties,
        }),
        { minItems: 1, maxItems: MAX_UPDATES },
      ),
    }),
    describeCall: (args: { slideId: string; updates: unknown[] }) =>
      `Update ${plural(args.updates.length, "shape")} on slide ${args.slideId}`,
    async execute(_id, { slideId, updates }) {
      await guard.beforeWrite([slideId]);
      const typed: ShapeUpdate[] = updates;
      return receiptResult(await write("update_shapes", { slideId, updates: typed }));
    },
  });

  const editTable = defineTool({
    name: "edit_table",
    label: "Edit table",
    access: "write",
    executionMode: sequential,
    description:
      "Creates a table (create) or edits one (shapeId): insert/delete rows and columns, set cell text, font, fill and alignment, merge cells, set column widths. " +
      "Row and column indexes are 0-based and refer to the grid after inserts and deletes, which apply first. " +
      "Use a table for exact values people compare; use insert_chart for trends and proportions.",
    parameters: Type.Object({
      slideId: SlideId,
      shapeId: Type.Optional(
        Type.String({ description: "Existing table shape ID. Omit when creating." }),
      ),
      create: Type.Optional(
        Type.Object({
          rows: Type.Integer({ minimum: 1, maximum: 75 }),
          columns: Type.Integer({ minimum: 1, maximum: 75 }),
          ...boundsProperties,
          values: Type.Optional(
            Type.Array(Type.Array(Type.String()), { description: "Rows of cell text." }),
          ),
          style: Type.Optional(
            Type.String({
              description:
                'TableStyle, e.g. "MediumStyle2Accent1", "LightStyle1Accent1", "NoStyleTableGrid".',
            }),
          ),
          name: Type.Optional(Type.String()),
        }),
      ),
      insertRows: Type.Optional(
        Type.Object({ index: Type.Integer({ minimum: 0 }), count: Type.Integer({ minimum: 1 }) }),
      ),
      insertColumns: Type.Optional(
        Type.Object({ index: Type.Integer({ minimum: 0 }), count: Type.Integer({ minimum: 1 }) }),
      ),
      deleteRows: Type.Optional(Type.Array(Type.Integer({ minimum: 0 }))),
      deleteColumns: Type.Optional(Type.Array(Type.Integer({ minimum: 0 }))),
      cells: Type.Optional(
        Type.Array(
          Type.Object({
            row: Type.Integer({ minimum: 0 }),
            column: Type.Integer({ minimum: 0 }),
            text: Type.Optional(Type.String()),
            font: Type.Optional(FontStyle),
            fill: Type.Optional(Color),
            align: Type.Optional(Type.Enum(["Left", "Center", "Right", "Justify"])),
          }),
        ),
      ),
      merge: Type.Optional(
        Type.Array(
          Type.Object({
            row: Type.Integer({ minimum: 0 }),
            column: Type.Integer({ minimum: 0 }),
            rowCount: Type.Integer({ minimum: 1 }),
            columnCount: Type.Integer({ minimum: 1 }),
          }),
        ),
      ),
      columnWidths: Type.Optional(
        Type.Array(Type.Number({ minimum: 1 }), { description: "Points, by column index." }),
      ),
    }),
    describeCall: (args: { slideId: string; create?: unknown }) =>
      args.create
        ? `Add a table to slide ${args.slideId}`
        : `Edit a table on slide ${args.slideId}`,
    async execute(_id, params) {
      await guard.beforeWrite([params.slideId]);
      const args: EditTableArgs = params;
      return receiptResult(await write("edit_table", args));
    },
  });

  const groupShapes = defineTool({
    name: "group_shapes",
    label: "Group shapes",
    access: "write",
    executionMode: sequential,
    description:
      "Groups two or more shapes on a slide into one (action group), or splits a group back into its shapes (action ungroup, one group ID).",
    parameters: Type.Object({
      slideId: SlideId,
      action: Type.Enum(["group", "ungroup"]),
      shapeIds: Type.Array(ShapeId, { minItems: 1 }),
      name: Type.Optional(Type.String({ description: "Name for the new group." })),
    }),
    describeCall: (args: { action: string; shapeIds: string[]; slideId: string }) =>
      args.action === "group"
        ? `Group ${plural(args.shapeIds.length, "shape")} on slide ${args.slideId}`
        : `Ungroup shape ${args.shapeIds[0]}`,
    async execute(_id, { slideId, action, shapeIds, name }) {
      await guard.beforeWrite([slideId]);
      if (action === "group") {
        if (shapeIds.length < 2) throw new Error("Grouping needs at least two shape IDs.");
        return receiptResult(
          await write("group_shapes", { slideId, shapeIds, ...(name && { name }) }),
        );
      }
      if (shapeIds.length !== 1) throw new Error("Ungroup takes exactly one group shape ID.");
      return receiptResult(await write("ungroup_shape", { slideId, shapeId: shapeIds[0] }));
    },
  });

  return [
    addSlide,
    deleteSlides,
    moveSlide,
    duplicateSlide,
    applyLayout,
    addShape,
    updateShapes,
    editTable,
    groupShapes,
  ];
}
