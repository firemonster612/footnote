import type { FootnoteTool, ToolEnv } from "@footnote/core/contracts";
import { Type } from "typebox";
import type {
  AddShapeArgs,
  AddSlideReceipt,
  EditTableArgs,
  SlideShapeUpdates,
  WriteReceipt,
} from "../ops/types.ts";
import { modelFacingError, type WriteGuard } from "../writeGuard.ts";
import { plural, receiptResult } from "./results.ts";
import {
  boundsProperties,
  Color,
  defineTool,
  FontStyle,
  Position,
  ShapeId,
  SlideId,
  type SlidePositions,
  slideLabel,
  slidesLabel,
  shapeStyleProperties,
} from "./schemas.ts";

const MAX_UPDATES = 50;
const MAX_UPDATED_SLIDES = 20;
const sequential = "sequential" as const;

export function createWriteTools(
  env: ToolEnv,
  guard: WriteGuard,
  positions: SlidePositions,
): FootnoteTool[] {
  const { host } = env;

  /**
   * Runs a write op in one host call, then records its receipt (undo snapshots, created slides, fresh fingerprints).
   * Guarded ops get `guard.writeArgs` in `args` and check freshness and export snapshots inside their own run.
   */
  async function write<R extends WriteReceipt>(
    op: string,
    args: object,
  ): Promise<Omit<R, "snapshots">> {
    let receipt: R;
    try {
      receipt = await host.call<R>(op, args);
    } catch (error) {
      throw modelFacingError(error);
    }
    guard.record(receipt);
    const { snapshots: _, ...withoutSnapshots } = receipt;
    return withoutSnapshots;
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
      "Returns the new slide's placeholder shapes so you can fill them with update_shapes; when adding several slides, add them all first and fill them in one update_shapes call.",
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
      const { shapes, ...receipt } = await write<AddSlideReceipt>("add_slide", {
        layoutId,
        slideMasterId,
        ...(position && { index: position - 1 }),
      });
      return receiptResult(receipt, { shapes });
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
      `Delete ${slidesLabel(positions, args.slideIds)}`,
    async execute(_id, { slideIds }) {
      return receiptResult(
        await write("delete_slides", { slideIds, ...guard.writeArgs(slideIds) }),
      );
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
      `Move ${slideLabel(positions, args.slideId)} to position ${args.position}`,
    async execute(_id, { slideId, position }) {
      // Moving doesn't overwrite content, so it only needs the undo snapshot, not the freshness check.
      const { snapshotSlideIds } = guard.writeArgs([slideId]);
      return receiptResult(
        await write("move_slide", { slideId, toIndex: position - 1, snapshotSlideIds }),
      );
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
    describeCall: (args: { slideId: string }) => `Duplicate ${slideLabel(positions, args.slideId)}`,
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
    describeCall: (args: { slideId: string }) =>
      `Change the layout of ${slideLabel(positions, args.slideId)}`,
    async execute(_id, { slideId, layoutId }) {
      return receiptResult(
        await write("apply_layout", { slideId, layoutId, ...guard.writeArgs([slideId]) }),
      );
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
      `Add a ${args.kind} to ${slideLabel(positions, args.slideId)}`,
    async execute(_id, { attachmentId, ...params }) {
      const args: AddShapeArgs = {
        ...params,
        ...(attachmentId && { imageBase64: imageFromAttachment(attachmentId) }),
      };
      return receiptResult(
        await write("add_shape", { ...args, ...guard.writeArgs([params.slideId]) }),
      );
    },
  });

  const shapeUpdate = Type.Object({
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
  });

  const updateShapes = defineTool({
    name: "update_shapes",
    label: "Update shapes",
    access: "write",
    executionMode: sequential,
    description:
      "Batch-edits shapes on one or more slides: text (plain or formatted runs), font, paragraph alignment/bullets, text frame, fill, line, " +
      "bounds, rotation, z-order, name, hyperlink, or delete. Everything in one call lands in PowerPoint at once, so put every shape " +
      "change you've planned, across all the slides involved, in a single call rather than one call per slide or shape. " +
      "Only listed fields change. Use theme fonts and colors unless asked otherwise; keep text inside its shape.",
    parameters: Type.Object({
      slides: Type.Array(
        Type.Object({
          slideId: SlideId,
          updates: Type.Array(shapeUpdate, { minItems: 1, maxItems: MAX_UPDATES }),
        }),
        { minItems: 1, maxItems: MAX_UPDATED_SLIDES },
      ),
    }),
    describeCall: (args: { slides: { slideId: string; updates: unknown[] }[] }) => {
      const count = args.slides.reduce((sum, slide) => sum + slide.updates.length, 0);
      const slideIds = args.slides.map((slide) => slide.slideId);
      return `Update ${plural(count, "shape")} on ${slidesLabel(positions, slideIds)}`;
    },
    async execute(_id, { slides }) {
      const edits: SlideShapeUpdates[] = slides;
      const slideIds = [...new Set(edits.map((edit) => edit.slideId))];
      return receiptResult(
        await write("update_shapes", { slides: edits, ...guard.writeArgs(slideIds) }),
      );
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
        ? `Add a table to ${slideLabel(positions, args.slideId)}`
        : `Edit a table on ${slideLabel(positions, args.slideId)}`,
    async execute(_id, params) {
      const args: EditTableArgs = params;
      return receiptResult(
        await write("edit_table", { ...args, ...guard.writeArgs([params.slideId]) }),
      );
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
        ? `Group ${plural(args.shapeIds.length, "shape")} on ${slideLabel(positions, args.slideId)}`
        : `Ungroup a group on ${slideLabel(positions, args.slideId)}`,
    async execute(_id, { slideId, action, shapeIds, name }) {
      const guardArgs = guard.writeArgs([slideId]);
      if (action === "group") {
        if (shapeIds.length < 2) throw new Error("Grouping needs at least two shape IDs.");
        return receiptResult(
          await write("group_shapes", { slideId, shapeIds, ...(name && { name }), ...guardArgs }),
        );
      }
      if (shapeIds.length !== 1) throw new Error("Ungroup takes exactly one group shape ID.");
      return receiptResult(
        await write("ungroup_shape", { slideId, shapeId: shapeIds[0], ...guardArgs }),
      );
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
