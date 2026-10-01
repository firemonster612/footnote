import { guardWrite, textShapeIds } from "./guard.ts";
import { batchFailureWarning, describeError, requireApi, shapeNotFound } from "./presentation.ts";
import { type PendingShapeLists, queueShapeLists } from "./shapeReader.ts";
import { applyFont, readBack, shapeReceipt, withSnapshots } from "./shapes.ts";
import type { EditTableArgs, WriteGuardArgs } from "./types.ts";

function createTable(
  shapes: PowerPoint.ShapeCollection,
  create: NonNullable<EditTableArgs["create"]>,
): PowerPoint.Shape {
  const { rows, columns, name, style, ...options } = create;
  const shape = shapes
    // Office validates the style name and fails the sync with InvalidArgument for unknown styles.
    .addTable(rows, columns, {
      ...options,
      ...(style && { style: style as PowerPoint.TableStyle }),
    })
    .load("id");
  if (name) shape.name = name;
  return shape;
}

/** Queues structure edits, then cell edits: in one batch they still run in that order, so cells address the final grid. */
function queueTableEdits(table: PowerPoint.Table, args: EditTableArgs): void {
  if (args.deleteRows)
    table.rows.deleteRows(args.deleteRows.map((index) => table.rows.getItemAt(index)));
  if (args.deleteColumns)
    table.columns.deleteColumns(args.deleteColumns.map((index) => table.columns.getItemAt(index)));
  if (args.insertRows) table.rows.add(args.insertRows.index, args.insertRows.count);
  if (args.insertColumns) table.columns.add(args.insertColumns.index, args.insertColumns.count);

  for (const { row, column, text, font, fill, align } of args.cells ?? []) {
    const cell = table.getCellOrNullObject(row, column);
    if (text !== undefined) cell.text = text;
    if (font) applyFont(cell.font, font);
    if (fill) cell.fill.setSolidColor(fill);
    if (align) cell.horizontalAlignment = align;
  }
  for (const { row, column, rowCount, columnCount } of args.merge ?? [])
    table.mergeCells(row, column, rowCount, columnCount);
  args.columnWidths?.forEach((width, index) => {
    table.columns.getItemAt(index).width = width;
  });
}

export const tableOps = {
  edit_table: ({
    expectedFingerprints,
    snapshotSlideIds,
    ...args
  }: EditTableArgs & WriteGuardArgs) =>
    PowerPoint.run(async (context) => {
      requireApi("1.8", "Tables");
      const edits =
        args.insertRows ||
        args.insertColumns ||
        args.deleteRows ||
        args.deleteColumns ||
        args.cells ||
        args.merge ||
        args.columnWidths;
      if (edits) requireApi("1.9", "Editing table structure and cells");

      const guard = guardWrite(context, [args.slideId], { expectedFingerprints, snapshotSlideIds });
      const shapes = context.presentation.slides.getItem(args.slideId).shapes.load("items/id");
      const existing =
        args.create || args.shapeId === undefined
          ? undefined
          : shapes.getItemOrNullObject(args.shapeId).load("id");
      const checked = await guard.check();
      if (existing?.isNullObject) throw shapeNotFound(args.shapeId ?? "", args.slideId);
      const before = new Set(shapes.items.map((shape) => shape.id));

      const shape = args.create ? createTable(shapes, args.create) : existing;
      if (!shape) throw new Error("Pass shapeId of an existing table, or create.");
      queueTableEdits(shape.getTable(), args);
      let pending: PendingShapeLists | undefined = queueShapeLists(context, [shapes], "styles", [
        textShapeIds(checked, args.slideId),
      ]);
      const warnings: string[] = [];
      let shapeId: string;
      try {
        await context.sync();
        shapeId = shape.id;
      } catch (error) {
        pending = undefined;
        if (!args.create) {
          shapeId = args.shapeId ?? "";
          warnings.push(batchFailureWarning(error));
        } else {
          // Creation and edits share one batch; a failed edit leaves the new table in place.
          const after = shapes.load("items/id");
          await context.sync();
          const created = after.items.find((item) => !before.has(item.id));
          if (!created) throw error;
          shapeId = created.id;
          warnings.push(
            `Table ${shapeId} was created but editing it failed: ${describeError(error)}`,
          );
        }
      }
      const [read = []] = await readBack(context, [shapes], pending);
      return withSnapshots(shapeReceipt(args.slideId, read, [shapeId], warnings), checked);
    }),
};
