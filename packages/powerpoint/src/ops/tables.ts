import { findShape, findSlide, requireApi } from "./presentation.ts";
import { applyFont, shapeReceipt } from "./shapes.ts";
import type { EditTableArgs } from "./types.ts";

async function createTable(
  slide: PowerPoint.Slide,
  create: NonNullable<EditTableArgs["create"]>,
): Promise<PowerPoint.Shape> {
  const { rows, columns, name, style, ...options } = create;
  const shape = slide.shapes
    // Office validates the style name and fails the sync with InvalidArgument for unknown styles.
    .addTable(rows, columns, {
      ...options,
      ...(style && { style: style as PowerPoint.TableStyle }),
    })
    .load("id");
  if (name) shape.name = name;
  await slide.context.sync();
  return shape;
}

export const tableOps = {
  edit_table: (args: EditTableArgs) =>
    PowerPoint.run(async (context) => {
      requireApi("1.8", "Tables");
      const { slide } = await findSlide(context, args.slideId);
      if (!args.create && !args.shapeId)
        throw new Error("Pass shapeId of an existing table, or create.");
      const shape = args.create
        ? await createTable(slide, args.create)
        : await findShape(slide, args.shapeId ?? "");
      const table = shape.getTable();
      const edits =
        args.insertRows ||
        args.insertColumns ||
        args.deleteRows ||
        args.deleteColumns ||
        args.cells ||
        args.merge ||
        args.columnWidths;
      if (edits) requireApi("1.9", "Editing table structure and cells");

      // Structure first, so cell coordinates refer to the final grid.
      if (args.deleteRows)
        table.rows.deleteRows(args.deleteRows.map((index) => table.rows.getItemAt(index)));
      if (args.deleteColumns)
        table.columns.deleteColumns(
          args.deleteColumns.map((index) => table.columns.getItemAt(index)),
        );
      if (args.insertRows) table.rows.add(args.insertRows.index, args.insertRows.count);
      if (args.insertColumns) table.columns.add(args.insertColumns.index, args.insertColumns.count);
      await context.sync();

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
      await context.sync();

      return shapeReceipt(context, slide, [shape.id], []);
    }),
};
