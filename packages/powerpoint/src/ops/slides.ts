import { findSlide, loadSlides, readBackFailureWarning, requireApi, slideIds } from "./presentation.ts";
import { readOutlines } from "./read.ts";
import type { InsertFormatting, WriteReceipt } from "./types.ts";

/** Receipt for slide-level writes: outlines of the affected slides as they are now. */
async function slideReceipt(
  context: PowerPoint.RequestContext,
  ids: string[],
  extra: Pick<WriteReceipt, "createdSlideIds" | "deletedSlideIds"> & { warnings?: string[] } = {},
): Promise<WriteReceipt> {
  let all: PowerPoint.Slide[];
  let outlines: Awaited<ReturnType<typeof readOutlines>>["outlines"];
  try {
    all = await loadSlides(context);
    const slides = all.filter((slide) => ids.includes(slide.id));
    ({ outlines } = await readOutlines(context, slides, (slide) => all.indexOf(slide)));
  } catch (error) {
    return {
      changed: ids,
      verified: {},
      warnings: [...(extra.warnings ?? []), readBackFailureWarning(error)],
      fingerprints: {},
      ...(extra.createdSlideIds && { createdSlideIds: extra.createdSlideIds }),
      ...(extra.deletedSlideIds && { deletedSlideIds: extra.deletedSlideIds }),
    };
  }
  const missing = ids.filter(
    (id) => !outlines.some((outline) => outline.id === id) && !extra.deletedSlideIds?.includes(id),
  );
  return {
    changed: ids,
    verified: {
      slides: outlines.map(({ id, index, layout, title, shapeCount }) => ({
        id,
        position: index + 1,
        layout,
        title,
        shapeCount,
      })),
      slideCount: all.length,
    },
    warnings: [
      ...(extra.warnings ?? []),
      ...missing.map((id) => `Slide ${id} not found after the write.`),
    ],
    fingerprints: Object.fromEntries(outlines.map((outline) => [outline.id, outline.fingerprint])),
    ...(extra.createdSlideIds && { createdSlideIds: extra.createdSlideIds }),
    ...(extra.deletedSlideIds && { deletedSlideIds: extra.deletedSlideIds }),
  };
}

/** Runs `insert` and returns the IDs of the slides it added, in deck order. */
async function insertedSlides(
  context: PowerPoint.RequestContext,
  insert: () => void,
): Promise<string[]> {
  const before = new Set(await slideIds(context));
  insert();
  await context.sync();
  return (await slideIds(context)).filter((id) => !before.has(id));
}

async function moveSlides(
  context: PowerPoint.RequestContext,
  ids: string[],
  toIndex: number,
): Promise<void> {
  ids.forEach((id, offset) => context.presentation.slides.getItem(id).moveTo(toIndex + offset));
  await context.sync();
}

interface InsertSlidesArgs {
  base64: string;
  /** Insert after this slide; omitted inserts at the beginning. */
  targetSlideId?: string;
  formatting?: InsertFormatting;
  sourceSlideIds?: string[];
}

async function insertSlides(
  context: PowerPoint.RequestContext,
  args: InsertSlidesArgs,
): Promise<string[]> {
  const { base64, ...options } = args;
  return insertedSlides(context, () =>
    context.presentation.insertSlidesFromBase64(base64, options),
  );
}

export const slideOps = {
  add_slide: ({
    layoutId,
    slideMasterId,
    index,
  }: {
    layoutId?: string;
    slideMasterId?: string;
    index?: number;
  }) =>
    PowerPoint.run(async (context) => {
      const created = await insertedSlides(context, () =>
        context.presentation.slides.add({
          ...(layoutId && { layoutId }),
          ...(slideMasterId && { slideMasterId }),
        }),
      );
      if (index !== undefined) {
        requireApi("1.8", "Positioning a new slide");
        await moveSlides(context, created, index);
      }
      return slideReceipt(context, created, { createdSlideIds: created });
    }),

  delete_slides: ({ slideIds: ids }: { slideIds: string[] }) =>
    PowerPoint.run(async (context) => {
      const existing = new Set(await slideIds(context));
      const missing = ids.filter((id) => !existing.has(id));
      if (missing.length > 0)
        throw new Error(
          `Slides not found: ${missing.join(", ")}. Use slide IDs from <deck_state> or get_deck.`,
        );
      for (const id of ids) context.presentation.slides.getItem(id).delete();
      await context.sync();
      return slideReceipt(context, ids, { deletedSlideIds: ids });
    }),

  move_slide: ({ slideId, toIndex }: { slideId: string; toIndex: number }) =>
    PowerPoint.run(async (context) => {
      requireApi("1.8", "Moving slides");
      await findSlide(context, slideId);
      await moveSlides(context, [slideId], toIndex);
      return slideReceipt(context, [slideId]);
    }),

  duplicate_slide: ({ slideId, toIndex }: { slideId: string; toIndex?: number }) =>
    PowerPoint.run(async (context) => {
      requireApi("1.8", "Duplicating slides");
      const { slide } = await findSlide(context, slideId);
      const exported = slide.exportAsBase64();
      await context.sync();
      const created = await insertSlides(context, {
        base64: exported.value,
        targetSlideId: slideId,
        formatting: "KeepSourceFormatting",
      });
      if (toIndex !== undefined) await moveSlides(context, created, toIndex);
      return slideReceipt(context, created, { createdSlideIds: created });
    }),

  apply_layout: ({ slideId, layoutId }: { slideId: string; layoutId: string }) =>
    PowerPoint.run(async (context) => {
      requireApi("1.8", "Applying layouts");
      const { slide } = await findSlide(context, slideId);
      const layout = slide.slideMaster.layouts.getItemOrNullObject(layoutId).load("id");
      await context.sync();
      if (layout.isNullObject)
        throw new Error(
          `Layout ${layoutId} not found on this slide's master. Call get_deck for layout IDs.`,
        );
      slide.applyLayout(layout);
      await context.sync();
      return slideReceipt(context, [slideId]);
    }),

  insert_slides: (args: InsertSlidesArgs) =>
    PowerPoint.run(async (context) => {
      const created = await insertSlides(context, args);
      return slideReceipt(context, created, { createdSlideIds: created });
    }),

  /** Undo: deletes slides, then reinserts exported slides at their original indexes (ascending). */
  restore_slides: ({
    deleteSlideIds,
    inserts,
  }: {
    deleteSlideIds: string[];
    inserts: { base64: string; index: number }[];
  }) =>
    PowerPoint.run(async (context) => {
      requireApi("1.8", "Undo");
      const existing = new Set(await slideIds(context));
      const toDelete = deleteSlideIds.filter((id) => existing.has(id));
      for (const id of toDelete) context.presentation.slides.getItem(id).delete();
      await context.sync();

      const restoredIds: string[] = [];
      for (const { base64, index } of [...inserts].sort((a, b) => a.index - b.index)) {
        const ids = await slideIds(context);
        const targetSlideId = ids[Math.min(index, ids.length) - 1];
        const created = await insertSlides(context, {
          base64,
          formatting: "KeepSourceFormatting",
          ...(targetSlideId && { targetSlideId }),
        });
        restoredIds.push(...created);
      }
      return { removedSlideIds: toDelete, restoredSlideIds: restoredIds };
    }),
};
