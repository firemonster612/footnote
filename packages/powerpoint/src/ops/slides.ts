import { guardWrite, textShapeIds } from "./guard.ts";
import {
  batchFailureWarning,
  describeError,
  findSlide,
  moveSelectionOff,
  readBackFailureWarning,
  requireApi,
  slideIds,
} from "./presentation.ts";
import { toOutline } from "./read.ts";
import { type PendingShapeLists, queueShapeLists } from "./shapeReader.ts";
import type {
  AddSlideReceipt,
  InsertFormatting,
  RestoreArgs,
  RestoreResult,
  ShapeInfo,
  SlideOutline,
  WriteGuardArgs,
  WriteReceipt,
} from "./types.ts";

/** Slide-level read-back: the deck's slide count and the affected slides' outlines and shape trees. */
interface SlideRead {
  slideCount: number;
  outlines: SlideOutline[];
  shapes: ShapeInfo[][];
}

interface QueuedSlideRead {
  deck: PowerPoint.SlideCollection;
  targets: PowerPoint.Slide[];
  tree: PendingShapeLists;
}

/** Queues a slide-level read-back of `targets` so it can ride on the write's own sync. */
function queueSlideRead(
  context: PowerPoint.RequestContext,
  targets: PowerPoint.Slide[],
  textIds?: readonly ReadonlySet<string>[],
): QueuedSlideRead {
  const deck = context.presentation.slides.load("items/id,items/layout/name");
  for (const target of targets) target.load("id");
  const tree = queueShapeLists(
    context,
    targets.map((target) => target.shapes),
    "summary",
    textIds,
  );
  return { deck, targets, tree };
}

async function finishSlideRead({ deck, targets, tree }: QueuedSlideRead): Promise<SlideRead> {
  const shapes = await tree.read();
  const all = deck.items;
  const outlines = targets.map(({ id }, i) => {
    const index = all.findIndex((slide) => slide.id === id);
    return toOutline(all[index]!, index, shapes[i] ?? []);
  });
  return { slideCount: all.length, outlines, shapes };
}

const slidesById = (context: PowerPoint.RequestContext, ids: string[]): PowerPoint.Slide[] =>
  ids.map((id) => context.presentation.slides.getItem(id));

/** Reads slides back in a new run: after a failed sync the write's own context is unusable. */
function readSlides(ids: string[]): Promise<SlideRead | { error: unknown }> {
  return PowerPoint.run(async (context) => {
    const queued = queueSlideRead(context, slidesById(context, ids));
    await context.sync();
    return finishSlideRead(queued);
  }).catch((error: unknown) => ({ error }));
}

/**
 * Syncs a slide-level write together with its read-back. A failed write is reported as a warning (part of it may
 * have applied) and read back on its own.
 */
async function commitSlideWrite(
  context: PowerPoint.RequestContext,
  ids: string[],
  warnings: string[],
  textIds?: readonly ReadonlySet<string>[],
): Promise<SlideRead | { error: unknown }> {
  const queued = queueSlideRead(context, slidesById(context, ids), textIds);
  try {
    await context.sync();
  } catch (error) {
    warnings.push(batchFailureWarning(error));
    return readSlides(ids);
  }
  return finishSlideRead(queued).catch((error: unknown) => ({ error }));
}

/** Receipt for slide-level writes: outlines of the affected slides as they are now. */
function slideReceipt(
  changed: string[],
  read: SlideRead | { error: unknown },
  {
    warnings = [],
    snapshots = {},
    ...slideLists
  }: Pick<WriteReceipt, "createdSlideIds" | "deletedSlideIds" | "snapshots"> & {
    warnings?: string[];
  } = {},
): WriteReceipt {
  const listed = { ...slideLists, ...(Object.keys(snapshots).length > 0 && { snapshots }) };
  if ("error" in read)
    return {
      changed,
      verified: {},
      warnings: [...warnings, readBackFailureWarning(read.error)],
      fingerprints: {},
      ...listed,
    };
  return {
    changed,
    verified: {
      slides: read.outlines.map(({ id, index, layout, title, shapeCount }) => ({
        id,
        position: index + 1,
        layout,
        title,
        shapeCount,
      })),
      slideCount: read.slideCount,
    },
    warnings,
    fingerprints: Object.fromEntries(
      read.outlines.map((outline) => [outline.id, outline.fingerprint]),
    ),
    ...listed,
  };
}

/**
 * Runs `insert` and returns the IDs of the slides it added, in deck order. The new slide list loads in the insert's
 * own sync, so a committed insert can't be followed by a failed read.
 */
async function insertedSlides(
  context: PowerPoint.RequestContext,
  insert: () => void,
): Promise<string[]> {
  const before = new Set(await slideIds(context));
  insert();
  const after = context.presentation.slides.load("items/id");
  await context.sync();
  return after.items.map((slide) => slide.id).filter((id) => !before.has(id));
}

async function moveSlides(
  context: PowerPoint.RequestContext,
  ids: string[],
  toIndex: number,
): Promise<void> {
  ids.forEach((id, offset) => context.presentation.slides.getItem(id).moveTo(toIndex + offset));
  await context.sync();
}

/**
 * Moves restored copies to their saved indexes (ascending) while the other slides keep their order. Placing the deck
 * front to back, each slide moves in from a later position, so the slides already placed stay put.
 */
async function moveToSavedIndexes(
  context: PowerPoint.RequestContext,
  copies: { id: string; index: number }[],
): Promise<void> {
  const current = await slideIds(context);
  const copyIds = new Set(copies.map((copy) => copy.id));
  const order = current.filter((id) => !copyIds.has(id));
  for (const { id, index } of copies) order.splice(Math.min(index, order.length), 0, id);
  const deck = [...current];
  let moved = false;
  order.forEach((id, target) => {
    const from = deck.indexOf(id);
    if (from === target) return;
    deck.splice(from, 1);
    deck.splice(target, 0, id);
    context.presentation.slides.getItem(id).moveTo(target);
    moved = true;
  });
  if (moved) await context.sync();
}

/** add_slide's receipt: the slide outline plus the new slide's placeholders. */
function addSlideReceipt(
  created: string[],
  read: SlideRead | { error: unknown },
  warnings: string[],
): AddSlideReceipt {
  const receipt = slideReceipt(created, read, { createdSlideIds: created, warnings });
  const shapes = "error" in read ? [] : (read.shapes[0] ?? []);
  return {
    ...receipt,
    shapes: shapes.map(({ id, name, placeholder, left, top, width, height }) => ({
      id,
      name,
      ...(placeholder && { placeholder }),
      left,
      top,
      width,
      height,
    })),
  };
}

/**
 * add_slide's batch failed after add() may have run: add() always appends, so a slide past `before` is the new one
 * (it's the move or the read-back that failed). Rethrows when nothing was added.
 */
async function recoverAddedSlide(before: string[], error: unknown): Promise<AddSlideReceipt> {
  let created: string[];
  try {
    created = (await PowerPoint.run(slideIds)).filter((id) => !before.includes(id));
  } catch (readError) {
    return {
      changed: [],
      verified: {},
      warnings: [
        `PowerPoint reported an error while adding the slide (${describeError(error)}) and the deck couldn't be read afterwards (${describeError(readError)}), so the slide may have been added. Call get_deck before adding it again.`,
      ],
      fingerprints: {},
      shapes: [],
    };
  }
  if (created.length === 0) throw error;
  return addSlideReceipt(created, await readSlides(created), [
    `The slide was added at the end, but moving it failed: ${describeError(error)}`,
  ]);
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
    PowerPoint.run(async (context): Promise<AddSlideReceipt> => {
      if (index !== undefined) requireApi("1.8", "Positioning a new slide");
      const slides = context.presentation.slides;
      const before = await slideIds(context);
      slides.add({
        ...(layoutId && { layoutId }),
        ...(slideMasterId && { slideMasterId }),
      });
      // add() always appends, so the slide at the old count is the new one. Moving it and reading it back in the
      // same batch makes it appear at its position at once.
      const added = slides.getItemAt(before.length);
      if (index !== undefined) added.moveTo(index);
      const queued = queueSlideRead(context, [added]);
      try {
        await context.sync();
      } catch (error) {
        return recoverAddedSlide(before, error);
      }
      const read = await finishSlideRead(queued).catch((error: unknown) => ({ error }));
      return addSlideReceipt([added.id], read, []);
    }),

  delete_slides: ({ slideIds: ids, ...guardArgs }: { slideIds: string[] } & WriteGuardArgs) =>
    PowerPoint.run(async (context) => {
      const checked = await guardWrite(context, ids, guardArgs).check();
      await moveSelectionOff(context, checked.slideIds, ids);
      for (const id of ids) context.presentation.slides.getItem(id).delete();
      const slides = context.presentation.slides.load("items/id");
      const warnings = [...checked.warnings];
      const remaining = (count: number): SlideRead => ({
        slideCount: count,
        outlines: [],
        shapes: [],
      });
      let read: SlideRead | { error: unknown };
      try {
        await context.sync();
        read = remaining(slides.items.length);
      } catch (error) {
        warnings.push(batchFailureWarning(error));
        read = await PowerPoint.run(slideIds).then(
          (left) => remaining(left.length),
          (readError: unknown) => ({ error: readError }),
        );
      }
      return slideReceipt(ids, read, {
        deletedSlideIds: ids,
        snapshots: checked.snapshots,
        warnings,
      });
    }),

  move_slide: ({
    slideId,
    toIndex,
    ...guardArgs
  }: { slideId: string; toIndex: number } & WriteGuardArgs) =>
    PowerPoint.run(async (context) => {
      requireApi("1.8", "Moving slides");
      const checked = await guardWrite(context, [slideId], guardArgs).check();
      context.presentation.slides.getItem(slideId).moveTo(toIndex);
      const warnings = [...checked.warnings];
      const read = await commitSlideWrite(context, [slideId], warnings, [
        textShapeIds(checked, slideId),
      ]);
      return slideReceipt([slideId], read, { snapshots: checked.snapshots, warnings });
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
      const warnings: string[] = [];
      if (toIndex !== undefined) {
        try {
          await moveSlides(context, created, toIndex);
        } catch (error) {
          warnings.push(
            `The copy was added after the original, but moving it failed: ${describeError(error)}`,
          );
        }
      }
      return slideReceipt(created, await readSlides(created), {
        createdSlideIds: created,
        warnings,
      });
    }),

  apply_layout: ({
    slideId,
    layoutId,
    ...guardArgs
  }: { slideId: string; layoutId: string } & WriteGuardArgs) =>
    PowerPoint.run(async (context) => {
      requireApi("1.8", "Applying layouts");
      const guard = guardWrite(context, [slideId], guardArgs);
      const slide = context.presentation.slides.getItem(slideId);
      const layout = slide.slideMaster.layouts.getItemOrNullObject(layoutId).load("id");
      const checked = await guard.check();
      if (layout.isNullObject)
        throw new Error(
          `Layout ${layoutId} not found on this slide's master. Call get_deck for layout IDs.`,
        );
      slide.applyLayout(layout);
      const warnings = [...checked.warnings];
      // No text hints: a new layout can rework the slide's placeholders, so frames are probed again.
      const read = await commitSlideWrite(context, [slideId], warnings);
      return slideReceipt([slideId], read, { snapshots: checked.snapshots, warnings });
    }),

  insert_slides: (args: InsertSlidesArgs) =>
    PowerPoint.run(async (context) => {
      const created = await insertSlides(context, args);
      return slideReceipt(created, await readSlides(created), {
        createdSlideIds: created,
      });
    }),

  /**
   * Undo: re-inserts the exported slides, deletes `deleteSlideIds`, then moves the copies to their saved indexes.
   * Copies go in before anything is deleted: deleting the on-screen slide first crashes PowerPoint for the web.
   * Stops at the first failure and reports what it did, so the caller keeps the rest of the turn for another try.
   */
  restore_slides: ({ deleteSlideIds, inserts }: RestoreArgs) =>
    PowerPoint.run(async (context): Promise<RestoreResult> => {
      requireApi("1.8", "Undo");
      const existing = await slideIds(context);
      const toDelete = deleteSlideIds.filter((id) => existing.includes(id));
      const sorted = [...inserts].sort((a, b) => a.index - b.index);
      const result: RestoreResult = { removedSlideIds: [], idMap: {} };
      try {
        for (const { slideId, base64, index } of sorted) {
          // A copy goes right after the slide it replaces, or near its old position among the slides that stay.
          const ids = await slideIds(context);
          const staying = ids.filter((id) => !toDelete.includes(id));
          const targetSlideId = ids.includes(slideId)
            ? slideId
            : staying[Math.min(index, staying.length) - 1];
          const [copy] = await insertSlides(context, {
            base64,
            formatting: "KeepSourceFormatting",
            ...(targetSlideId && { targetSlideId }),
          });
          if (copy) result.idMap[slideId] = copy;
        }
        const copies = sorted.flatMap(({ slideId, index }) => {
          const id = result.idMap[slideId];
          return id ? [{ id, index }] : [];
        });
        await moveSelectionOff(context, await slideIds(context), toDelete, copies[0]?.id);
        for (const id of toDelete) context.presentation.slides.getItem(id).delete();
        await context.sync();
        result.removedSlideIds = toDelete;
        await moveToSavedIndexes(context, copies);
      } catch (error) {
        return { ...result, error: describeError(error) };
      }
      return result;
    }),
};
