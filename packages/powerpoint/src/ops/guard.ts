// Read-before-write and undo snapshots, done inside a write op's own PowerPoint.run so a guarded write is one host call.
import { fingerprintShapes } from "./fingerprint.ts";
import { slideIds as loadSlideIds, slideNotFound, supportsApi } from "./presentation.ts";
import { queueShapeLists } from "./shapeReader.ts";
import type { ShapeInfo, SlideSnapshot, WriteGuardArgs } from "./types.ts";

export const staleSlideMessage = (position: number): string =>
  `Slide ${position} changed since you last read it; call get_slide again.`;

export interface CheckedWrite {
  /** Every slide ID in deck order, before the write. */
  slideIds: string[];
  /** Summary shape trees, before the write, of the slides checked against an expected fingerprint. */
  shapes: Map<string, ShapeInfo[]>;
  snapshots: Record<string, SlideSnapshot>;
  /** Start of the write's warnings: says when undo snapshots couldn't be taken. */
  warnings: string[];
}

export const NO_SNAPSHOT_WARNING =
  "This PowerPoint lacks PowerPointApi 1.8, so the change was made without an undo checkpoint: Undo and Revert can't restore these slides.";

/**
 * Queues what a guarded write needs before it changes anything: the slide list, shape trees of the slides the model
 * has read (for the freshness check) and undo exports. The op queues its own lookups next, then awaits `check()`:
 * one sync carries all of it, and only the freshness check adds read-only syncs. Throws the model-facing
 * "not found" or "changed since you last read it" error before any write is queued.
 */
export function guardWrite(
  context: PowerPoint.RequestContext,
  targetSlideIds: string[],
  { expectedFingerprints = {}, snapshotSlideIds = [] }: WriteGuardArgs,
): { check(): Promise<CheckedWrite> } {
  // Exporting needs 1.8. Without it the write still goes ahead; blocking every edit for the sake of undo would be worse.
  const canSnapshot = supportsApi("1.8");
  const warnings = snapshotSlideIds.length > 0 && !canSnapshot ? [NO_SNAPSHOT_WARNING] : [];
  const slides = context.presentation.slides.load("items/id");
  const checkedIds = targetSlideIds.filter((id) => expectedFingerprints[id] !== undefined);
  const trees = queueShapeLists(
    context,
    checkedIds.map((id) => slides.getItem(id).shapes),
    "summary",
  );
  const exports = (canSnapshot ? snapshotSlideIds : []).map((id) => ({
    id,
    result: slides.getItem(id).exportAsBase64(),
  }));

  return {
    async check() {
      try {
        await context.sync();
      } catch (error) {
        // A missing slide fails the batch at its getItem; report it the way the model can act on. A failed sync
        // leaves its context unusable, so the slide list comes from a new run.
        const existing = await PowerPoint.run(loadSlideIds);
        const missing = targetSlideIds.find((id) => !existing.includes(id));
        throw missing ? slideNotFound(missing) : error;
      }
      const ids = slides.items.map((slide) => slide.id);
      const missing = targetSlideIds.find((id) => !ids.includes(id));
      if (missing) throw slideNotFound(missing);

      const lists = await trees.read();
      const shapes = new Map(checkedIds.map((id, i) => [id, lists[i] ?? []]));
      for (const [id, current] of shapes) {
        if (fingerprintShapes(current) !== expectedFingerprints[id])
          throw new Error(staleSlideMessage(ids.indexOf(id) + 1));
      }
      const snapshots = Object.fromEntries(
        exports.map(({ id, result }) => [id, { index: ids.indexOf(id), base64: result.value }]),
      );
      return { slideIds: ids, shapes, snapshots, warnings };
    },
  };
}

/** IDs of shapes on `slideId` that the check read with text, so a read-back can load their text without probing frames. */
export function textShapeIds(checked: CheckedWrite, slideId: string): Set<string> {
  const ids = new Set<string>();
  const visit = (shapes: ShapeInfo[]) => {
    for (const shape of shapes) {
      if (shape.text !== undefined) ids.add(shape.id);
      visit(shape.children ?? []);
    }
  };
  visit(checked.shapes.get(slideId) ?? []);
  return ids;
}
