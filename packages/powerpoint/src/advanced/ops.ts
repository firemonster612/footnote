/// <reference types="office-js" />
// Office-realm ops for the advanced tools. Keep this file free of agent-side imports (PptxGenJS, JSZip, typebox).

import type { OpRegistry } from "@footnote/core/contracts";
import { guardWrite } from "../ops/guard.ts";
import { describeError, moveSelectionOff, readBackFailureWarning } from "../ops/presentation.ts";
import type { WriteGuardArgs } from "../ops/types.ts";

export interface ReplaceSlideArgs extends Pick<WriteGuardArgs, "expectedFingerprints"> {
  slideId: string;
  /** Single-slide PPTX (an edited export of `slideId`). */
  base64: string;
}

export interface ReplaceSlideResult {
  slideId: string;
  shapes: { id: string; name: string; type: string }[];
  warnings: string[];
}

/**
 * Swaps a slide for an edited copy at the same position: inserts the copy right after it, then deletes the original.
 * `expectedFingerprints` holds the slide's fingerprint when it was exported, so a user edit made since then fails the
 * replace instead of being lost. The copy uses the destination theme because it was exported from this deck.
 * Once the copy is in, later failures become warnings: the caller must record the copy either way.
 */
async function replaceSlide({
  slideId,
  base64,
  expectedFingerprints,
}: ReplaceSlideArgs): Promise<ReplaceSlideResult> {
  return PowerPoint.run(async (context) => {
    const checked = await guardWrite(context, [slideId], {
      ...(expectedFingerprints && { expectedFingerprints }),
    }).check();
    const before = new Set(checked.slideIds);

    context.presentation.insertSlidesFromBase64(base64, {
      targetSlideId: slideId,
      formatting: "UseDestinationTheme",
    });
    const slides = context.presentation.slides.load("items/id");
    await context.sync();
    const created = slides.items.filter((s) => !before.has(s.id));
    const replacement = created[0];
    if (created.length !== 1 || !replacement)
      throw new Error(`Expected 1 inserted slide, got ${created.length}`);

    const warnings: string[] = [];
    try {
      await moveSelectionOff(
        context,
        [...checked.slideIds, replacement.id],
        [slideId],
        replacement.id,
      );
      context.presentation.slides.getItem(slideId).delete();
      await context.sync();
    } catch (error) {
      warnings.push(
        `The edited copy ${replacement.id} was inserted, but deleting the original slide ${slideId} failed (${describeError(error)}). Delete the original with delete_slides.`,
      );
      return { slideId: replacement.id, shapes: [], warnings };
    }
    try {
      const shapes = replacement.shapes.load("items/id,items/name,items/type");
      await context.sync();
      const list = shapes.items.map((s) => ({ id: s.id, name: s.name, type: s.type }));
      return { slideId: replacement.id, shapes: list, warnings };
    } catch (error) {
      return { slideId: replacement.id, shapes: [], warnings: [readBackFailureWarning(error)] };
    }
  });
}

export const advancedOps: OpRegistry = {
  replace_slide: replaceSlide,
};
