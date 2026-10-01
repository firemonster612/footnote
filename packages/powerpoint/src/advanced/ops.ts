/// <reference types="office-js" />
// Office-realm ops for the advanced tools. Keep this file free of agent-side imports (PptxGenJS, JSZip, typebox).

import type { OpRegistry } from "@footnote/core/contracts";

export interface ReplaceSlideArgs {
  slideId: string;
  /** Single-slide PPTX (an edited export of `slideId`). */
  base64: string;
}

export interface ReplaceSlideResult {
  slideId: string;
  shapes: { id: string; name: string; type: string }[];
}

/**
 * Swaps a slide for an edited copy at the same position: inserts the copy right after it, then deletes the original.
 * The copy uses the destination theme because it was exported from this deck.
 */
async function replaceSlide({ slideId, base64 }: ReplaceSlideArgs): Promise<ReplaceSlideResult> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    const before = new Set(slides.items.map((s) => s.id));
    if (!before.has(slideId)) throw new Error(`Slide ${slideId} not found`);

    context.presentation.insertSlidesFromBase64(base64, {
      targetSlideId: slideId,
      formatting: "UseDestinationTheme",
    });
    slides.load("items/id");
    await context.sync();
    const created = slides.items.filter((s) => !before.has(s.id));
    const replacement = created[0];
    if (created.length !== 1 || !replacement)
      throw new Error(`Expected 1 inserted slide, got ${created.length}`);

    // Deleting the on-screen slide crashes PowerPoint for the web; select the replacement first.
    if (Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
      context.presentation.setSelectedSlides([replacement.id]);
      await context.sync();
    }
    slides.getItem(slideId).delete();
    const shapes = replacement.shapes;
    shapes.load("items/id,items/name,items/type");
    await context.sync();
    return {
      slideId: replacement.id,
      shapes: shapes.items.map((s) => ({ id: s.id, name: s.name, type: s.type })),
    };
  });
}

export const advancedOps: OpRegistry = {
  replace_slide: replaceSlide,
};
