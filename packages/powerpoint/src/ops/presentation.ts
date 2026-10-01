// Office-realm lookups shared by the ops.
import { describeOfficeError } from "@footnote/core/contracts";

export const supportsApi = (version: string): boolean =>
  Office.context.requirements.isSetSupported("PowerPointApi", version);

export function requireApi(version: string, feature: string): void {
  if (!supportsApi(version)) {
    throw new Error(
      `${feature} needs PowerPointApi ${version}, which this PowerPoint doesn't support. Try execute_office_js or ask the user to update Office.`,
    );
  }
}

/** All slides in deck order with id and layout name loaded. */
export async function loadSlides(context: PowerPoint.RequestContext): Promise<PowerPoint.Slide[]> {
  const slides = context.presentation.slides;
  slides.load("items/id,items/layout/name");
  await context.sync();
  return slides.items;
}

export const slideNotFound = (slideId: string): Error =>
  new Error(`Slide ${slideId} not found. Use slide IDs from <deck_state> or get_deck.`);

export async function findSlide(
  context: PowerPoint.RequestContext,
  slideId: string,
): Promise<{ slide: PowerPoint.Slide; index: number; slides: PowerPoint.Slide[] }> {
  const slides = await loadSlides(context);
  const index = slides.findIndex((slide) => slide.id === slideId);
  const slide = slides[index];
  if (!slide) throw slideNotFound(slideId);
  return { slide, index, slides };
}

export async function slideIds(context: PowerPoint.RequestContext): Promise<string[]> {
  return (await loadSlides(context)).map((slide) => slide.id);
}

export const shapeNotFound = (shapeId: string, slideId: string): Error =>
  new Error(
    `Shape ${shapeId} not found on slide ${slideId}. Call get_slide for current shape IDs.`,
  );

/** Office errors with their code and failing statement; other errors by message. */
export function describeError(error: unknown): string {
  if (error instanceof OfficeExtension.Error)
    return describeOfficeError({
      message: error.message,
      code: error.code,
      debugInfo: error.debugInfo,
    });
  return error instanceof Error ? error.message : String(error);
}

/**
 * Read-back after a committed write must never fail the op: the model would redo a write that already
 * happened (duplicate slides). Report the failed verification as a warning instead.
 */
export function readBackFailureWarning(error: unknown): string {
  return `The write was applied, but reading it back failed (${describeError(error)}). Call get_slide before editing the affected slides again; do not repeat the write.`;
}

/** A write's batch failed part-way. Office batches aren't transactional, so edits queued before the failure stay. */
export function batchFailureWarning(error: unknown): string {
  return `PowerPoint rejected part of this write (${describeError(error)}); the edits before the failing one may have applied. The read-back shows the current state; fix what's missing rather than repeating the whole write.`;
}

/**
 * PowerPoint for the web crashes its editor ("Sorry, we ran into a problem") when the API deletes the slide that's
 * open on screen. Select a slide that's staying, in its own sync, before deleting.
 * `slideIds` is the deck as it is now, in order.
 */
export async function moveSelectionOff(
  context: PowerPoint.RequestContext,
  slideIds: string[],
  deleting: string[],
  preferredId?: string,
): Promise<void> {
  if (deleting.length === 0 || !supportsApi("1.5")) return;
  const staying = slideIds.filter((id) => !deleting.includes(id));
  const target = preferredId && staying.includes(preferredId) ? preferredId : staying[0];
  if (!target) return;
  context.presentation.setSelectedSlides([target]);
  await context.sync();
}
