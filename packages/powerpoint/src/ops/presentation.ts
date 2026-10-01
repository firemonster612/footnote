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

export async function findSlide(
  context: PowerPoint.RequestContext,
  slideId: string,
): Promise<{ slide: PowerPoint.Slide; index: number; slides: PowerPoint.Slide[] }> {
  const slides = await loadSlides(context);
  const index = slides.findIndex((slide) => slide.id === slideId);
  const slide = slides[index];
  if (!slide)
    throw new Error(`Slide ${slideId} not found. Use slide IDs from <deck_state> or get_deck.`);
  return { slide, index, slides };
}

export async function slideIds(context: PowerPoint.RequestContext): Promise<string[]> {
  return (await loadSlides(context)).map((slide) => slide.id);
}

export async function findShape(
  slide: PowerPoint.Slide,
  shapeId: string,
): Promise<PowerPoint.Shape> {
  const shape = slide.shapes.getItemOrNullObject(shapeId);
  shape.load("id");
  await slide.context.sync();
  if (shape.isNullObject)
    throw new Error(
      `Shape ${shapeId} not found on slide ${slide.id}. Call get_slide for current shape IDs.`,
    );
  return shape;
}

/**
 * Read-back after a committed write must never fail the op: the model would redo a write that already
 * happened (duplicate slides). Report the failed verification as a warning instead.
 */
export function readBackFailureWarning(error: unknown): string {
  const detail =
    error instanceof OfficeExtension.Error
      ? describeOfficeError({
          message: error.message,
          code: error.code,
          debugInfo: error.debugInfo,
        })
      : error instanceof Error
        ? error.message
        : String(error);
  return `The write was applied, but reading it back failed (${detail}). Call get_slide before editing the affected slides again; do not repeat the write.`;
}

/**
 * PowerPoint for the web crashes its editor ("Sorry, we ran into a problem") when the API deletes the slide that's
 * open on screen. Select a slide that's staying before deleting.
 */
export async function moveSelectionOff(
  context: PowerPoint.RequestContext,
  deleting: string[],
  preferredId?: string,
): Promise<void> {
  if (deleting.length === 0 || !supportsApi("1.5")) return;
  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();
  const staying = slides.items.map((slide) => slide.id).filter((id) => !deleting.includes(id));
  const target = preferredId && staying.includes(preferredId) ? preferredId : staying[0];
  if (!target) return;
  context.presentation.setSelectedSlides([target]);
  await context.sync();
}
