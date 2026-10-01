// Reduces a PPTX to a subset of its slides so `insertSlidesFromBase64` imports only those.

import {
  PRESENTATION_PART,
  findRelatedPart,
  listSlideParts,
  loadPptx,
  readPart,
  relTypes,
  relsPath,
  removeContentTypeOverride,
  removeRelationship,
  savePptx,
} from "./ooxml.ts";

/** Keeps the 1-based `slideNumbers` (in source order) and drops every other slide with its notes. */
export async function keepSlides(pptxBase64: string, slideNumbers: number[]): Promise<string> {
  const zip = await loadPptx(pptxBase64);
  const slides = await listSlideParts(zip);
  const invalid = slideNumbers.filter((n) => !Number.isInteger(n) || n < 1 || n > slides.length);
  if (invalid.length > 0) {
    throw new Error(
      `Slide numbers ${invalid.join(", ")} are out of range; the file has ${slides.length} slides.`,
    );
  }

  const keep = new Set(slideNumbers);
  let presentation = await readPart(zip, PRESENTATION_PART);
  for (const [i, slide] of slides.entries()) {
    if (keep.has(i + 1)) continue;
    presentation = presentation.replace(
      new RegExp(`<p:sldId\\b[^>]*\\bid="${slide.sldId}"[^>]*/>`),
      "",
    );
    await removeRelationship(zip, PRESENTATION_PART, slide.relId);
    const notesPath = await findRelatedPart(zip, slide.path, relTypes.notesSlide);
    for (const part of notesPath ? [slide.path, notesPath] : [slide.path]) {
      zip.remove(part);
      zip.remove(relsPath(part));
      await removeContentTypeOverride(zip, part);
    }
  }
  zip.file(PRESENTATION_PART, presentation);
  return savePptx(zip);
}
