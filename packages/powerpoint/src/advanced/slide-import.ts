// Reduces a PPTX to a subset of its slides so `insertSlidesFromBase64` imports only those.

import type JSZip from "jszip";
import {
  PRESENTATION_PART,
  findRelatedPart,
  listSlideParts,
  loadPptx,
  readPart,
  readRelationships,
  relTypes,
  relsPath,
  removeContentTypeOverride,
  removeRelationship,
  resolveTarget,
  savePptx,
} from "./ooxml.ts";

/**
 * Keeps the 1-based `slideNumbers` (in source order) and drops every other slide with its notes, its section entry,
 * and the links kept slides had to it.
 */
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
  const dropped = slides.filter((_, i) => !keep.has(i + 1));
  let presentation = await readPart(zip, PRESENTATION_PART);
  for (const slide of dropped) {
    presentation = presentation
      .replace(new RegExp(`<p:sldId\\b[^>]*\\bid="${slide.sldId}"[^>]*/>`), "")
      .replace(new RegExp(`<p14:sldId\\b[^>]*\\bid="${slide.sldId}"[^>]*/>`), "");
    await removeRelationship(zip, PRESENTATION_PART, slide.relId);
    const notesPath = await findRelatedPart(zip, slide.path, relTypes.notesSlide);
    for (const part of notesPath ? [slide.path, notesPath] : [slide.path]) {
      zip.remove(part);
      zip.remove(relsPath(part));
      await removeContentTypeOverride(zip, part);
    }
  }
  zip.file(PRESENTATION_PART, presentation);
  const droppedPaths = new Set(dropped.map((slide) => slide.path));
  for (const [i, slide] of slides.entries())
    if (keep.has(i + 1)) await dropLinksTo(zip, slide.path, droppedPaths);
  return savePptx(zip);
}

/** Removes a slide's relationships to dropped slides and the hyperlinks that use them, which would otherwise dangle. */
async function dropLinksTo(
  zip: JSZip,
  slidePath: string,
  droppedPaths: Set<string>,
): Promise<void> {
  const dangling = (await readRelationships(zip, slidePath)).filter(
    (rel) => rel.type === relTypes.slide && droppedPaths.has(resolveTarget(slidePath, rel.target)),
  );
  if (dangling.length === 0) return;
  let xml = await readPart(zip, slidePath);
  for (const rel of dangling) {
    await removeRelationship(zip, slidePath, rel.id);
    xml = xml.replace(
      new RegExp(
        `<a:(hlinkClick|hlinkHover)\\b[^>]*\\br:id="${rel.id}"[^>]*?(?:/>|>[\\s\\S]*?</a:\\1>)`,
        "g",
      ),
      "",
    );
  }
  zip.file(slidePath, xml);
}
