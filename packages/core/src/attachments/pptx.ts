import JSZip from "jszip";
import type { ExtractedContent } from "./extracted-content.ts";

// OOXML here is machine-written and flat enough for targeted patterns: <p:sp> never nests inside
// another <p:sp>, and <a:t> holds only escaped text. A DOM parser would need a dependency in Node.
const RELATIONSHIP = /<Relationship\b[^>]*>/g;
const SLIDE_ID = /<p:sldId\b[^>]*\br:id="([^"]+)"/g;
const SHAPE = /<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g;
const PARAGRAPH = /<a:p>([\s\S]*?)<\/a:p>/g;
const TEXT_RUN_OR_BREAK = /<a:t>([^<]*)<\/a:t>|<a:br\b[^>]*\/>/g;
const TITLE_PLACEHOLDER = /<p:ph\b[^>]*type="(?:title|ctrTitle)"/;
const NOTES_BODY_PLACEHOLDER = /<p:ph\b[^>]*type="body"/;
const XML_ENTITY = /&(?:#x([0-9a-fA-F]+)|#(\d+)|(amp|lt|gt|quot|apos));/g;
const NAMED_ENTITIES: Record<string, string> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };

/** Slide titles, text, and speaker notes per slide, in presentation order. */
export async function processPptx(bytes: ArrayBuffer): Promise<ExtractedContent> {
  const zip = await JSZip.loadAsync(bytes);
  const presentation = await readPart(zip, "ppt/presentation.xml");
  const presentationRels = await readRelationships(zip, "ppt/presentation.xml");
  const slidePaths = [...presentation.matchAll(SLIDE_ID)].flatMap(([, relId]) => {
    const target = relId ? presentationRels.get(relId) : undefined;
    return target ? [target] : [];
  });

  const slides = await Promise.all(
    slidePaths.map((path, index) => readSlide(zip, path, index + 1)),
  );
  return {
    text: slides.join("\n\n"),
    summary: slidePaths.length === 1 ? "1 slide" : `${slidePaths.length} slides`,
  };
}

async function readSlide(zip: JSZip, path: string, slideNumber: number): Promise<string> {
  const xml = await readPart(zip, path);
  const shapes = xml.match(SHAPE) ?? [];
  const titleShape = shapes.find((shape) => TITLE_PLACEHOLDER.test(shape));
  const title = titleShape ? paragraphs(titleShape).join(" ") : "";
  const body = paragraphs(titleShape ? xml.replace(titleShape, "") : xml);

  const notesPath = [...(await readRelationships(zip, path)).values()].find((target) =>
    target.includes("/notesSlides/"),
  );
  const notesXml = notesPath ? await readPart(zip, notesPath) : "";
  const notesShape = (notesXml.match(SHAPE) ?? []).find((shape) =>
    NOTES_BODY_PLACEHOLDER.test(shape),
  );
  const notes = notesShape ? paragraphs(notesShape) : [];

  return [
    title ? `## Slide ${slideNumber}: ${title}` : `## Slide ${slideNumber}`,
    ...body,
    ...(notes.length > 0 ? ["", "Notes:", ...notes] : []),
  ].join("\n");
}

function paragraphs(xml: string): string[] {
  return [...xml.matchAll(PARAGRAPH)]
    .map(([, inner = ""]) =>
      [...inner.matchAll(TEXT_RUN_OR_BREAK)]
        .map(([, text]) => (text === undefined ? "\n" : decodeXml(text)))
        .join(""),
    )
    .filter((text) => text.trim().length > 0);
}

/** Relationship targets of a part, by ID, resolved to zip paths. */
async function readRelationships(zip: JSZip, partPath: string): Promise<Map<string, string>> {
  const directory = partPath.slice(0, partPath.lastIndexOf("/") + 1);
  const relsPath = `${directory}_rels/${partPath.slice(directory.length)}.rels`;
  const xml = zip.file(relsPath) ? await readPart(zip, relsPath) : "";
  const targets = new Map<string, string>();
  for (const [element] of xml.matchAll(RELATIONSHIP)) {
    const id = /\bId="([^"]+)"/.exec(element)?.[1];
    const target = /\bTarget="([^"]+)"/.exec(element)?.[1];
    if (id && target) targets.set(id, new URL(target, `file:///${directory}`).pathname.slice(1));
  }
  return targets;
}

async function readPart(zip: JSZip, path: string): Promise<string> {
  const file = zip.file(path);
  if (!file) throw new Error(`Not a valid PowerPoint file: missing ${path}.`);
  return file.async("string");
}

function decodeXml(text: string): string {
  return text.replace(XML_ENTITY, (entity, hex?: string, decimal?: string, named?: string) => {
    if (hex) return String.fromCodePoint(Number.parseInt(hex, 16));
    if (decimal) return String.fromCodePoint(Number(decimal));
    return NAMED_ENTITIES[named ?? ""] ?? entity;
  });
}
