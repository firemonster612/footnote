// Minimal OPC/PresentationML plumbing for editing PPTX packages with JSZip.
// String-level edits on known part shapes; no DOM, so it runs in the side panel and in tests alike.

import JSZip from "jszip";

const EMU_PER_POINT = 12700;
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

export const relTypes = {
  chart: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
  notesMaster: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
  notesSlide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
  package: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
  slide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
  theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
} as const;

export const contentTypes = {
  chart: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
  notesMaster: "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml",
  notesSlide: "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
} as const;

export const PRESENTATION_PART = "ppt/presentation.xml";

export interface Relationship {
  id: string;
  type: string;
  target: string;
}

export interface SlidePart {
  /** `<p:sldId id>` in presentation.xml. */
  sldId: string;
  /** Relationship ID from presentation.xml to the slide part. */
  relId: string;
  path: string;
}

export interface SlideSize {
  /** Points. */
  width: number;
  height: number;
}

export function loadPptx(base64: string): Promise<JSZip> {
  return JSZip.loadAsync(base64, { base64: true });
}

export function savePptx(zip: JSZip): Promise<string> {
  return zip.generateAsync({ type: "base64", compression: "DEFLATE" });
}

export async function readPart(zip: JSZip, path: string): Promise<string> {
  const file = zip.file(path);
  if (!file) throw new Error(`PPTX part missing: ${path}`);
  return file.async("string");
}

export function parseAttributes(tag: string): Record<string, string> {
  return Object.fromEntries([...tag.matchAll(/([\w:.-]+)="([^"]*)"/g)].map((m) => [m[1], m[2]]));
}

export function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

export function unescapeXml(text: string): string {
  return text
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec: string) => String.fromCodePoint(Number(dec)))
    .replace(/&amp;/g, "&");
}

export function relsPath(partPath: string): string {
  const slash = partPath.lastIndexOf("/");
  return `${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`;
}

/** Resolves a relationship target (relative to the source part, or absolute from the package root) to a zip path. */
export function resolveTarget(sourcePart: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const segments = sourcePart.split("/").slice(0, -1);
  for (const segment of target.split("/")) {
    if (segment === "..") segments.pop();
    else if (segment !== ".") segments.push(segment);
  }
  return segments.join("/");
}

export async function readRelationships(zip: JSZip, partPath: string): Promise<Relationship[]> {
  const file = zip.file(relsPath(partPath));
  if (!file) return [];
  const xml = await file.async("string");
  return [...xml.matchAll(/<Relationship\b[^>]*>/g)].map((m) => {
    const a = parseAttributes(m[0]);
    return { id: a.Id ?? "", type: a.Type ?? "", target: a.Target ?? "" };
  });
}

/** Finds the single relationship of `type` from `partPath` and returns the target's zip path, or null. */
export async function findRelatedPart(
  zip: JSZip,
  partPath: string,
  type: string,
): Promise<string | null> {
  const rel = (await readRelationships(zip, partPath)).find((r) => r.type === type);
  return rel ? resolveTarget(partPath, rel.target) : null;
}

/** Adds a relationship from `partPath` to the absolute part `targetPath` and returns its new ID. */
export async function addRelationship(
  zip: JSZip,
  partPath: string,
  type: string,
  targetPath: string,
): Promise<string> {
  const path = relsPath(partPath);
  const xml =
    (await zip.file(path)?.async("string")) ??
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${RELS_NS}"></Relationships>`;
  const used = new Set([...xml.matchAll(/\bId="([^"]*)"/g)].map((m) => m[1]));
  let n = 1;
  while (used.has(`rId${n}`)) n++;
  const id = `rId${n}`;
  const rel = `<Relationship Id="${id}" Type="${type}" Target="/${targetPath}"/>`;
  zip.file(path, xml.replace("</Relationships>", `${rel}</Relationships>`));
  return id;
}

export async function removeRelationship(
  zip: JSZip,
  partPath: string,
  relId: string,
): Promise<void> {
  const path = relsPath(partPath);
  const xml = await readPart(zip, path);
  zip.file(path, xml.replace(new RegExp(`<Relationship\\b[^>]*\\bId="${relId}"[^>]*/>`), ""));
}

export async function ensureContentTypeOverride(
  zip: JSZip,
  partPath: string,
  contentType: string,
): Promise<void> {
  const xml = await readPart(zip, "[Content_Types].xml");
  if (xml.includes(`PartName="/${partPath}"`)) return;
  const override = `<Override PartName="/${partPath}" ContentType="${contentType}"/>`;
  zip.file("[Content_Types].xml", xml.replace("</Types>", `${override}</Types>`));
}

export async function ensureContentTypeDefault(
  zip: JSZip,
  extension: string,
  contentType: string,
): Promise<void> {
  const xml = await readPart(zip, "[Content_Types].xml");
  if (new RegExp(`<Default\\b[^>]*Extension="${extension}"`, "i").test(xml)) return;
  const entry = `<Default Extension="${extension}" ContentType="${contentType}"/>`;
  zip.file(
    "[Content_Types].xml",
    xml.replace(/<Types\b[^>]*>/, (open) => open + entry),
  );
}

export async function removeContentTypeOverride(zip: JSZip, partPath: string): Promise<void> {
  const xml = await readPart(zip, "[Content_Types].xml");
  const partName = partPath.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  zip.file(
    "[Content_Types].xml",
    xml.replace(new RegExp(`<Override\\b[^>]*PartName="/${partName}"[^>]*/>`), ""),
  );
}

/** First path `${prefix}${n}${suffix}` (n ≥ 1) not present in the package. */
export function nextFreePath(zip: JSZip, prefix: string, suffix: string): string {
  let n = 1;
  while (zip.file(`${prefix}${n}${suffix}`)) n++;
  return `${prefix}${n}${suffix}`;
}

/** Slides in presentation order. */
export async function listSlideParts(zip: JSZip): Promise<SlidePart[]> {
  const xml = await readPart(zip, PRESENTATION_PART);
  const rels = await readRelationships(zip, PRESENTATION_PART);
  return [...xml.matchAll(/<p:sldId\b[^>]*>/g)].map((m) => {
    const a = parseAttributes(m[0]);
    const relId = a["r:id"] ?? "";
    const rel = rels.find((r) => r.id === relId);
    if (!rel) throw new Error(`presentation.xml references missing relationship ${relId}`);
    return { sldId: a.id ?? "", relId, path: resolveTarget(PRESENTATION_PART, rel.target) };
  });
}

/** The package's first slide: an `exportAsBase64` result contains exactly one. */
export async function firstSlidePath(zip: JSZip): Promise<string> {
  const [first] = await listSlideParts(zip);
  if (!first) throw new Error("PPTX has no slides");
  return first.path;
}

export async function readSlideSize(zip: JSZip): Promise<SlideSize> {
  const xml = await readPart(zip, PRESENTATION_PART);
  const tag = xml.match(/<p:sldSz\b[^>]*>/)?.[0];
  if (!tag) throw new Error("presentation.xml has no <p:sldSz>");
  const { cx, cy } = parseAttributes(tag);
  return { width: Number(cx) / EMU_PER_POINT, height: Number(cy) / EMU_PER_POINT };
}

/** Next free `<p:cNvPr id>` in a slide's shape tree. */
export function nextShapeId(slideXml: string): number {
  const ids = [...slideXml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map((m) => Number(m[1]));
  return Math.max(0, ...ids) + 1;
}

export function appendToShapeTree(slideXml: string, elementXml: string): string {
  const end = slideXml.lastIndexOf("</p:spTree>");
  if (end < 0) throw new Error("Slide XML has no </p:spTree>");
  return slideXml.slice(0, end) + elementXml + slideXml.slice(end);
}
