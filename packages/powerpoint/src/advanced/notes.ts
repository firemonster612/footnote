// Speaker notes on the first slide of a PPTX package (an `exportAsBase64` result).
// Office.js has no notes API, so notes go through export → edit notesSlide XML → re-insert.

import type JSZip from "jszip";
import PptxGenJS from "pptxgenjs";
import {
  PRESENTATION_PART,
  addRelationship,
  appendToShapeTree,
  contentTypes,
  ensureContentTypeOverride,
  escapeXml,
  findRelatedPart,
  firstSlidePath,
  loadPptx,
  nextFreePath,
  nextShapeId,
  readPart,
  relTypes,
  savePptx,
  unescapeXml,
} from "./ooxml.ts";

const SHAPE_PATTERN = /<p:sp(?:\s[^>]*)?>[\s\S]*?<\/p:sp>/g;
const BODY_PLACEHOLDER_PATTERN = /<p:ph\b[^>]*\btype="body"/;
const PARAGRAPH_PATTERN = /<a:p(?:\s[^>]*)?\/>|<a:p(?:\s[^>]*)?>([\s\S]*?)<\/a:p>/g;
const RUN_TEXT_PATTERN = /<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>|<a:br\b/g;

/** Notes text of the package's first slide; "" when it has none. Paragraphs are joined with "\n". */
export async function readNotes(pptxBase64: string): Promise<string> {
  const zip = await loadPptx(pptxBase64);
  const notesPath = await findRelatedPart(zip, await firstSlidePath(zip), relTypes.notesSlide);
  if (!notesPath) return "";
  const body = findBodyShape(await readPart(zip, notesPath));
  return body ? paragraphsText(body) : "";
}

/** Replaces the notes of the package's first slide, creating the notes slide (and notes master) when absent. */
export async function writeNotes(pptxBase64: string, text: string): Promise<string> {
  const zip = await loadPptx(pptxBase64);
  const slidePath = await firstSlidePath(zip);
  const notesPath =
    (await findRelatedPart(zip, slidePath, relTypes.notesSlide)) ??
    (await createNotesSlide(zip, slidePath));
  zip.file(notesPath, replaceBodyText(await readPart(zip, notesPath), text));
  return savePptx(zip);
}

function findBodyShape(notesXml: string): string | undefined {
  return notesXml.match(SHAPE_PATTERN)?.find((shape) => BODY_PLACEHOLDER_PATTERN.test(shape));
}

function paragraphsText(shapeXml: string): string {
  return [...shapeXml.matchAll(PARAGRAPH_PATTERN)]
    .map((p) =>
      [...(p[1] ?? "").matchAll(RUN_TEXT_PATTERN)]
        .map((m) => (m[1] === undefined ? "\n" : unescapeXml(m[1])))
        .join(""),
    )
    .join("\n");
}

function paragraphsXml(text: string): string {
  return text
    .split(/\r?\n/)
    .map((line) =>
      line
        ? `<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${escapeXml(line)}</a:t></a:r></a:p>`
        : "<a:p/>",
    )
    .join("");
}

function replaceBodyText(notesXml: string, text: string): string {
  const paragraphs = paragraphsXml(text);
  const body = findBodyShape(notesXml);
  if (!body) return appendToShapeTree(notesXml, bodyShapeXml(nextShapeId(notesXml), paragraphs));

  const txBody = body.match(/<p:txBody>([\s\S]*?)<\/p:txBody>/);
  let newBody: string;
  if (txBody) {
    // Keep <a:bodyPr>/<a:lstStyle>; paragraphs are always the trailing children of txBody.
    const inner = txBody[1] ?? "";
    const firstParagraph = inner.search(/<a:p[\s>/]/);
    const kept = firstParagraph < 0 ? inner : inner.slice(0, firstParagraph);
    newBody = body.replace(txBody[0], () => `<p:txBody>${kept}${paragraphs}</p:txBody>`);
  } else {
    newBody = body.replace(
      "</p:sp>",
      () => `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`,
    );
  }
  // Function replacers: user text may contain "$&"-style replacement patterns.
  return notesXml.replace(body, () => newBody);
}

function bodyShapeXml(id: number, paragraphs: string): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Notes Placeholder ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
    `<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`
  );
}

const EMPTY_NOTES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
  `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>` +
  `<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>` +
  `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;

async function createNotesSlide(zip: JSZip, slidePath: string): Promise<string> {
  const masterPath =
    (await findRelatedPart(zip, PRESENTATION_PART, relTypes.notesMaster)) ??
    (await addNotesMaster(zip));
  const notesPath = nextFreePath(zip, "ppt/notesSlides/notesSlide", ".xml");
  zip.file(notesPath, EMPTY_NOTES_XML);
  await addRelationship(zip, notesPath, relTypes.notesMaster, masterPath);
  await addRelationship(zip, notesPath, relTypes.slide, slidePath);
  await addRelationship(zip, slidePath, relTypes.notesSlide, notesPath);
  await ensureContentTypeOverride(zip, notesPath, contentTypes.notesSlide);
  return notesPath;
}

/** Notes slides need a notes master. Borrow PptxGenJS's and point it at this package's theme. */
async function addNotesMaster(zip: JSZip): Promise<string> {
  const themePath = await findRelatedPart(zip, PRESENTATION_PART, relTypes.theme);
  if (!themePath) throw new Error("PPTX has no presentation theme to attach a notes master to");

  const template = new PptxGenJS();
  template.addSlide();
  const templateZip = await loadPptx((await template.write({ outputType: "base64" })) as string);
  const masterXml = await readPart(templateZip, "ppt/notesMasters/notesMaster1.xml");

  const masterPath = nextFreePath(zip, "ppt/notesMasters/notesMaster", ".xml");
  zip.file(masterPath, masterXml);
  await addRelationship(zip, masterPath, relTypes.theme, themePath);
  await ensureContentTypeOverride(zip, masterPath, contentTypes.notesMaster);

  const relId = await addRelationship(zip, PRESENTATION_PART, relTypes.notesMaster, masterPath);
  const presentation = await readPart(zip, PRESENTATION_PART);
  // Schema order: sldMasterIdLst, notesMasterIdLst, handoutMasterIdLst, sldIdLst, ...
  const list = `<p:notesMasterIdLst><p:notesMasterId r:id="${relId}"/></p:notesMasterIdLst>`;
  zip.file(
    PRESENTATION_PART,
    presentation.replace("</p:sldMasterIdLst>", `</p:sldMasterIdLst>${list}`),
  );
  return masterPath;
}
