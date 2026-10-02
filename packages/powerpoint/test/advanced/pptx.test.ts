import PptxGenJS from "pptxgenjs";
import { describe, expect, it } from "vitest";
import { buildChartPptx, mergeChartIntoSlide, type ChartSpec } from "../../src/advanced/charts.ts";
import { readNotes, writeNotes } from "../../src/advanced/notes.ts";
import {
  PRESENTATION_PART,
  findRelatedPart,
  listSlideParts,
  loadPptx,
  readPart,
  readRelationships,
  readSlideSize,
  relTypes,
  relsPath,
  resolveTarget,
  savePptx,
} from "../../src/advanced/ooxml.ts";
import { keepSlides } from "../../src/advanced/slide-import.ts";

type SlideFiller = (slide: PptxGenJS.Slide, pptx: PptxGenJS) => void;

async function makePptx(...slides: SlideFiller[]): Promise<string> {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE"; // 13.33 x 7.5 in
  for (const fill of slides) fill(pptx.addSlide(), pptx);
  return (await pptx.write({ outputType: "base64" })) as string;
}

const text =
  (t: string): SlideFiller =>
  (slide) =>
    slide.addText(t, { x: 1, y: 1, w: 4, h: 1 });

/** Mimics a deck with no notes at all: drops the notes slide and the notes master. */
async function withoutNotes(base64: string): Promise<string> {
  const zip = await loadPptx(base64);
  const slidePath = (await listSlideParts(zip))[0]!.path;
  const slideRels = await readPart(zip, relsPath(slidePath));
  zip.file(relsPath(slidePath), slideRels.replace(/<Relationship\b[^>]*notesSlide"[^>]*\/>/, ""));
  for (const part of await zip.file(/notesSlides|notesMasters/)) zip.remove(part.name);
  const presentation = await readPart(zip, PRESENTATION_PART);
  zip.file(
    PRESENTATION_PART,
    presentation.replace(/<p:notesMasterIdLst>[\s\S]*?<\/p:notesMasterIdLst>/, ""),
  );
  const presRels = await readPart(zip, relsPath(PRESENTATION_PART));
  zip.file(
    relsPath(PRESENTATION_PART),
    presRels.replace(/<Relationship\b[^>]*notesMaster"[^>]*\/>/, ""),
  );
  return savePptx(zip);
}

describe("speaker notes", () => {
  it("reads notes written by PptxGenJS", async () => {
    const pptx = await makePptx((slide) => {
      slide.addText("Hi", { x: 1, y: 1, w: 4, h: 1 });
      slide.addNotes("Open with the Q3 number.");
    });
    expect(await readNotes(pptx)).toBe("Open with the Q3 number.");
  });

  it("returns empty notes for a slide without a notes part", async () => {
    expect(await readNotes(await withoutNotes(await makePptx(text("Hi"))))).toBe("");
  });

  it("replaces existing notes, keeping paragraphs and escaping XML", async () => {
    const pptx = await makePptx((slide) => {
      slide.addText("Hi", { x: 1, y: 1, w: 4, h: 1 });
      slide.addNotes("old notes");
    });
    const notes = 'Revenue <up> 12% & "steady"\n\nCosts cost $& more';
    const edited = await writeNotes(pptx, notes);
    expect(await readNotes(edited)).toBe(notes);
  });

  it("creates the notes slide and notes master when the deck has neither", async () => {
    const edited = await writeNotes(await withoutNotes(await makePptx(text("Hi"))), "Fresh notes");
    expect(await readNotes(edited)).toBe("Fresh notes");

    const zip = await loadPptx(edited);
    const presentation = await readPart(zip, PRESENTATION_PART);
    expect(presentation).toMatch(
      /<\/p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rId\d+"\/><\/p:notesMasterIdLst>/,
    );
    const masterPath = await findRelatedPart(zip, PRESENTATION_PART, relTypes.notesMaster);
    expect(masterPath).toBe("ppt/notesMasters/notesMaster1.xml");
    expect(await findRelatedPart(zip, masterPath!, relTypes.theme)).toBe("ppt/theme/theme1.xml");

    const notesPath = await findRelatedPart(
      zip,
      (await listSlideParts(zip))[0]!.path,
      relTypes.notesSlide,
    );
    const types = await readPart(zip, "[Content_Types].xml");
    expect(types).toContain(`PartName="/${notesPath}"`);
    expect(types).toContain('PartName="/ppt/notesMasters/notesMaster1.xml"');
  });
});

describe("charts", () => {
  const spec: ChartSpec = {
    type: "column",
    categories: ["Q1", "Q2", "Q3"],
    series: [
      { name: "2025", values: [1, 2, 3] },
      { name: "2026", values: [2, 3, 4] },
    ],
    bounds: { left: 72, top: 108, width: 576, height: 324 },
    title: "Revenue",
  };

  it("generates a chart PPTX at the deck's slide size", async () => {
    const pptx = await buildChartPptx(spec, { width: 720, height: 405 });
    const zip = await loadPptx(pptx);
    expect(await readSlideSize(zip)).toEqual({ width: 720, height: 405 });
    const slidePath = (await listSlideParts(zip))[0]!.path;
    const chartPath = await findRelatedPart(zip, slidePath, relTypes.chart);
    expect(chartPath).toMatch(/^ppt\/charts\/chart\d+\.xml$/);
    expect(await readPart(zip, chartPath!)).toContain('<c:barDir val="col"/>');
    const workbook = await findRelatedPart(zip, chartPath!, relTypes.package);
    expect(zip.file(workbook!)).not.toBeNull();
    expect(await readPart(zip, slidePath)).toContain('name="Chart: Revenue"');
  });

  it("merges charts onto a slide, keeping its content and avoiding part and ID collisions", async () => {
    const target = await makePptx((slide, pptx) => {
      slide.addText("Title", { x: 1, y: 0.5, w: 6, h: 1 });
      slide.addChart(
        pptx.ChartType.pie,
        [{ name: "Share", labels: ["A", "B"], values: [60, 40] }],
        { x: 7, y: 2, w: 4, h: 4 },
      );
    });
    const chart = await buildChartPptx(spec, await readSlideSize(await loadPptx(target)));
    // Merging twice forces the second chart to dodge the part names the first one took.
    const merged = await loadPptx(
      await mergeChartIntoSlide(await mergeChartIntoSlide(target, chart), chart),
    );

    const slidePath = (await listSlideParts(merged))[0]!.path;
    const slideXml = await readPart(merged, slidePath);
    expect(slideXml).toContain("Title");
    expect(slideXml.match(/<p:graphicFrame\b/g)).toHaveLength(3);
    const shapeIds = [...slideXml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(shapeIds).size).toBe(shapeIds.length);

    const chartRels = (await readRelationships(merged, slidePath)).filter(
      (r) => r.type === relTypes.chart,
    );
    const frameRelIds = [...slideXml.matchAll(/<c:chart\b[^>]*r:id="([^"]+)"/g)].map((m) => m[1]);
    expect(frameRelIds.sort()).toEqual(chartRels.map((r) => r.id).sort());

    const charts = chartRels.map((r) => resolveTarget(slidePath, r.target));
    const workbooks = await Promise.all(
      charts.map((c) => findRelatedPart(merged, c, relTypes.package)),
    );
    expect(new Set(charts).size).toBe(3);
    expect(new Set(workbooks).size).toBe(3);
    for (const workbook of workbooks) expect(merged.file(workbook!)).not.toBeNull();
    const chartXml = await Promise.all(charts.map((c) => readPart(merged, c)));
    expect(chartXml.filter((xml) => xml.includes('<c:barDir val="col"/>'))).toHaveLength(2);

    const types = await readPart(merged, "[Content_Types].xml");
    for (const c of charts) expect(types).toContain(`PartName="/${c}"`);
  });
});

describe("keepSlides", () => {
  it("keeps the selected slides with their notes and drops the rest", async () => {
    const withNotes =
      (t: string): SlideFiller =>
      (slide) => {
        slide.addText(t, { x: 1, y: 1, w: 4, h: 1 });
        slide.addNotes(`notes ${t}`);
      };
    const source = await makePptx(withNotes("one"), withNotes("two"), withNotes("three"));
    const sourceRelCount = (await readRelationships(await loadPptx(source), PRESENTATION_PART))
      .length;
    const reduced = await loadPptx(await keepSlides(source, [3, 1]));

    const slides = await listSlideParts(reduced);
    expect(slides.map((s) => s.path)).toEqual(["ppt/slides/slide1.xml", "ppt/slides/slide3.xml"]);
    expect(await readPart(reduced, "ppt/slides/slide3.xml")).toContain("three");
    expect(reduced.file("ppt/slides/slide2.xml")).toBeNull();
    expect(reduced.file("ppt/notesSlides/notesSlide2.xml")).toBeNull();
    const types = await readPart(reduced, "[Content_Types].xml");
    expect(types).not.toContain("/ppt/slides/slide2.xml");
    expect(types).not.toContain("/ppt/notesSlides/notesSlide2.xml");
    expect(await readRelationships(reduced, PRESENTATION_PART)).toHaveLength(sourceRelCount - 1);
  });

  it("drops links and section entries that point at removed slides", async () => {
    const pptx = new PptxGenJS();
    pptx.addSection({ title: "Main" });
    pptx
      .addSlide({ sectionTitle: "Main" })
      .addText("Jump", { x: 1, y: 1, w: 4, h: 1, hyperlink: { slide: 2 } });
    pptx.addSlide({ sectionTitle: "Main" }).addText("Target", { x: 1, y: 1, w: 4, h: 1 });
    const source = (await pptx.write({ outputType: "base64" })) as string;

    const reduced = await loadPptx(await keepSlides(source, [1]));

    const rels = await readRelationships(reduced, "ppt/slides/slide1.xml");
    expect(rels.filter((rel) => rel.type === relTypes.slide)).toEqual([]);
    const slideXml = await readPart(reduced, "ppt/slides/slide1.xml");
    expect(slideXml).toContain("Jump");
    expect(slideXml).not.toContain("hlinkClick");
    const presentation = await readPart(reduced, PRESENTATION_PART);
    expect(presentation).toContain('<p14:sldId id="256"/>');
    expect(presentation).not.toContain('<p14:sldId id="257"/>');
  });

  it("rejects slide numbers outside the file", async () => {
    await expect(keepSlides(await makePptx(text("one")), [2])).rejects.toThrow(
      "out of range; the file has 1 slides",
    );
  });
});
