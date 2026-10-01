// Native charts: PptxGenJS generates a one-slide PPTX holding the chart; the merge path grafts that chart
// (chart part + embedded workbook + rels + content types + graphic frame) onto an exported slide.

import PptxGenJS from "pptxgenjs";
import type { Bounds } from "../ops/types.ts";
import {
  addRelationship,
  appendToShapeTree,
  contentTypes,
  ensureContentTypeDefault,
  ensureContentTypeOverride,
  firstSlidePath,
  loadPptx,
  nextFreePath,
  nextShapeId,
  readPart,
  readRelationships,
  relTypes,
  resolveTarget,
  savePptx,
  type SlideSize,
} from "./ooxml.ts";

export const chartKinds = ["bar", "column", "line", "pie", "doughnut", "area", "scatter"] as const;
export type ChartKind = (typeof chartKinds)[number];

export const legendPositions = ["top", "bottom", "left", "right", "none"] as const;
export type LegendPosition = (typeof legendPositions)[number];

export interface ChartSpec {
  type: ChartKind;
  /** Category labels; for scatter, the numeric X values. */
  categories: string[];
  series: { name: string; values: number[] }[];
  /** Points. */
  bounds: Bounds;
  title?: string;
  /** Defaults to shown when there are several series or slices. */
  legend?: LegendPosition;
  dataLabels?: boolean;
  /** Excel number format for value axis and data labels, e.g. "0%" or "$#,##0". */
  numberFormat?: string;
  /** Hex colors (with or without "#"), one per series (or per slice for pie/doughnut). */
  colors?: string[];
  fontFace?: string;
}

const POINTS_PER_INCH = 72;
const legendCodes = { top: "t", bottom: "b", left: "l", right: "r" } as const;

export function chartShapeName(spec: Pick<ChartSpec, "title">): string {
  return spec.title ? `Chart: ${spec.title}` : "Chart";
}

/** Throws a model-facing message when the data doesn't fit the chart type. */
export function validateChartSpec(spec: Pick<ChartSpec, "type" | "categories" | "series">): void {
  if (spec.series.length === 0) throw new Error("A chart needs at least one series.");
  for (const s of spec.series) {
    if (s.values.length !== spec.categories.length) {
      throw new Error(
        `Series "${s.name}" has ${s.values.length} values but there are ${spec.categories.length} categories.`,
      );
    }
  }
  if (spec.type === "scatter" && spec.categories.some((c) => Number.isNaN(Number(c)))) {
    throw new Error("Scatter charts need numeric categories (the X values).");
  }
}

/** A one-slide PPTX of the deck's size with the chart at `spec.bounds`. */
export async function buildChartPptx(spec: ChartSpec, slideSize: SlideSize): Promise<string> {
  const pptx = new PptxGenJS();
  pptx.defineLayout({
    name: "deck",
    width: slideSize.width / POINTS_PER_INCH,
    height: slideSize.height / POINTS_PER_INCH,
  });
  pptx.layout = "deck";

  const isPie = spec.type === "pie" || spec.type === "doughnut";
  const legend = spec.legend ?? (spec.series.length > 1 || isPie ? "bottom" : "none");
  const data =
    spec.type === "scatter"
      ? [{ name: "X", values: spec.categories.map(Number) }, ...spec.series]
      : spec.series.map((s) => ({ name: s.name, labels: spec.categories, values: s.values }));
  const chartType = spec.type === "column" ? pptx.ChartType.bar : pptx.ChartType[spec.type];
  const { left, top, width, height } = spec.bounds;

  pptx.addSlide().addChart(chartType, data, {
    x: left / POINTS_PER_INCH,
    y: top / POINTS_PER_INCH,
    w: width / POINTS_PER_INCH,
    h: height / POINTS_PER_INCH,
    objectName: chartShapeName(spec),
    ...(spec.type === "bar" && { barDir: "bar" }),
    ...(spec.type === "column" && { barDir: "col" }),
    ...(spec.title && { showTitle: true, title: spec.title }),
    showLegend: legend !== "none",
    ...(legend !== "none" && { legendPos: legendCodes[legend] }),
    ...(spec.dataLabels && { showValue: true }),
    ...(spec.numberFormat && {
      dataLabelFormatCode: spec.numberFormat,
      valAxisLabelFormatCode: spec.numberFormat,
    }),
    ...(spec.colors && { chartColors: spec.colors.map((c) => c.replace(/^#/, "")) }),
    ...(spec.fontFace && {
      titleFontFace: spec.fontFace,
      legendFontFace: spec.fontFace,
      catAxisLabelFontFace: spec.fontFace,
      valAxisLabelFontFace: spec.fontFace,
      dataLabelFontFace: spec.fontFace,
    }),
  });
  return (await pptx.write({ outputType: "base64" })) as string;
}

/**
 * Copies the chart from `chartPptx` (as made by buildChartPptx) onto the first slide of `targetPptx`
 * (an exported slide), keeping everything already on that slide. Returns the edited target package.
 */
export async function mergeChartIntoSlide(targetPptx: string, chartPptx: string): Promise<string> {
  const source = await loadPptx(chartPptx);
  const target = await loadPptx(targetPptx);

  const sourceSlidePath = await firstSlidePath(source);
  const sourceSlide = await readPart(source, sourceSlidePath);
  const chartRel = (await readRelationships(source, sourceSlidePath)).find(
    (r) => r.type === relTypes.chart,
  );
  const frame = sourceSlide.match(/<p:graphicFrame\b[\s\S]*?<\/p:graphicFrame>/)?.[0];
  if (!chartRel || !frame) throw new Error("Chart PPTX has no chart on its slide");
  const sourceChartPath = resolveTarget(sourceSlidePath, chartRel.target);

  const chartPath = nextFreePath(target, "ppt/charts/chart", ".xml");
  target.file(chartPath, await readPart(source, sourceChartPath));
  await ensureContentTypeOverride(target, chartPath, contentTypes.chart);

  for (const rel of await readRelationships(source, sourceChartPath)) {
    if (rel.type !== relTypes.package) throw new Error(`Unexpected chart relationship ${rel.type}`);
    const sourceWorkbookPath = resolveTarget(sourceChartPath, rel.target);
    const workbook = source.file(sourceWorkbookPath);
    if (!workbook) throw new Error(`PPTX part missing: ${sourceWorkbookPath}`);
    const workbookPath = nextFreePath(target, "ppt/embeddings/Microsoft_Excel_Worksheet", ".xlsx");
    target.file(workbookPath, await workbook.async("uint8array"));
    await addRelationship(target, chartPath, rel.type, workbookPath);
  }
  await ensureContentTypeDefault(target, "xlsx", contentTypes.xlsx);

  const targetSlidePath = await firstSlidePath(target);
  const relId = await addRelationship(target, targetSlidePath, relTypes.chart, chartPath);
  const targetSlide = await readPart(target, targetSlidePath);
  const graftedFrame = frame
    .replace(/(<p:cNvPr\b[^>]*\bid=")\d+"/, `$1${nextShapeId(targetSlide)}"`)
    .replace(`r:id="${chartRel.id}"`, `r:id="${relId}"`);
  target.file(targetSlidePath, appendToShapeTree(targetSlide, graftedFrame));
  return savePptx(target);
}
