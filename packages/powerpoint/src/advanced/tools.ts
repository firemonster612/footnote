// Agent-side tools for what Office.js can't do directly: charts, speaker notes, slide import.
// All three edit PPTX packages (export → JSZip → re-insert), so slides they touch come back with new IDs.

import type { FootnoteTool, ToolEnv } from "@footnote/core/contracts";
import { Type, type TSchema } from "typebox";
import type { DeckInfo, WriteReceipt } from "../ops/types.ts";
import {
  buildChartPptx,
  chartKinds,
  chartShapeName,
  legendPositions,
  mergeChartIntoSlide,
  validateChartSpec,
  type ChartSpec,
} from "./charts.ts";
import type { AdvancedToolDeps } from "./deps.ts";
import { readNotes, writeNotes } from "./notes.ts";
import { loadPptx, readSlideSize, type SlideSize } from "./ooxml.ts";
import type { ReplaceSlideArgs, ReplaceSlideResult } from "./ops.ts";
import { keepSlides } from "./slide-import.ts";

/** Used when the host can't report slide size (no PowerPointApi 1.10) and the deck has no slide to export. */
const DEFAULT_SLIDE_SIZE: SlideSize = { width: 960, height: 540 };

function defineTool<T extends TSchema>(tool: FootnoteTool<T>): FootnoteTool {
  // execute's params are contravariant, so a specific tool isn't assignable to FootnoteTool<TSchema>.
  // Sound because the agent validates arguments against `parameters` before calling execute.
  return tool as FootnoteTool;
}

function jsonResult(value: unknown) {
  return { content: [{ type: "text" as const, text: JSON.stringify(value) }], details: value };
}

const position = Type.Optional(
  Type.Integer({
    minimum: 1,
    description:
      "1-based slide number the first inserted slide should get. Omit to append at the end.",
  }),
);

const chartParameters = Type.Object({
  slideId: Type.Optional(
    Type.String({
      description:
        "Existing slide to add the chart to. Omit to insert the chart on a new blank slide.",
    }),
  ),
  position,
  type: Type.Enum(chartKinds),
  categories: Type.Array(Type.String(), {
    description:
      "Category labels along the X axis (or pie slices). For scatter: the numeric X values.",
  }),
  series: Type.Array(
    Type.Object({
      name: Type.String(),
      values: Type.Array(Type.Number(), { description: "One value per category." }),
    }),
    { minItems: 1 },
  ),
  bounds: Type.Optional(
    Type.Object(
      { left: Type.Number(), top: Type.Number(), width: Type.Number(), height: Type.Number() },
      {
        description:
          "Chart frame in points. Defaults to the area below a title (85% wide, 70% tall).",
      },
    ),
  ),
  title: Type.Optional(Type.String()),
  legend: Type.Optional(
    Type.Enum(legendPositions, {
      description: "Default: bottom when there are several series or slices.",
    }),
  ),
  dataLabels: Type.Optional(Type.Boolean({ description: "Show values on bars/points/slices." })),
  numberFormat: Type.Optional(
    Type.String({
      description: 'Excel number format for the value axis and labels, e.g. "0%", "$#,##0".',
    }),
  ),
  colors: Type.Optional(
    Type.Array(Type.String(), {
      description:
        "#RRGGBB per series (per slice for pie/doughnut). Use theme colors from <deck_state>.",
    }),
  ),
  fontFace: Type.Optional(
    Type.String({ description: "Font for all chart text; use the theme body font." }),
  ),
});

export function createAdvancedTools(env: ToolEnv, deps: AdvancedToolDeps): FootnoteTool[] {
  const { host } = env;
  const exportSlide = (slideId: string) => host.call<string>("export_slide", { slideId });
  const replaceSlide = (args: ReplaceSlideArgs) =>
    host.call<ReplaceSlideResult>("replace_slide", args);
  const insertSlides = (args: {
    base64: string;
    targetSlideId?: string;
    formatting: "KeepSourceFormatting" | "UseDestinationTheme";
  }) => host.call<WriteReceipt>("insert_slides", args);

  /** Slide ID to insert after for a 1-based position (undefined = beginning); appends when position is omitted. */
  function insertAfter(deck: DeckInfo, position: number | undefined): string | undefined {
    const ids = deck.slides.map((s) => s.id);
    if (position === undefined) return ids.at(-1);
    if (position > ids.length + 1)
      throw new Error(`Position ${position} is past the end; the deck has ${ids.length} slides.`);
    return ids[position - 2];
  }

  async function deckSlideSize(deck: DeckInfo): Promise<SlideSize> {
    if (deck.slideWidth && deck.slideHeight)
      return { width: deck.slideWidth, height: deck.slideHeight };
    const first = deck.slides[0];
    return first ? readSlideSize(await loadPptx(await exportSlide(first.id))) : DEFAULT_SLIDE_SIZE;
  }

  const getNotes = defineTool({
    name: "get_notes",
    label: "Read speaker notes",
    access: "read",
    description:
      "Read the speaker notes of one or more slides. Returns plain text per slide (paragraphs separated by newlines).",
    parameters: Type.Object({ slideIds: Type.Array(Type.String(), { minItems: 1 }) }),
    describeCall: ({ slideIds }: { slideIds: string[] }) =>
      `Read notes of ${slideIds.length} slide(s)`,
    async execute(_id, { slideIds }) {
      const notes = [];
      for (const slideId of slideIds)
        notes.push({ slideId, text: await readNotes(await exportSlide(slideId)) });
      return jsonResult({ notes });
    },
  });

  const setNotes = defineTool({
    name: "set_notes",
    label: "Write speaker notes",
    access: "write",
    description:
      "Replace the speaker notes of one or more slides (batch them in one call). Plain text; newlines start paragraphs. " +
      "Office.js has no notes API, so each slide is exported, edited and re-inserted at the same position: " +
      "its slide ID changes. Use the new IDs from the result for any later call.",
    parameters: Type.Object({
      notes: Type.Array(Type.Object({ slideId: Type.String(), text: Type.String() }), {
        minItems: 1,
      }),
    }),
    describeCall: ({ notes }: { notes: unknown[] }) => `Write notes on ${notes.length} slide(s)`,
    async execute(_id, { notes }) {
      await deps.beforeWrite(notes.map((n) => n.slideId));
      const replaced: Record<string, string> = {};
      const verified: Record<string, boolean> = {};
      const failed: { slideId: string; error: string }[] = [];
      for (const { slideId, text } of notes) {
        try {
          const edited = await writeNotes(await exportSlide(slideId), text);
          const { slideId: newId } = await replaceSlide({ slideId, base64: edited });
          replaced[slideId] = newId;
          await deps.afterWrite({ createdSlideIds: [newId] });
          verified[newId] =
            (await readNotes(await exportSlide(newId))) === text.replace(/\r\n/g, "\n");
        } catch (error) {
          failed.push({ slideId, error: error instanceof Error ? error.message : String(error) });
        }
      }
      const result = {
        changed: Object.values(replaced),
        replacedSlideIds: replaced,
        verified: { notesMatch: verified },
        warnings: [
          "Slides were re-inserted to edit notes: old slide IDs are gone (see replacedSlideIds).",
        ],
        ...(failed.length > 0 && { failed }),
      };
      return { ...jsonResult(result), ...(failed.length === notes.length && { isError: true }) };
    },
  });

  const insertChart = defineTool({
    name: "insert_chart",
    label: "Insert chart",
    access: "write",
    description:
      "Insert a native, editable PowerPoint chart (bar, column, line, pie, doughnut, area, scatter). " +
      "With slideId, the chart is added to that existing slide; everything already on it is kept, but the slide is " +
      "re-inserted and gets a new ID (returned). For a chart on a new slide that uses a deck layout, add_slide first, then " +
      "call this with its slideId; omitting slideId inserts a blank slide holding only the chart. " +
      "Bounds are in points. Pass theme colors and fonts so the chart matches the deck. Edit the chart's data later by calling again " +
      "and deleting the old chart shape.",
    parameters: chartParameters,
    describeCall: ({ type, slideId, title }: { type: string; slideId?: string; title?: string }) =>
      `Insert ${type} chart${title ? ` "${title}"` : ""} ${slideId ? `on slide ${slideId}` : "on a new slide"}`,
    async execute(_id, { slideId, position, bounds, ...chart }) {
      const boundsFor = (size: SlideSize) =>
        bounds ?? {
          left: size.width * 0.075,
          top: size.height * 0.22,
          width: size.width * 0.85,
          height: size.height * 0.7,
        };

      validateChartSpec(chart);

      if (slideId) {
        await deps.beforeWrite([slideId]);
        const exported = await exportSlide(slideId);
        const size = await readSlideSize(await loadPptx(exported));
        const spec: ChartSpec = { ...chart, bounds: boundsFor(size) };
        const merged = await mergeChartIntoSlide(exported, await buildChartPptx(spec, size));
        const { slideId: newId, shapes } = await replaceSlide({ slideId, base64: merged });
        await deps.afterWrite({ createdSlideIds: [newId] });
        const chartShape = shapes.find(
          (s) => s.type === "Chart" && s.name === chartShapeName(spec),
        );
        return jsonResult({
          changed: [newId],
          replacedSlideIds: { [slideId]: newId },
          verified: { chartShapeId: chartShape?.id ?? null, shapeCount: shapes.length },
          warnings: [
            `Slide ${slideId} was re-inserted with the chart and is now ${newId}.`,
            ...(chartShape
              ? []
              : ["Chart shape not found on the re-inserted slide; call get_slide to check."]),
          ],
        });
      }

      const deck = await host.call<DeckInfo>("get_deck");
      const size = await deckSlideSize(deck);
      const spec: ChartSpec = { ...chart, bounds: boundsFor(size) };
      const targetSlideId = insertAfter(deck, position);
      const receipt = await insertSlides({
        base64: await buildChartPptx(spec, size),
        formatting: "UseDestinationTheme",
        ...(targetSlideId && { targetSlideId }),
      });
      const createdSlideIds = receipt.createdSlideIds ?? [];
      await deps.afterWrite({ createdSlideIds });
      return jsonResult({
        changed: createdSlideIds,
        verified: receipt.verified,
        warnings: receipt.warnings,
      });
    },
  });

  const insertSlidesFromFile = defineTool({
    name: "insert_slides_from_file",
    label: "Insert slides from file",
    access: "write",
    description:
      "Copy slides from an attached .pptx into this deck. Choose slides by their 1-based numbers in the attachment " +
      "(see the attachment's slide list via read_attachment). By default the slides adopt this deck's theme; set " +
      "keepSourceFormatting to keep their original look.",
    parameters: Type.Object({
      attachmentId: Type.String(),
      slides: Type.Optional(
        Type.Array(Type.Integer({ minimum: 1 }), {
          minItems: 1,
          description: "1-based slide numbers in the file. Omit for all.",
        }),
      ),
      position,
      keepSourceFormatting: Type.Optional(Type.Boolean()),
    }),
    describeCall: ({ slides }: { slides?: number[] }) =>
      `Insert ${slides ? `${slides.length} slide(s)` : "all slides"} from an attached presentation`,
    async execute(_id, { attachmentId, slides, position, keepSourceFormatting }) {
      const attachment = env.attachments.get(attachmentId);
      if (attachment?.kind !== "pptx") {
        const available = env.attachments
          .list()
          .filter((a) => a.kind === "pptx")
          .map((a) => `${a.id} (${a.name})`);
        throw new Error(
          `No .pptx attachment ${attachmentId}. Available: ${available.join(", ") || "none"}.`,
        );
      }
      const deck = await host.call<DeckInfo>("get_deck");
      const targetSlideId = insertAfter(deck, position);
      const receipt = await insertSlides({
        base64: slides ? await keepSlides(attachment.base64, slides) : attachment.base64,
        formatting: keepSourceFormatting ? "KeepSourceFormatting" : "UseDestinationTheme",
        ...(targetSlideId && { targetSlideId }),
      });
      const createdSlideIds = receipt.createdSlideIds ?? [];
      await deps.afterWrite({ createdSlideIds });
      return jsonResult({
        changed: createdSlideIds,
        verified: receipt.verified,
        warnings: receipt.warnings,
      });
    },
  });

  return [getNotes, setNotes, insertChart, insertSlidesFromFile];
}
