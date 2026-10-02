import { fingerprintShapes } from "./fingerprint.ts";
import { findSlide, loadSlides, requireApi, supportsApi } from "./presentation.ts";
import { readShapeLists } from "./shapeReader.ts";
import type {
  DeckInfo,
  DeckState,
  LayoutsInfo,
  SelectionInfo,
  ShapeInfo,
  SlideDetail,
  SlideOutline,
  SlideState,
  ThemeInfo,
} from "./types.ts";

const TITLE_PLACEHOLDERS = new Set(["Title", "CenterTitle", "VerticalTitle"]);
const BODY_PLACEHOLDERS = new Set(["Body", "Content", "Subtitle", "VerticalBody"]);
const THEME_COLORS = [
  "Dark1",
  "Light1",
  "Dark2",
  "Light2",
  "Accent1",
  "Accent2",
  "Accent3",
  "Accent4",
  "Accent5",
  "Accent6",
  "Hyperlink",
  "FollowedHyperlink",
] as const;
export const MAX_RENDER_WIDTH = 1280;

// ---------------------------------------------------------------------------
// Context-level readers, shared with the write ops' receipts.
// ---------------------------------------------------------------------------

export function slideTitle(shapes: ShapeInfo[]): string | undefined {
  const title = shapes.find(
    (shape) => shape.placeholder && TITLE_PLACEHOLDERS.has(shape.placeholder) && shape.text?.trim(),
  );
  return title?.text
    ?.trim()
    .split(/[\r\n\v]/)[0]
    ?.slice(0, 100);
}

/** A slide's outline from its summary shape tree. `slide` needs id and layout/name loaded. */
export function toOutline(
  slide: PowerPoint.Slide,
  index: number,
  shapes: ShapeInfo[],
): SlideOutline {
  const title = slideTitle(shapes);
  return {
    id: slide.id,
    index,
    layout: slide.layout.name,
    ...(title && { title }),
    shapeCount: shapes.length,
    fingerprint: fingerprintShapes(shapes),
  };
}

/** Outlines (with fingerprints) of the given slides, plus their summary shape trees. */
export async function readOutlines(
  context: PowerPoint.RequestContext,
  slides: PowerPoint.Slide[],
  indexOf: (slide: PowerPoint.Slide) => number,
): Promise<{ outlines: SlideOutline[]; shapes: ShapeInfo[][] }> {
  const shapes = await readShapeLists(
    context,
    slides.map((slide) => slide.shapes),
    "summary",
  );
  const outlines = slides.map((slide, i) => toOutline(slide, indexOf(slide), shapes[i] ?? []));
  return { outlines, shapes };
}

async function readDeck(
  context: PowerPoint.RequestContext,
): Promise<{ deck: DeckInfo; shapes: ShapeInfo[][] }> {
  const slides = await loadSlides(context);
  const pageSetup = supportsApi("1.10")
    ? context.presentation.pageSetup.load("slideWidth,slideHeight")
    : undefined;
  const { outlines, shapes } = await readOutlines(context, slides, (slide) =>
    slides.indexOf(slide),
  );
  const deck: DeckInfo = { slides: outlines };
  if (pageSetup) {
    deck.slideWidth = pageSetup.slideWidth;
    deck.slideHeight = pageSetup.slideHeight;
  }
  return { deck, shapes };
}

async function readSelection(context: PowerPoint.RequestContext): Promise<SelectionInfo> {
  const slides = context.presentation.getSelectedSlides().load("items/id");
  const shapes = context.presentation.getSelectedShapes().load("items/id");
  const range = context.presentation.getSelectedTextRangeOrNullObject().load("text,start,length");
  await context.sync();

  const selection: SelectionInfo = {
    slideIds: slides.items.map((slide) => slide.id),
    shapeIds: shapes.items.map((shape) => shape.id),
  };
  if (range.isNullObject || !range.text) return selection;
  selection.text = { text: range.text, start: range.start, length: range.length };
  try {
    const parent = range.getParentTextFrame().getParentShape().load("id");
    await context.sync();
    selection.text.shapeId = parent.id;
  } catch {
    // Text selected inside a table cell has no parent shape frame; the selection is still useful without it.
  }
  return selection;
}

async function readTheme(context: PowerPoint.RequestContext): Promise<ThemeInfo> {
  const master = context.presentation.slideMasters.getItemAt(0);
  const colorResults = supportsApi("1.10")
    ? THEME_COLORS.map((name) => [name, master.themeColorScheme.getThemeColor(name)] as const)
    : [];
  const [masterShapes = []] = await readShapeLists(context, [master.shapes], "styles");

  const placeholderFont = (types: Set<string>) =>
    masterShapes.find(
      (shape) => shape.placeholder && types.has(shape.placeholder) && shape.font?.name,
    )?.font?.name ?? undefined;
  const fonts = {
    heading: placeholderFont(TITLE_PLACEHOLDERS),
    body: placeholderFont(BODY_PLACEHOLDERS),
  };

  if (colorResults.length > 0) {
    return {
      colors: Object.fromEntries(colorResults.map(([name, result]) => [name, result.value])),
      sampled: false,
      fonts,
    };
  }
  return { colors: sampleColors(masterShapes), sampled: true, fonts };
}

/** Fallback when the theme API is missing: text and fill colors used on the master, labeled by where they appear. */
function sampleColors(shapes: ShapeInfo[]): Record<string, string> {
  const colors: Record<string, string> = {};
  for (const shape of shapes) {
    const role = shape.placeholder ?? shape.name;
    if (shape.font?.color) colors[`${role} text`] = shape.font.color;
    if (shape.fill?.color) colors[`${role} fill`] = shape.fill.color;
  }
  return colors;
}

// ---------------------------------------------------------------------------
// Ops
// ---------------------------------------------------------------------------

export const readOps = {
  get_deck: (): Promise<DeckInfo> =>
    PowerPoint.run(async (context) => (await readDeck(context)).deck),

  get_theme: (): Promise<ThemeInfo> => PowerPoint.run(readTheme),

  get_layouts: (): Promise<LayoutsInfo> =>
    PowerPoint.run(async (context) => {
      const masters = context.presentation.slideMasters.load(
        "items/id,items/name,items/layouts/items/id,items/layouts/items/name",
      );
      await context.sync();
      return {
        masters: masters.items.map((master) => ({
          id: master.id,
          name: master.name,
          layouts: master.layouts.items.map((layout) => ({ id: layout.id, name: layout.name })),
        })),
      };
    }),

  get_slide: ({ slideId }: { slideId: string }): Promise<SlideDetail> =>
    PowerPoint.run(async (context) => {
      const { slide, index } = await findSlide(context, slideId);
      const [shapes = []] = await readShapeLists(context, [slide.shapes], "paragraphs");
      return {
        id: slide.id,
        index,
        layout: slide.layout.name,
        fingerprint: fingerprintShapes(shapes),
        shapes,
      };
    }),

  get_selection: (): Promise<SelectionInfo> => PowerPoint.run(readSelection),

  get_deck_state: (): Promise<DeckState> =>
    PowerPoint.run(async (context) => {
      const { deck, shapes } = await readDeck(context);
      const selection = await readSelection(context);
      const theme = await readTheme(context);
      const selectedIndex = deck.slides.findIndex((slide) => slide.id === selection.slideIds[0]);
      const selectedShapes = shapes[selectedIndex];
      return { deck, selection, theme, ...(selectedShapes && { selectedShapes }) };
    }),

  /**
   * Index and fingerprint of each listed slide that exists, read in one run. Also exports the ones in exportSlideIds
   * when PowerPointApi 1.8 is available (undo snapshots, PPTX edits); without it they come back without base64.
   */
  get_slide_states: ({
    slideIds,
    exportSlideIds = [],
  }: {
    slideIds: string[];
    exportSlideIds?: string[];
  }): Promise<Record<string, SlideState>> =>
    PowerPoint.run(async (context) => {
      const exportIds = supportsApi("1.8") ? exportSlideIds : [];
      const all = await loadSlides(context);
      const slides = all.filter(
        (slide) => slideIds.includes(slide.id) || exportSlideIds.includes(slide.id),
      );
      const exports = new Map(
        slides
          .filter((slide) => exportIds.includes(slide.id))
          .map((slide) => [slide.id, slide.exportAsBase64()]),
      );
      const { outlines } = await readOutlines(context, slides, (slide) => all.indexOf(slide));
      return Object.fromEntries(
        outlines.map((outline) => {
          const base64 = exports.get(outline.id)?.value;
          return [
            outline.id,
            { index: outline.index, fingerprint: outline.fingerprint, ...(base64 && { base64 }) },
          ];
        }),
      );
    }),

  render_slide: ({
    slideId,
    width = MAX_RENDER_WIDTH,
  }: {
    slideId: string;
    width?: number;
  }): Promise<string> =>
    PowerPoint.run(async (context) => {
      requireApi("1.8", "Rendering slides");
      const { slide } = await findSlide(context, slideId);
      const image = slide.getImageAsBase64({ width: Math.min(width, MAX_RENDER_WIDTH) });
      await context.sync();
      return image.value;
    }),

  export_slide: ({ slideId }: { slideId: string }): Promise<string> =>
    PowerPoint.run(async (context) => {
      requireApi("1.8", "Exporting slides");
      const { slide } = await findSlide(context, slideId);
      const base64 = slide.exportAsBase64();
      await context.sync();
      return base64.value;
    }),
};
