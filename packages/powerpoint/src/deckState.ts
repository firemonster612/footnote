// Builds the per-turn <deck_state> block. Pure: takes op results, returns text.
import type { DeckState, ShapeInfo, SlideOutline, ThemeInfo } from "./ops/types.ts";

const MAX_CHARS = 8_000; // ≈2k tokens
const FULL_OUTLINE_LIMIT = 40;
const MAX_SELECTED_SHAPES = 15;

export interface DeckChanges {
  added: SlideOutline[];
  removed: string[];
  moved: SlideOutline[];
  modified: SlideOutline[];
}

/** Slide-level changes between two outlines. Moves are the minimal set: slides outside the longest run that kept its relative order. */
export function diffDecks(previous: SlideOutline[], current: SlideOutline[]): DeckChanges {
  const before = new Map(previous.map((slide, order) => [slide.id, { slide, order }]));
  const currentIds = new Set(current.map((slide) => slide.id));
  const common = current.filter((slide) => before.has(slide.id));
  const kept = longestIncreasingRun(common.map((slide) => before.get(slide.id)?.order ?? 0));
  return {
    added: current.filter((slide) => !before.has(slide.id)),
    removed: previous.filter((slide) => !currentIds.has(slide.id)).map((slide) => slide.id),
    moved: common.filter((_, i) => !kept.has(i)),
    modified: common.filter(
      (slide) => before.get(slide.id)?.slide.fingerprint !== slide.fingerprint,
    ),
  };
}

/** Indexes (into `values`) of one longest strictly increasing subsequence. */
function longestIncreasingRun(values: number[]): Set<number> {
  const tails: number[] = []; // tails[k] = index of the smallest tail of an increasing run of length k+1
  const parent: number[] = [];
  values.forEach((value, i) => {
    let low = 0;
    let high = tails.length;
    while (low < high) {
      const mid = (low + high) >> 1;
      if ((values[tails[mid] ?? 0] ?? 0) < value) low = mid + 1;
      else high = mid;
    }
    parent[i] = low > 0 ? (tails[low - 1] ?? -1) : -1;
    tails[low] = i;
  });
  const run = new Set<number>();
  for (let i = tails[tails.length - 1] ?? -1; i >= 0; i = parent[i] ?? -1) run.add(i);
  return run;
}

export function hasChanges(changes: DeckChanges): boolean {
  return Object.values(changes).some((list) => list.length > 0);
}

/** Deck text is untrusted: quote it and keep it from closing our tags. */
export function quote(text: string, max: number): string {
  const oneLine = text.replace(/\s+/g, " ").trim();
  const clipped = oneLine.length > max ? `${oneLine.slice(0, max)}…` : oneLine;
  return JSON.stringify(clipped).replaceAll("<", "\\u003c");
}

const slideRef = (slide: SlideOutline): string => `${slide.index + 1} (${slide.id})`;

function outlineLine(slide: SlideOutline): string {
  const title = slide.title ? ` · ${quote(slide.title, 70)}` : "";
  return `${slide.index + 1}. ${slide.id} · ${slide.layout}${title}`;
}

/** Full outline for small decks; otherwise the ends, the neighborhood of each focus slide, and gap markers. */
function outlineLines(slides: SlideOutline[], focusIds: Set<string>): string[] {
  if (slides.length <= FULL_OUTLINE_LIMIT) return slides.map(outlineLine);
  const keep = new Set<number>();
  const keepRange = (from: number, to: number) => {
    for (let i = Math.max(0, from); i <= Math.min(slides.length - 1, to); i++) keep.add(i);
  };
  keepRange(0, 9);
  keepRange(slides.length - 5, slides.length - 1);
  slides.forEach((slide, i) => {
    if (focusIds.has(slide.id)) keepRange(i - 3, i + 3);
  });

  const lines: string[] = [];
  let skipped = 0;
  slides.forEach((slide, i) => {
    if (!keep.has(i)) {
      skipped++;
      return;
    }
    if (skipped) lines.push(`… ${skipped} slides (call get_deck for all) …`);
    skipped = 0;
    lines.push(outlineLine(slide));
  });
  return lines;
}

function shapeLine(shape: ShapeInfo): string {
  const kind = shape.placeholder ? `${shape.placeholder} placeholder` : shape.type;
  const bounds = `${shape.left},${shape.top} ${shape.width}×${shape.height}`;
  const text = shape.text?.trim() ? ` · ${quote(shape.text, 60)}` : "";
  return `- ${shape.id} · ${quote(shape.name, 40)} · ${kind} · ${bounds}${text}`;
}

function themeLine(theme: ThemeInfo): string {
  const fonts = [
    theme.fonts.heading && `heading ${theme.fonts.heading}`,
    theme.fonts.body && `body ${theme.fonts.body}`,
  ]
    .filter(Boolean)
    .join(", ");
  const colors = Object.entries(theme.colors)
    .map(([name, color]) => `${name} ${color}`)
    .join(", ");
  return `Theme: fonts ${fonts || "unknown"}; ${theme.sampled ? "colors sampled from the master" : "colors"}: ${colors || "unknown"}`;
}

function changeLines(changes: DeckChanges): string[] {
  const parts = [
    changes.added.length > 0 && `added ${changes.added.map(slideRef).join(", ")}`,
    changes.removed.length > 0 && `removed ${changes.removed.join(", ")}`,
    changes.moved.length > 0 &&
      `moved ${changes.moved.map((slide) => `${slide.id} → ${slide.index + 1}`).join(", ")}`,
    changes.modified.length > 0 && `modified ${changes.modified.map(slideRef).join(", ")}`,
  ].filter(Boolean);
  return [`Changes since last update: ${parts.length > 0 ? parts.join("; ") : "none"}`];
}

export interface DeckStateInput {
  state: DeckState;
  documentName?: string;
  /** Omitted on the first block of a chat. */
  changes?: DeckChanges;
  includeTheme: boolean;
  /** Slides that changed after the model last read them. */
  staleSlideIds: string[];
}

export function formatDeckState({
  state,
  documentName,
  changes,
  includeTheme,
  staleSlideIds,
}: DeckStateInput): string {
  const { deck, selection } = state;
  const size =
    deck.slideWidth && deck.slideHeight ? ` · ${deck.slideWidth}×${deck.slideHeight} pt` : "";
  const name = documentName ? `${quote(documentName, 80)} · ` : "";
  const byId = new Map(deck.slides.map((slide) => [slide.id, slide]));
  const focus = new Set([
    ...selection.slideIds,
    ...staleSlideIds,
    ...(changes ? [...changes.added, ...changes.modified].map((slide) => slide.id) : []),
  ]);

  const lines = [
    `Deck: ${name}${deck.slides.length} slides${size}`,
    "Slides (position. ID · layout · title):",
    ...outlineLines(deck.slides, focus),
  ];

  const selectedSlides = selection.slideIds.flatMap((id) => byId.get(id) ?? []);
  if (selectedSlides.length > 0) {
    const shapes = selection.shapeIds.length > 0 ? `; shapes ${selection.shapeIds.join(", ")}` : "";
    const text = selection.text
      ? `; text ${quote(selection.text.text, 80)} in shape ${selection.text.shapeId}`
      : "";
    lines.push(`Selection: slide ${selectedSlides.map(slideRef).join(", ")}${shapes}${text}`);
  }
  if (state.selectedShapes && selectedSlides[0]) {
    lines.push(
      `Shapes on slide ${selectedSlides[0].index + 1} (ID · name · type · left,top width×height · text):`,
    );
    lines.push(...state.selectedShapes.slice(0, MAX_SELECTED_SHAPES).map(shapeLine));
    const more = state.selectedShapes.length - MAX_SELECTED_SHAPES;
    if (more > 0) lines.push(`- … ${more} more (get_slide)`);
  }
  if (includeTheme) lines.push(themeLine(state.theme));
  if (changes) lines.push(...changeLines(changes));
  const stale = staleSlideIds.flatMap((id) => byId.get(id) ?? []);
  if (stale.length > 0) {
    lines.push(
      `Changed since you last read them (call get_slide before editing): ${stale.map(slideRef).join(", ")}`,
    );
  }

  let body = lines.join("\n");
  if (body.length > MAX_CHARS) body = `${body.slice(0, MAX_CHARS)}\n… truncated (call get_deck)`;
  return `<deck_state>\n${body}\n</deck_state>`;
}
