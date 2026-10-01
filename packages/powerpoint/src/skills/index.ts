import type { SkillDefinition } from "@footnote/core/contracts";
import buildDeckFromSource from "./build-deck-from-source.md?raw";
import chartsAndTables from "./charts-and-tables.md?raw";
import deckDesign from "./deck-design.md?raw";
import editExistingDeck from "./edit-existing-deck.md?raw";
import layoutQa from "./layout-qa.md?raw";
import speakerNotes from "./speaker-notes.md?raw";

const bundled = (name: string, description: string, body: string): SkillDefinition => ({
  name,
  description,
  body,
  source: "bundled",
});

export const powerpointSkills: SkillDefinition[] = [
  bundled(
    "deck-design",
    "Visual rules for slides: grid, margins, type scale, theme colors, contrast, one idea per slide. Use when creating or restyling slides.",
    deckDesign,
  ),
  bundled(
    "build-deck-from-source",
    "Turn an attachment, notes, web pages, or an outline into a deck: storyline first, then layouts, placeholders, and a render check.",
    buildDeckFromSource,
  ),
  bundled(
    "edit-existing-deck",
    "Change slides that already exist: respect the template, keep diffs minimal, run consistency passes across slides.",
    editExistingDeck,
  ),
  bundled(
    "charts-and-tables",
    "Presenting numbers: chart vs table, choosing a chart type, labeling, and using insert_chart and edit_table.",
    chartsAndTables,
  ),
  bundled(
    "layout-qa",
    "Render-and-fix loop for overflow, alignment, margins, contrast, overlap, and leftover placeholders. Use after changing slides.",
    layoutQa,
  ),
  bundled(
    "speaker-notes",
    "Writing speaker notes, scripts, or talking points with get_notes and set_notes.",
    speakerNotes,
  ),
];
