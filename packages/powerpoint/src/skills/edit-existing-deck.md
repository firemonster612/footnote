# Edit an existing deck

Use this when the user wants changes to slides that already exist: rewording, restyling, restructuring, or adding slides that must match the rest.

## Read before you write

- Start from `<deck_state>`: slide IDs, titles, layouts, and the user's selection. "This slide" or "these" means the selection.
- Call `get_slide` on every slide you will change. Write tools refuse to act on a slide that changed since you last read it; if that happens, read it again and redo your edit against the new state.
- Call `render_slide` on a representative slide before restyling anything, so you see what the template actually looks like.

## Respect the template

The deck's existing choices are the spec, even where they differ from general design advice.

- Keep the theme fonts, colors, and slide masters. Don't swap fonts or recolor unless asked.
- Match existing slides: title position and size, body font size, bullet style, margins, accent usage. Measure them with `get_slide` (bounds are in points) and reuse the numbers.
- New slides use the same layouts the deck already uses for similar content. Use `duplicate_slide` on a well-formed slide when no layout fits, then replace its content.
- Keep logos, footers, slide numbers, and confidentiality marks where the template puts them.

## Make minimal diffs

- Change only what the request covers. Rewording a title doesn't license restyling the slide.
- Edit text in place with `update_shapes` instead of deleting and recreating shapes; recreated shapes lose animations, links, and their IDs.
- Keep run formatting when rewriting text: if a phrase was bold or colored, carry that formatting over to the new wording where it still applies.
- Don't reorder, delete, or merge slides unless the user asked. When a request implies it ("tighten this deck"), say what you plan to cut before cutting.

## Consistency passes

For deck-wide requests ("make the titles consistent", "fix the fonts"), work in two steps:

1. Survey: read every affected slide with `get_slide`. Note the variants you find (title sizes, fonts, bullet styles, capitalization, punctuation).
2. Pick the majority or template value as the standard, tell the user what you chose, then apply it slide by slide with one batched `update_shapes` call per slide.

Common consistency fixes:

- Title capitalization (sentence case or title case, one style).
- Ending punctuation on bullets (all or none).
- Font sizes drifting by a point or two.
- Content shifted a few points off the shared margins.
- Number formats (12% vs 12 percent, $1.2M vs $1,200,000).

## Rewriting text

- Keep the author's voice and terminology. Shorten rather than paraphrase.
- Keep every number, name, and claim unless the user asked you to change it.
- If shorter text no longer fills the placeholder well, leave the layout alone unless it looks broken.

## After editing

- Render each changed slide and compare against the slides around it. Load `layout-qa` for the checklist.
- Report which slides changed and what you changed, briefly. Mention anything you noticed but left alone.
- The user can undo the whole turn; still, don't make sweeping changes they didn't ask for.
