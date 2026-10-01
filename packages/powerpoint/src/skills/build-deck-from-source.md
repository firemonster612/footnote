# Build a deck from source material

Use this when the user hands you a document, notes, data, or an outline and wants slides made from it.

## 1. Read the source

- For attachments, the message holds an excerpt. Call `read_attachment` with `offset` until you have read everything that matters. For PDFs, pass `pages` to look at charts and figures the text extraction misses.
- For web sources, use `fetch_page`. Record the URL so you can cite it on the slide or in the notes.
- Note the audience, purpose, and length if the user gave them. If length isn't given, aim for one slide per main point plus a title and a closing slide; 8–12 slides suits most source documents.

## 2. Write the storyline before touching slides

Draft the deck as a list of slide titles in your reply or your reasoning. Each title is a full sentence carrying that slide's point. Read the titles in order: they should tell the whole story on their own.

A structure that works for most business material:

1. Title slide: topic, presenter or organization, date.
2. The answer or recommendation, up front.
3. Context: the situation and why it matters now.
4. Evidence: one slide per supporting point, each with its key number, chart, or example.
5. Implications, options, or plan.
6. Next steps or the ask.

Show the user the outline first when the request is ambiguous or the deck is long. For short, clear requests, go ahead.

## 3. Choose layouts

- Call `get_deck` for the layouts and theme. Map each slide in the outline to a layout: Title Slide, Title and Content, Two Content, Section Header, Title Only (for charts and diagrams), Blank (rarely).
- If the deck already has slides, match their style. Load the `edit-existing-deck` skill.
- Load `deck-design` for spacing, type, and color rules if you will place anything outside placeholders.

## 4. Build

- Create slides with `add_slide` and the chosen layout, in order. Fill placeholders with `update_shapes` (set text on the title and body placeholders by shape ID from `get_slide`).
- Batch: fill the new slides with one `update_shapes` call covering every shape on all of them.
- Body text: short bullets or a few sentences. Move detail, caveats, and sources to speaker notes with `set_notes` (see `speaker-notes`).
- Numbers: use `insert_chart` for trends and comparisons, `edit_table` for exact values people will look up. Load `charts-and-tables`.
- Images from the user's attachments: `add_shape` with kind image and the attachment ID.
- Delete placeholders you don't use. Empty placeholders show "Click to add text" in edit mode.
- If the user attached a .pptx with slides to reuse, `insert_slides_from_file` copies them in.

## 5. Check

Render every new slide with `render_slide`. Fix overflow, crowding, and misalignment before reporting back. The `layout-qa` skill has the checklist and fix loop.

## 6. Report

Tell the user what you built in a few lines: slide count, the storyline in brief, and anything you left out or assumed. Don't restate every slide.

## Faithfulness

- Use the source's numbers and claims exactly. Don't round, extrapolate, or add figures the source doesn't contain.
- When you summarize an argument, keep its qualifiers ("in the pilot", "estimated").
- If the source contradicts itself or lacks something the deck needs, say so instead of inventing it.
