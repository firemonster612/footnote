# Deck design

A deck reads as designed when every slide follows the same few rules. Decide the rules once, then apply them everywhere.

## Start from what the deck already has

- Call `get_deck` and read the theme fonts, theme colors, slide size, and layouts before choosing anything.
- Use theme fonts and theme colors. Introduce a new font or color only when the user asks or the deck has no theme worth keeping.
- Prefer a layout with placeholders over free-floating text boxes. Placeholders keep position, size, and type consistent and survive theme changes.

## Grid and margins

Slide sizes are in points. A 16:9 deck is usually 960 × 540; 4:3 is 720 × 540. Read the real size from `get_deck`.

- Margins: about 5% of the slide width on the left and right (48 pt on 960 wide) and a similar top margin. Keep them identical on every slide.
- Use a 12-column grid inside the margins for placing content. Common splits: full width, 6/6, 4/8, 4/4/4.
- Align edges to shared lines. Two boxes that are almost aligned look worse than two boxes that are obviously not.
- Leave at least 16–24 pt between separate blocks.

## Type scale

Pick four sizes and use only those:

| Role                         | 16:9 at 960 pt wide |
| ---------------------------- | ------------------- |
| Slide title                  | 28–36 pt            |
| Subtitle or section lead     | 20–24 pt            |
| Body                         | 16–20 pt            |
| Captions, sources, footnotes | 10–12 pt            |

- Body text below 14 pt is hard to read when projected. If text only fits at 12 pt, cut words or split the slide.
- Left-align body text. Center only short titles or single statements.
- One font family for everything is fine. Two at most: one for headings, one for body.
- Bold for emphasis, sparingly. Avoid underline (it reads as a link) and all caps for more than a few words.

## Color

- Theme colors: dark text on a light background, or the reverse. Keep one background treatment across the deck.
- Choose one accent color for emphasis and highlights. Use a second accent only to separate data series.
- Never encode meaning in color alone; pair it with a label, position, or shape.
- Contrast: body text needs at least 4.5:1 against its background. Light gray on white and mid-tone text on photos usually fail.

## One idea per slide

- The title states the point of the slide as a sentence ("Churn fell 18% after the pricing change"), not a topic ("Churn").
- Each slide supports that one point. If you need "and" in the title, it's probably two slides.
- Bullets: at most 5–6 per slide, one line each where possible, no more than two levels.
- Replace a bullet list with a diagram, chart, or table when the content has structure (steps, comparison, numbers).

## Images and shapes

- Images fill a grid area or bleed to the slide edge; avoid small images floating in empty space.
- Keep aspect ratios; never stretch.
- Use simple shapes (rectangles, lines, circles) in theme colors. Skip shadows, bevels, gradients, and 3D effects unless the template already uses them.
- Repeated elements (cards, icons, numbered steps) share size, spacing, and alignment exactly. Set bounds with numbers, not by eye.

## Before you finish

Render each slide you changed with `render_slide` and check it against these rules. The `layout-qa` skill has the full checklist.
