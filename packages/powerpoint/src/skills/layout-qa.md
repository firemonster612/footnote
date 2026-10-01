# Layout QA

Use this after creating or changing slides, or when the user asks you to check or clean up a deck. Shape data tells you where things are; only a render shows what the audience sees.

## The loop

1. `render_slide` for each slide you changed.
2. Look for the problems below. Use `get_slide` to get exact bounds and font sizes for anything that looks off.
3. Fix everything you found with one batched `update_shapes` call covering all the slides.
4. Render again. Repeat until clean, but stop after three rounds on one slide and tell the user what's still wrong rather than looping.

Check every changed slide, not just the first. Problems cluster on the slides with the most content.

## What to look for

**Overflow and clipping**

- Text running past the bottom of its box or off the slide.
- Text auto-shrunk far below the body size of other slides.
- Fix: cut words first, then widen or heighten the box within the grid, then split the slide. Shrinking the font is the last resort and never below 14 pt for body text.

**Alignment**

- Edges that are almost but not exactly aligned (a few points off).
- Repeated items (cards, columns, icons) with unequal sizes or gaps.
- Fix: set `left`, `top`, `width`, `height` to shared values. Compute equal gaps: gap = (available width − total item width) / (count − 1).

**Margins and spacing**

- Content closer to the slide edge than the deck's margin (about 48 pt on a 960 pt wide slide, or whatever the other slides use).
- Blocks touching or crowding each other; aim for 16–24 pt between separate blocks.
- Large empty areas next to cramped ones: rebalance.

**Contrast and legibility**

- Light text on light backgrounds, dark on dark, or text over busy image areas.
- Fix: use the theme's text color for the background, add a solid shape behind the text, or move the text.

**Overlap**

- Shapes covering each other unintentionally, text over an image edge, a chart legend over data.
- Fix: move or resize; check z-order if something is hidden.

**Consistency with the deck**

- Title position, size, and font matching other slides.
- Same bullet style, same accent color, same footer and slide-number placement.

**Leftovers**

- Empty placeholders ("Click to add text"), placeholder text like "Lorem ipsum" or "Title", stray empty text boxes.
- Fix: delete them.

## Measuring instead of guessing

- All bounds are points. The slide size comes from `get_deck`.
- A shape's right edge is `left + width`; bottom is `top + height`. A shape fits the slide if right ≤ slide width − margin and bottom ≤ slide height − margin.
- To center a shape horizontally: left = (slide width − width) / 2.
- When many shapes need the same treatment across slides, compute values once and apply them identically.

## Reporting

Say what you checked and what you fixed in a sentence or two per slide that needed work. If a slide still has a problem you couldn't fix cleanly (for example, the content needs cutting and that's the user's call), say so plainly.
