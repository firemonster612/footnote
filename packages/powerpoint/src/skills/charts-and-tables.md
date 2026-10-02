# Charts and tables

Use this when slides need numbers: choosing between a chart and a table, picking the chart type, and inserting it with `insert_chart` or `edit_table`.

## Chart or table?

- **Chart** when the point is a shape: a trend, a ranking, a share, a gap between two things. The audience should get it in five seconds.
- **Table** when people need exact values, compare many attributes per item, or will look things up. Keep it under about 6 columns and 8 rows on a slide.
- **One big number** when a single figure carries the point. Large text plus a short label beats a chart with one bar.
- Put the takeaway in the slide title either way ("Q3 revenue grew 14%, the fastest in two years").

## Choosing a chart type

| Message                                 | Chart                                    |
| --------------------------------------- | ---------------------------------------- |
| Change over time                        | line (many points), column (few periods) |
| Ranking or comparison across categories | bar (horizontal), sorted by value        |
| Part of a whole, 2–5 parts              | pie or doughnut                          |
| Part of a whole, more than 5 parts      | bar (horizontal), sorted by value        |
| Relationship between two measures       | scatter                                  |
| Cumulative volume over time             | area                                     |

Avoid pie charts with more than five slices or slices that are close in size; use a sorted bar chart instead. Don't use 3D.

## Inserting a chart

`insert_chart` builds a native, editable PowerPoint chart.

- To put a chart on an existing slide, pass its `slideId`. The slide is re-inserted, so it gets a **new ID**; use the ID from the result in later calls.
- For a new chart slide that matches the deck, `add_slide` with a Title Only layout, set the title, then `insert_chart` with that slide's ID and bounds below the title.
- Bounds are in points. Leave the slide margins clear and keep space under the title (for 960 × 540: roughly left 48, top 110, width 864, height 390).
- Pass theme colors (hex from `get_deck`) for series, and the theme body font, so the chart matches the deck.
- To change a chart's data later, insert a new one and delete the old chart shape with `update_shapes`.

## Making charts readable

- Label directly: data labels on bars or the last point of a line instead of a legend when there are few series. Turn the legend off when there is one series.
- Highlight the point: accent color on the series or bar that matters, gray for the rest.
- Number format matches the data: `0%`, `$#,##0`, `#,##0.0`. Include units in the axis title or chart title, not on every label.
- Start bar and column value axes at zero.
- Sort categories by value unless they have a natural order (time, stages).
- Keep chart text at least 12 pt.

## Tables with edit_table

- Create with a header row and the values in one call. Use the deck's table style if other tables exist; otherwise a light style such as `LightStyle1Accent1`.
- Right-align numbers and use the same decimals within a column. Left-align text.
- Bold the header row; avoid heavy borders and per-cell fills. Highlight at most one row or column with the accent color.
- Set column widths so long text doesn't wrap to many lines; abbreviate headers if needed.
- If a table runs past about 8 rows, split it, cut rows to the ones that support the point, or move the full table to an appendix slide.

## Check

Render the slide. Confirm labels don't overlap, nothing is clipped, the legend (if any) doesn't cover data, and the numbers match the source exactly.
