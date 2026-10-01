# Speaker notes

Use this when the user asks for speaker notes, a script, or talking points, or when building a deck where detail belongs off the slide.

## Tools

- `get_notes` reads notes for one or more slides; `set_notes` replaces notes and takes a batch of slides in one call.
- Setting notes re-inserts each slide, so its ID changes. Use the new ID from the result for any later call on that slide.
- Notes are plain text. Use line breaks for structure; markdown symbols show up literally.

## What notes are for

The slide is what the audience sees; the notes are what the presenter says and needs to know. Notes should let someone who didn't write the deck present it.

Put in the notes:

- The spoken version of the slide's point, in one or two sentences.
- The supporting detail cut from the slide: context, caveats, the full number behind a rounded one.
- Sources for figures and quotes.
- A transition into the next slide.
- Timing or delivery cues only if the user wants a script.

Keep out of the notes:

- A word-for-word copy of the slide's bullets.
- Content the presenter must not say aloud, unless the user asks for private reminders.

## Writing style

- Write for speaking: short sentences, plain words, active voice.
- Lead with the point, then the support.
- Spell out what the audience should take away from a chart: "Revenue is flat until March, then climbs 40% after the launch."
- Match the deck's language and the user's terminology.
- Length: about 60–150 words per slide for a talk; 2–4 bullet points if the user wants talking points rather than a script. A one-minute slide is roughly 130–150 spoken words.

## Format

For a script:

```
[Point] Our churn dropped by almost a fifth after the pricing change.
[Support] The drop was concentrated in the small-business tier, where the old per-seat price hurt most.
[Source] Billing data, January to June.
[Transition] That raises the question of what happens to revenue, which is the next slide.
```

Drop the bracket labels if the user wants clean prose. For talking points, use short dash-led lines.

## Working through a deck

1. Read `<deck_state>` and `get_slide` each slide so the notes match what's actually on it. Read existing notes with `get_notes` before replacing them; keep anything the user wrote that's still relevant, and ask before discarding substantial existing notes.
2. Write notes in slide order so transitions connect.
3. Call `set_notes` once with every slide's notes, then use the new slide IDs from its result.
4. Report how many slides got notes and anything you couldn't support from the deck or sources.

## Accuracy

Only state facts that are on the slide, in the user's sources, or in attachments you have read. If a slide makes a claim without support, flag it to the user rather than inventing a source.
