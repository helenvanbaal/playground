# Self-critique rubric

Run this after the render-and-look pass in `SKILL.md`, once every slide
image has been viewed and obvious visual bugs are already fixed. This pass
is about judgment, not just defect-spotting.

## Score every slide 1–5 on each dimension

| Dimension | 1 (fail) | 3 (ok) | 5 (strong) |
|---|---|---|---|
| **Hierarchy** | No clear focal point; everything competes | One clear focal element, but supporting items are ambiguous | Eye goes exactly where it should, in order, in under 2 seconds |
| **Alignment** | Elements don't share edges or a grid | Mostly aligned, a few stray elements | Every element aligns to the layout's placeholders/grid |
| **Whitespace** | Cramped, text touches edges or other elements | Adequate but uneven margins | Generous, consistent margins; breathing room around the focal element |
| **Contrast** | Text hard to read against its background | Readable but not crisp (e.g. grey-on-white body text) | Strong, accessible contrast everywhere text appears |
| **Consistency** | Fonts/colors/positions vary slide-to-slide without reason | Mostly consistent, a couple of one-off choices | Title position, type scale, and palette match every other slide of the same layout |
| **Text density** | Wall of text, tiny font, or 6+ bullets | At the content-rules limit (≤5 bullets, ≤40 words) | Clearly within budget; a chart/visual carries what bullets would have |

## Threshold

Any dimension scoring **below 4** on a given slide must be fixed before
delivery — either by editing the outline/content, swapping the layout, or
making a targeted fix directly in the generated deck. Re-render and re-view
the corrected slide to confirm the fix actually worked; don't assume a code
change produced the intended visual result.

## Explicit check: avoid list

While scoring, flag any of these on sight — they almost always mean a
dimension above should be capped at 2 or lower:

- Default Office blue (`#4472C4`) or any other non-palette color
- Centered body text/bullets (this template left-aligns; centered body text
  reads as off-brand and harder to scan)
- Bullets on every single slide (if more than ~60% of content slides are
  plain bulleted lists, go back and apply "prefer a visual to a bullet
  list" from `content-rules.md`)
- A decorative line/rule under a title that isn't part of the actual
  template layout
- Text smaller than 14pt that you authored yourself (template-default
  narrow-column text is the only exception)
- Low-contrast text over a busy photo (no scrim/overlay, or wrong text
  color for the image)
- Inconsistent title position/size across slides that use the same layout
- Images stretched/distorted off their native aspect ratio
- Orphaned single words wrapping alone onto a final line
- Clipped or overflowing text (spilling past a placeholder's visible
  boundary, or truncated by autofit shrinking text below 14pt) — check
  `big-number-stat` slides specifically: a full sentence in the big-number
  slot overflows into the row below and looks like garbled overlapping text
- A chart with its own auto-generated title repeating the slide title
  almost verbatim — turn the chart's internal title off and let the slide
  title carry the message once
- A table header that's hard to read against its own fill (e.g. white text
  on a light fill) — set both the fill and the font color explicitly rather
  than trusting the default table style
- A slide missing its intended title entirely because the chosen layout
  has no title/headline/heading placeholder (e.g. a timeline built on
  "Empty page") — confirm the title actually rendered, not just that the
  code ran without error

## Builder bias warning

A model that built the deck tends to approve its own work. When the deck is
high-stakes (leadership review, external audience, final delivery), spawn a
**separate reviewer** — a fresh `task` agent that receives only the
rendered slide images and this rubric, not the build transcript or outline —
and treat its scores as the ones that count. Reconcile any disagreement by
re-inspecting the actual rendered image, not by trusting the builder's
original intent.

## Recording the pass

Report, per slide: the six scores, any avoid-list flags, and either "no
action" or the specific fix applied. Do not report a deck as complete if any
slide is still below the threshold on any dimension.
