# Good and bad examples

## Good: real slides from the brand template

These are rendered screenshots of actual slides from the source brand file,
chosen because they show the template's layouts used well. Compare any
generated slide against the one with the matching layout before delivering.

| Image | Layout | Why it works |
|---|---|---|
| `title-black.jpg` / `title-yellow.jpg` | `title` | Huge, confident display type; one line of supporting text; nothing else competing for attention |
| `agenda.png` | `section-divider` | Big heading, short left-aligned list, generous whitespace — a navigation aid, not a content slide |
| `big-statement.jpg` | `big-statement` | One sentence, large type, grounded in a real photo — this is what "the title is the takeaway" looks like at full size |
| `statement-highlights.jpg` | `big-number-stat` (photo mood) | Photo carries the emotional weight; three short captions add just enough specificity |
| `dual-image-statement.jpg` | `comparison` (dual-image) | Two real photos, two-word captions each — a comparison without a single bullet |
| `timeline-simple.png` | `timeline` | Horizontal line, sized circles, alternating labels above/below so nothing collides |
| `end-slide-green.png` / `end-slide-black.png` | `closing` | Minimal, on-brand, matches the opening mood |
| `brand-colors.png` | — | The literal brand swatches; cross-check any custom color against this before using it |

Note: `big-statement.jpg`, `dual-image-statement.jpg`, and
`statement-highlights.jpg` are recompressed JPEGs (originals were
multi-megabyte PNGs of full photos) — this is fine for visual reference, and
is also a good model for keeping the skill package itself small.

## Bad: a deliberately broken slide (`bad-example.png`)

This slide was built outside the brand template to concentrate nearly every
"ugly slide" failure mode into one image for calibration. If a generated
slide resembles this in any respect, fix it before delivering.

What's wrong, specifically:

1. **Default Office blue background** (`#1F497D`) — not a brand color. Compare
   against `brand-colors.png`: the only blue in this system is `#0058A3`.
2. **Centered title, centered body text** — this template is left-aligned
   throughout. Centered body text is an instant off-brand signal.
2. **A decorative yellow line under the title** — the real template never
   underlines titles with a rule. If a slide "needs" a divider line to feel
   finished, that's a sign the layout/hierarchy is doing too little work.
3. **Eight bullets, single slide** — blows through the 5-bullet limit by 60%.
   This is a topic dump, not one idea.
4. **13pt body text** — below the 14pt floor, and still wraps to full-width
   lines because there are far too many words for the slide.
5. **Low-contrast text** (`#9DB2D4` on `#1F497D`) — technically legible,
   genuinely tiring to read; nothing like the strong white/black contrast in
   every real template example.
6. **Calibri**, not Noto IKEA Latin — a font not in the two-font system.
7. **The title is a topic, not a takeaway** — "Q3 Results" says nothing a
   reader couldn't guess from the agenda. Every one of the eight bullets
   could instead be a `big-number-stat`, a `three-or-four-column` pillars
   slide, or a native chart.

This single slide violates design-system.md (color, font, contrast),
layouts.md (no named layout looks like this), content-rules.md (word/bullet
limits, title-as-takeaway), and five separate rubric.md avoid-list items —
which is exactly why the render-and-look pass and rubric scoring catch it
immediately once rendered, even though the underlying python-pptx code runs
without any error.
