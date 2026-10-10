# Content rules

Most "ugly" slides are actually "too much text" or "the wrong content type
for the point being made." Apply these rules while planning the outline,
before any layout or visual work starts.

## One idea per slide

Every slide carries exactly one idea. If you can't state the slide's point
in one sentence, it's two slides.

## The title is the takeaway, not the topic

Write titles as the conclusion, not the subject:

- Bad: "Q3 Results"
- Good: "Revenue grew 40% in Q3, led by EMEA"
- Bad: "Customer Feedback"
- Good: "Customers want faster checkout, not more features"

Exceptions: `title` slides (the deck's own title) and `section-divider`
slides (agenda/section names) are labels by nature — don't force them into
takeaway sentences.

## Hard limits

- Body text: **≤ 30–40 words** per slide outside the title.
- Bullets: **≤ 5**, each **≤ 10 words** where possible. If a bullet needs two
  lines to say something simple, cut it.
- Font size: **never below 14pt** for text you author yourself. (Template
  placeholders in narrow 3–4 column layouts default to 11–12pt for short
  labels — that's a template exception, not a license to shrink paragraphs.)
- If content exceeds these limits: split into two slides, or convert to a
  chart/diagram/table that carries more information without more words.

## Prefer a visual to a bullet list

For every slide, ask: **can this be a chart, diagram, icon row, big number,
or table instead of bullets?** Default answer should be yes more often than
not. Concretely:

| If the content is... | Use... |
|---|---|
| A trend over time | A native line/bar chart (`add_chart`), not a bulleted list of numbers |
| A comparison across 2+ named items/criteria | A table (`add_table`) or `comparison` layout, not parallel bullet lists |
| One number that matters | `big-number-stat` layout, not a sentence burying the number |
| A sequence of steps/dates | `timeline` layout, not a bulleted chronology |
| 3–4 parallel short items | `three-or-four-column`, not a single bulleted list |
| A claim about the real world | `image-and-text` with a real photo, not a text-only slide |
| Genuinely sequential prose/reasoning | Bullets are fine — don't force a chart where there isn't one |

Native charts only: build charts with `python-pptx`'s chart API
(`add_chart` in `scripts/generate_deck.py`), not pasted screenshot images of
charts from Excel/other tools. Native charts stay editable, match the brand
palette, and render crisply at any size. Always set chart colors from the
palette in `references/design-system.md` — never leave default Office chart
colors.

## Speaker notes carry the detail

Anything that matters but would crowd the slide goes in speaker notes: data
sources, caveats, methodology, the detail behind a headline number. The
slide shows the takeaway; the notes carry the proof.

## Do not invent facts

Mark uncertain or missing content as `[to confirm]` rather than inventing a
number, quote, or source. This applies to chart data too — do not fabricate
data points to make a chart look complete.
