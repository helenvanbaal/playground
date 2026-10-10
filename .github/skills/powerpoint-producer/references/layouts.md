# Layout library

Choose exactly one named layout per slide from this menu. Each maps to a real
layout in `assets/base-template.pptx` (inspect with
`python3 scripts/generate_deck.py --inspect-template` to confirm current
placeholder `idx` values before relying on them in code). Do not invent a
layout outside this menu, and do not build a slide on a blank canvas when a
matching named layout exists.

For every layout, "lead" means a one-sentence takeaway, not a topic label.

## 1. `title`

**Template layout:** `Title - Black` / `Title - White` / `Title - Yellow`
(pick the mood that fits the deck's tone; black = serious/default, yellow =
energetic/kickoff, white = light/neutral).

**Placeholders:** `idx=0` title (big display, autosized), `idx=10` body
(subtitle/presenter line/date, keep to one short line).

**Use for:** the opening slide, and optionally a closing section title
restated in the deck's mood color.

**Budget:** title ≤ 8 words. Subtitle ≤ 12 words (e.g. "Prepared for
Leadership Review — October 2026").

## 2. `section-divider`

**Template layout:** `Agenda - Simple` (or `1_Agenda - Simple`).

**Placeholders:** `idx=12` heading (the big word, e.g. "Agenda" or a section
number/name), `idx=15` list body (one line per agenda item/section).

**Use for:** agenda slides and section breaks in a longer deck.

**Budget:** heading ≤ 3 words. Up to 6 list items, each ≤ 8 words, one line
each — this is a navigation aid, not a content slide.

## 3. `big-statement`

**Template layout:** `50/50 - big title` (image fills the right half, huge
autosized quote fills the left) or `Statement title - image` (full-bleed
photo behind a centered headline) or `Headline only - Blue/Red/White bgr`
(pure color-field, no photo).

**Placeholders:** `idx=0` or the layout's headline placeholder (one
sentence, large autosized type); image placeholder if present.

**Use for:** the single most important takeaway in the deck — a finding, a
decision, a number framed as a sentence. See
`references/examples/big-statement.jpg`.

**Budget:** one sentence, ≤ 18 words, no bullets. If you need supporting
detail, put it in speaker notes, not on the slide.

## 4. `two-column`

**Template layout:** `Two columns` (title + kicker + two text blocks) or
`two columns - colorised` (two large equal panels, no title — good for a
stark before/after or definition pair).

**Placeholders:** `idx=0` title, `idx=14` kicker (one line), `idx=12`/`idx=15`
left/right bodies.

**Use for:** two parallel ideas, a before/after, or pros/cons — when there
are exactly two things to hold side by side.

**Budget:** ≤ 4 short bullets per column, ≤ 10 words per bullet. If content
is really a comparison across named attributes, use `comparison` (a table)
instead of two free bullet lists.

## 5. `three-or-four-column`

**Template layout:** `Three columns` / `1_Three columns - images` (adds a
picture above each column) or `Four columns` / `Four column - images`.

**Placeholders:** `idx=0` title, `idx=14` kicker, then one body placeholder
per column (narrow — 11–14pt by template default; see design-system.md).

**Use for:** 3–4 parallel short items (pillars, phases, workstreams). This is
the template's native alternative to a 3–5 bullet list when the items are
genuinely parallel and each deserves equal visual weight.

**Budget:** ≤ 3 short lines per column (a label + one short supporting
phrase). This is not a place for paragraphs — the columns are narrow by
design. If any column needs more than that, you have a content slide, not a
pillars slide: switch to `two-column` with real bullets, or split into
multiple slides.

## 6. `image-and-text`

**Template layout:** `50/50 - image + large title` / `50/50 - image + medium
title` / `1_50/50 - image + small title` (title size varies; pick based on
how much the title should dominate) or `50/50 - 2rd img+description` (image
top, short description below it).

**Placeholders:** `idx=0` title, `idx=10` picture (always real photography —
see Content rules), `idx=13` body.

**Use for:** grounding a claim in a real photo (a product, a place, a
customer moment) alongside a short explanation.

**Budget:** title ≤ 10 words, body ≤ 3 short bullets or 1 short paragraph
(≤ 40 words total).

## 7. `big-number-stat`

**Template layout:** `three headliner` (three big-number rows stacked, each
with a wide 70pt stat slot and a narrow annotation column) or `Statement -
image + highlights` (full photo, headline, and three short highlight
captions along the bottom).

**Placeholders:** big stat number (autosized ~70pt, **number only** — e.g.
"40%", not a sentence; a full sentence at 70pt overflows into the row below,
see `references/rubric.md`), small annotation/caption text (12–16pt, carries
the label and optional source note).

**Use for:** 1–3 headline numbers that are the actual point of the slide —
this is the template's native "big number" layout and should be preferred
any time a slide's content is fundamentally "here is a number that matters."

**Budget:** each big number is a bare figure (≤ 6 characters, e.g. "18%",
"2.3x"). Its label is a ≤ 8-word phrase in the small annotation slot, plus an
optional ≤ 6-word source note. Never put the label's words into the number
field itself.

## 8. `comparison`

**Template layout:** `Statement - dual image` (two large images side by side,
each with a short caption overlaid — see
`references/examples/dual-image-statement.jpg`, "My home / Our home") for a
visual comparison; or a genuine data table (`add_table` in
`scripts/generate_deck.py`) placed in a `Simple paragraph` / `Empty page`
layout's safe area for a comparison across named attributes/criteria.

**Use for:** two options, two states, or a structured comparison across
several named criteria.

**Budget:** two captions ≤ 4 words each for the dual-image form. For a table,
≤ 5 rows × ≤ 4 columns before it needs to split across slides.

## 9. `timeline`

**No dedicated template layout** — build on `Empty page` using the simple
milestone-row recipe in `scripts/generate_deck.py`
(`add_timeline(slide, milestones)`), modeled on the brand's own timeline
slide (`references/examples/timeline-simple.jpg`): a horizontal line, one
circle per milestone (larger circle = bigger milestone), a short label
above or below each, alternating to avoid collisions.

**Use for:** a sequence of dated or ordered milestones. Prefer this over a
bulleted list of dates.

**Budget:** ≤ 6 milestones per slide, each label ≤ 5 words. Do not attempt
the brand's circular "year wheel" diagram (`timeline-year-cycle.jpg`) —
it's hand-built in the source file and too complex to generate reliably;
use the linear milestone row instead.

## 10. `closing`

**Template layout:** `End slide - yellow` / `1_End slide - Green` / `End
slide - Black bgr` / `End slide - image bgr` (match the mood used on the
`title` slide).

**Placeholders:** headline ("Thank you!" or a closing line), small caption
(presenter/team name).

**Use for:** the last slide of the deck.

**Budget:** headline ≤ 4 words. Caption ≤ 2 short lines.

---

## Choosing between a column layout and a chart/table

Before using `two-column` / `three-or-four-column` for anything numeric, ask:
"is this actually comparable data?" If yes, use a native chart (bar/line/pie
via `scripts/generate_deck.py`'s `add_chart`) or a table instead — see
`references/content-rules.md`.

## Charts and tables draw independently of the layout's own placeholders

`add_chart` and `add_table` always draw into the full safe-content band
(below the title, above the footer, margin-to-margin) regardless of which
named layout/mood you picked — they do not use that layout's own body/image
placeholder geometry. This means any layout with a title placeholder works
as a chart/table host; pick the one whose mood and title size fit the
slide, then add the chart/table on top. `add_chart` disables the chart's own
auto-generated title by default (the slide title already states the
takeaway — a second chart-internal title is a duplicate, not a clarification).
`add_table` always sets its own explicit header (blue fill, white bold text)
and alternating white/light-grey body rows — it does not rely on
PowerPoint's default table style, which can otherwise produce invisible
header text against a custom fill.
