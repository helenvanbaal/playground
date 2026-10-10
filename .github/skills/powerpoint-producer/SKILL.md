---
name: powerpoint-producer
description: "Use when creating, generating, drafting, or redesigning PowerPoint PPTX presentations on the S&D brand template -- executive decks, strategy stories, workshop decks, status updates. Produces decks via a fixed design system (named layouts, brand palette/type scale), native charts/tables/timelines instead of bullet walls, and a mandatory render-to-image self-review pass before delivery. Use for 'make me a deck', 'turn this into a presentation', 'build a PowerPoint', or improving an existing slide deck's visual quality."
argument-hint: "Create a PowerPoint from this information. Mention audience, purpose, and desired length if known."
user-invocable: true
disable-model-invocation: false
---

# PowerPoint Producer

Produce decision-ready, on-brand PowerPoint decks on the S&D template. The
single biggest lever for quality is **looking at the rendered result before
delivering it** — this skill makes that step mandatory, not optional.

## Why this skill exists

Slides are usually mediocre not because the content is wrong, but because
the model never sees the rendered result, has unlimited design freedom, and
defaults to bullets for everything. This skill fixes all three: a fixed
design system (colors/fonts/margins), a small library of named layouts
instead of "design a slide," and a required render-and-look pass with a
scoring rubric before any deck is called done.

## Resources

- [Design system](references/design-system.md) — palette, fonts, type scale, margins (read first; everything else assumes this)
- [Layout library](references/layouts.md) — the 10 named layouts you must choose from, with real template mappings and word budgets
- [Content rules](references/content-rules.md) — one idea per slide, word/bullet limits, chart-over-bullets decision table
- [Self-critique rubric](references/rubric.md) — the scoring pass after rendering, and the avoid-list
- [Good and bad examples](references/good-bad-examples.md) — real brand screenshots plus one deliberately broken slide
- [outline.schema.json](outline.schema.json) — the structured outline format
- [scripts/generate_deck.py](scripts/generate_deck.py) — builds the .pptx from an outline
- [scripts/render_preview.py](scripts/render_preview.py) — renders the .pptx to PNGs for the required look pass
- [scripts/trim_template.py](scripts/trim_template.py) — regenerates `assets/base-template.pptx` if the brand source file changes
- `assets/base-template.pptx` — the brand template (masters/layouts/theme/fonts only, demo slides stripped)

## Inputs to request or identify

| Input | Required | Notes |
|---|---:|---|
| Source material | Yes | Notes, analysis, research, a brief, or pasted content |
| Audience | Yes | Leadership, team, customer, workshop participants, etc. |
| Purpose | Yes | Inform, decide, align, persuade, train, review |
| Desired length | No | If absent, choose the shortest deck that carries the message |
| Images | No | Real photos for `image-and-text`/`comparison`/`big-statement` slides; ask for them rather than inventing stock-photo placeholders. If none are available, prefer a chart/table/stat layout over a slide that needed a photo it doesn't have |
| Output path | No | Default to `outputs/<slug>-deck.pptx` |

Do not invent facts, metrics, or quotes. Mark missing content `[to confirm]`.

## Workflow

### 1. Plan: write the outline before touching layout

For every slide, decide the **key message** (one sentence, the takeaway),
the **named layout** (one of the 10 in `references/layouts.md`), and the
**visual element** (bullets/chart/table/image/big-number/timeline — default
away from bullets per `references/content-rules.md`). Reviewing this
structure catches a weak narrative before any visual work starts.

Write this as JSON matching [outline.schema.json](outline.schema.json). Keep
it in the conversation or save it alongside the output — it's the editable
source of truth, not a throwaway.

### 2. Build the deck

```bash
python3 -m pip install -r .github/skills/powerpoint-producer/requirements.txt   # once per environment
python3 .github/skills/powerpoint-producer/scripts/generate_deck.py \
  --outline outline.json --output outputs/deck.pptx
```

Use `--inspect-template` to list real layout names/placeholder indices if
you need to confirm something beyond what `references/layouts.md` documents
(e.g. after running `trim_template.py` against an updated brand file).

Read any printed content-rule warnings (bullet/word-count overages) and fix
the outline before proceeding — don't carry a warning into delivery.

### 3. Render and look — required, not optional

```bash
python3 .github/skills/powerpoint-producer/scripts/render_preview.py \
  --pptx outputs/deck.pptx --outdir outputs/deck_previews
```

**View every rendered PNG.** This is the step that catches what code review
cannot: overflow, overlap, cramped layouts, broken contrast. Check explicitly
for:

- Text clipped, truncated, or overflowing its placeholder
- Elements overlapping (especially multi-part layouts like `big-number-stat`
  and `timeline`, where one overflowing element can bleed into the next)
- Uneven margins, or anything inside the right-edge copyright strip / bottom
  footer band (see `references/design-system.md` safe area)
- Low contrast, especially text drawn over a photo
- Orphaned single words wrapping alone onto a final line
- A slide's title or other expected text actually missing — not just that
  `generate_deck.py` exited without error

Fix defects by editing the outline/script, not by eyeballing the XML. Then
**regenerate and re-render** — never assume a code fix produced the intended
visual result without looking again.

### 4. Self-critique against the rubric

Run the scoring pass in [references/rubric.md](references/rubric.md) on every
slide: hierarchy, alignment, whitespace, contrast, consistency, text density.
Fix anything scoring below 4, then re-render and re-check it. For a
high-stakes deck (leadership review, external audience), spawn a separate
`task` agent with only the rendered images and the rubric — a builder that
hasn't seen its own reasoning scores more honestly than the one that built it.

### 5. Deliver

Report the output path, slide count, any `[to confirm]` markers left in the
deck, and confirm the render-and-look and rubric passes were both completed
(don't report a deck as done otherwise). Mention any content that had to be
split across slides or converted to a different visual element than
requested, and why.

## Content rules (summary — see references/content-rules.md for detail)

- One idea per slide. Title states the takeaway, not the topic.
- ≤ 30–40 words of body text, ≤ 5 bullets, nothing under 14pt that you
  authored yourself.
- For every slide, ask "can this be a chart, table, big number, timeline, or
  real photo instead of bullets?" Default to yes.
- Native charts/tables only (`add_chart`/`add_table`) — never a pasted
  screenshot of a chart from another tool.
- Speaker notes carry the detail the slide doesn't show.

## Updating the brand template

If the user provides a new/updated brand `.pptx`, rebuild the base template
rather than hand-editing `assets/base-template.pptx`:

```bash
python3 .github/skills/powerpoint-producer/scripts/trim_template.py \
  --source "/path/to/New Brand Template.pptx" \
  --output .github/skills/powerpoint-producer/assets/base-template.pptx
```

Then re-verify `references/design-system.md` and the `LAYOUT_ROLES` mapping
in `generate_deck.py` still match (layout names/placeholder indices can shift
between template versions) using `--inspect-template`, and update both if
they've changed.

## Output quality bar

A delivered deck must:

- Use only layouts from `references/layouts.md` and colors/fonts from
  `references/design-system.md` — no invented layouts, no off-palette colors
- Have had every slide rendered to an image and visually reviewed, with any
  defects fixed and re-rendered
- Score ≥ 4 on every rubric dimension for every slide, or have a documented
  reason it doesn't
- Carry the deck's narrative takeaway in slide titles, not slide topics
- Use speaker notes for the detail that would otherwise crowd a slide
- Save to a named `.pptx` under `outputs/` unless told otherwise
