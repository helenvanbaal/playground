---
name: powerpoint-producer
description: "Create and redesign PowerPoint decks using the S&D template's real layouts and placeholders. Builds a varied visual story, runs a fast structural lint, and renders a preview for visual review."
argument-hint: "Create a PowerPoint from this information. Mention audience, purpose, and desired length if known."
user-invocable: true
disable-model-invocation: false
---

# PowerPoint Producer

Create a deck from the S&D template. Use the real template layouts and their
placeholders; do not invent slide geometry, draw a custom timeline, or create
charts/tables outside the template's layouts.

## Sources of truth

- [`layouts.json`](layouts.json) is the machine-readable catalogue of the
  template's 39 real layouts, placeholder indices, geometry, estimated text
  capacity, and brand colors/typeface.
- [`outline.schema.json`](outline.schema.json) defines the editable outline.
- [`assets/base-template.pptx`](assets/base-template.pptx) provides the master
  and layouts.
- [`scripts/generate_deck.py`](scripts/generate_deck.py) fills only native
  placeholders. Outline `layout` values must exactly match a real template
  layout name; text and image values are keyed by the placeholder `idx` listed
  in the catalogue. It also removes copyright and leftover tracker-demo text
  from masters/layouts so they do not appear in generated decks.
- [`scripts/lint_deck.py`](scripts/lint_deck.py) checks placeholder use,
  approximate text capacity, low placeholder fill, font overrides, and
  out-of-slide geometry.
- [`scripts/render_preview.py`](scripts/render_preview.py) renders preview
  images. LibreOffice may substitute Noto IKEA Latin, so use the preview to
  spot gross spacing/overlap issues, not to certify exact PowerPoint font fit.

The checked-in base template currently contains **zero sample slides**; the
original source deck was already trimmed. Its raster examples remain under
`references/examples/`, but are not editable source slides. When rebuilding
from a source brand deck, keep its sample slides (the default); use
`--strip-slides` only when a layout-only file is intentionally wanted. The
generator removes retained samples from the output and creates slides from
their real layouts. It does not clone sample compositions.

## Workflow

1. Identify the audience, purpose, source material, and desired length. Do not
   invent facts, metrics, or quotes; mark missing details `[to confirm]`.
2. Plan the deck's visual rhythm before writing the full outline. Make a
   slide-by-slide run of show that alternates appropriate real layout families
   (for example: statement, comparison, big-number, agenda, image) and moods.
   Avoid using the same layout three slides in a row or letting one layout
   family dominate the content slides. Repetition is fine when the story truly
   calls for a repeated structure, but it should be a conscious choice.
3. Choose a **real layout** by exact name from `layouts.json`, then select only
   its native placeholders. Match the amount and shape of content to the
   layout, not vice versa:
   - If a slide has only one or two short points, prefer an intentionally
     spacious layout with a large headline, big-number, or color panel. Do not
     park a few short lines in a dense or narrow-column layout.
   - For several genuinely parallel facts, use columns or a comparison
     composition; for a sequence, use the template's agenda/list layout.
   - When a slide looks empty, first choose a more suitable real layout or a
     larger-type layout variant. Add detail only when it is supported by the
     source; never pad copy to satisfy a fill percentage.
   - Use distinct template background moods (black, white, yellow, blue, red)
     where they support the story: keep color-field slides for emphasis or
     transitions, and use light layouts for detailed comparison. Do not apply
     one background treatment to every slide.
   - Never fake photography or leave empty image placeholders. If no real
     images are available, choose a photo-free layout from the catalogue.
   - Respect the layout's actual type scale. Avoid narrow-column layouts when
     their small template type makes the copy hard to read; choose a wider
     placeholder or a larger-type composition instead of shrinking text.
   Keep the outline's `key_message` and notes, but write visible content in
   `text` and `images` using the selected layout's real placeholder indices.
4. Generate and lint:

   ```bash
   python3 .github/skills/powerpoint-producer/scripts/generate_deck.py \
     --outline outline.json --output outputs/deck.pptx
   python3 .github/skills/powerpoint-producer/scripts/lint_deck.py \
     --pptx outputs/deck.pptx \
     --catalog .github/skills/powerpoint-producer/layouts.json
   ```

   Text capacity and fill are estimates, not hard rendering limits. Treat
   under-fill warnings as a prompt to reconsider that placeholder/layout, and
   repeated-family warnings as a prompt to improve visual rhythm. Fix them by
   changing to a better-fitting layout or, only when source-supported, adding
   missing content. Do not waive a whole-deck pattern just because each slide
   is structurally valid.
5. Render all slides once for visual review:

   ```bash
   python3 .github/skills/powerpoint-producer/scripts/render_preview.py \
     --pptx outputs/deck.pptx --outdir /tmp/deck-previews
   ```

   Check for obvious clipping, overlaps, missing text, and elements outside
   the safe area. If you revise slides, rerender only those slide images with
   `--slides 2,4`; LibreOffice still converts the complete deck to PDF.
6. Use a separate reviewer only when requested or when the deck is high-stakes.
   Deliver the named `.pptx`, outline path, slide count, unresolved
   `[to confirm]` markers, and any preview limitations.

## Refresh the catalogue/template

After receiving a new brand source file, preserve its sample slides and
rebuild the base template:

```bash
python3 .github/skills/powerpoint-producer/scripts/trim_template.py \
  --source "/path/to/Brand Template.pptx" \
  --output .github/skills/powerpoint-producer/assets/base-template.pptx
python3 .github/skills/powerpoint-producer/scripts/build_layout_catalog.py \
  --template .github/skills/powerpoint-producer/assets/base-template.pptx \
  --output .github/skills/powerpoint-producer/layouts.json
```

Review catalogue changes when the brand source changes. The heuristic text
capacity and fill checks should not override the source slide's intended
composition or PowerPoint's actual rendering.
