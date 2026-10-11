---
name: powerpoint-producer
description: "Create and redesign PowerPoint decks using the S&D template's real layouts and placeholders. Builds an editable outline, runs a fast structural lint, and renders a preview for visual review."
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
  in the catalogue. It also normalizes the inherited copyright line once on
  the slide master to 6pt Noto IKEA Latin in brand grey, keeping its placement
  consistent and subdued across slides.
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
2. Draft the outline JSON first. Choose a **real layout** by name from
   `layouts.json`, then select only the placeholders appropriate to that
   layout. Match content volume to the placeholder's capacity and the real
   examples; do not fill a dense layout with a few isolated labels.
3. Generate and lint:

   ```bash
   python3 .github/skills/powerpoint-producer/scripts/generate_deck.py \
     --outline outline.json --output outputs/deck.pptx
   python3 .github/skills/powerpoint-producer/scripts/lint_deck.py \
     --pptx outputs/deck.pptx \
     --catalog .github/skills/powerpoint-producer/layouts.json
   ```

   Text capacity and the 40% fill threshold are estimates, not hard rendering
   limits. Treat low-fill warnings as a prompt to reconsider the layout, not a
   mandate to add copy. Fix all structural errors.
4. Render all slides once for visual review:

   ```bash
   python3 .github/skills/powerpoint-producer/scripts/render_preview.py \
     --pptx outputs/deck.pptx --outdir /tmp/deck-previews
   ```

   Check for obvious clipping, overlaps, missing text, and elements outside
   the safe area. If you revise slides, rerender only those slide images with
   `--slides 2,4`; LibreOffice still converts the complete deck to PDF.
5. Use a separate reviewer only when requested or when the deck is high-stakes.
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
