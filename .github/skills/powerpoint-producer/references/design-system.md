# Design system — S&D brand template

This is the fixed design system extracted directly from the brand template
(`assets/base-template.pptx`, trimmed from "S&D template - 26.pptx"). Treat
every value below as a constraint, not a suggestion. Do not invent colors,
fonts, sizes, or margins outside this system.

## Colors

| Role | Name | Hex | Theme slot | Use for |
|---|---|---|---|---|
| Dominant | IKEA Blue | `#0058A3` | `dk2` | Divider/mood backgrounds, primary chart series, key emphasis |
| Accent | Yellow | `#FFDB00` | `accent2` / `accent5` | Title-slide mood, small accents, highlight callouts — use sparingly |
| Neutral (ink) | Black | `#000000` | `dk1` / `tx1` | Body text on light backgrounds, black-mood backgrounds |
| Neutral (paper) | White | `#FFFFFF` | `lt1` / `bg1` | Default slide background, text on dark backgrounds |
| Neutral (surface) | Light Grey | `#F0F0F0` | `lt2` | Subtle panel/card backgrounds, chart gridlines |
| Neutral (mid) | Grey | `#575757` | `accent1` / `accent3` | Secondary text, muted chart series, captions |
| Supporting | Orange/Red | `#F54C22` | `accent4` | Warning/negative emphasis, one of several chart series, "Red" mood divider |
| Supporting | Green | `#36C336` | `accent6` | Positive/growth emphasis, one of several chart series |

Rules:
- Backgrounds are almost always **black or white**. Color is used as an accent
  (a shape, a chip, a divider slide's mood), not as a wash across every slide.
  Only the explicit mood layouts (title, headline-only, end-slide) use a full
  color background — and even then, several "mood" names are really a white
  or black slide with a small color accent, not a full color fill. Check the
  named layout's actual rendered look in `references/examples/` before
  assuming a full color background.
- Never introduce a color outside this table. Never use default Office blue
  (`#4472C4`) or default Office theme colors.
- For charts with 2–4 series, use Blue, Grey, Orange, Green in that order
  (reserve Yellow for a single highlighted series or callout, since yellow
  text/fills have low contrast on white).
- For a single highlighted data point or "this one matters" emphasis, Yellow
  or Blue on a black/white/grey field reads strongest.

## Fonts

Two fonts, both embedded in the template:

1. **Noto IKEA Latin** — the only typeface for all real text: titles, body,
   labels, numbers, chart text. It is a serif face used at large display
   sizes for titles/statements and reads cleanly at small sizes for labels.
2. **Hej Script** — a handwritten display font used **only** as a decorative
   graphical highlight (e.g. circling a word, an underline squiggle, a
   "note-to-self" accent), copied in from the sticker pack. Never set real
   body or title copy in Hej Script, and never use it for anything that must
   stay legible at small sizes.

Do not introduce a third font. If a rendered preview shows a fallback serif
(e.g. because Noto IKEA Latin isn't installed on the rendering machine),
that's a preview-only substitution — judge layout/spacing/overflow from the
preview, not exact letterforms; the saved `.pptx` still references the
correct theme font for when the user opens it in PowerPoint.

## Type scale (real values used in the template)

| Role | Size | Where |
|---|---|---|
| Display / hero title | 60–100pt | Title slides, statement/headline-only divider slides |
| Big stat / number | 70pt | "three headliner" big-number layout, statement highlights |
| Slide title (content slides) | 32pt | Two/Three/Four-column, simple paragraph layouts |
| Large content title | 40–60pt | "Simple large title+paragraph", "Headline + paragraph" |
| Lead / kicker line | 14–18pt | The one-line sentence under a title before body content starts |
| Body text (roomy layouts) | 18pt | Simple paragraph, two-columns-colorised |
| Body text (dense layouts) | 11–14pt | Three-column, four-column body text (columns are narrow, ~2.8–3.9in) |
| Section eyebrow (top-left) | 9pt | "STRATEGY & DEVELOPMENT"-style kicker above a title |
| Footer / copyright | 8pt | Fixed footer element — never edit or repopulate this |

Authoring rule: when you add your **own** text (not filling an existing
narrow-column placeholder), never go below 14pt. The 11–12pt column sizes
are a template-defined exception for short labels in narrow columns, not a
license to shrink paragraph text generally. If a multi-column layout's
narrow placeholder can't hold your content at a readable size, you have too
much text for that slide — cut it, move it to notes, or pick a layout with
fewer/wider columns.

## Margins and safe area

Slide size is 13.33in × 7.5in (16:9, EMU 12192000 × 6858000).

- **Left margin:** 0.22–0.27in (consistent across almost every layout).
- **Right margin:** keep custom content clear of **1.0in** from the right
  edge. The brand footer includes a rotated copyright mark
  ("© Ingka Holding B.V. 2026") printed vertically near the right edge on
  every slide; placing text or important visuals there causes overlap.
- **Top margin:** 0.56in for titles. A small eyebrow/kicker label sits at
  y=0–0.3in on the left — don't duplicate or collide with it.
- **Bottom margin:** keep content above **y = 7.2in**. A small forward-arrow
  mark and the page number sit at y≈7.23–7.37in on every slide.
- There is no drawn 12-column grid in this template, but placeholder widths
  consistently divide the safe content width (≈11.86in, from x=0.21 to
  x≈12.07) into halves, thirds, or quarters with ~0.2–0.3in gutters. When
  building custom multi-column content, reuse these column math patterns
  rather than inventing new ones:
  - 2 columns: 5.88in wide, 0.2in gutter
  - 3 columns: 3.94in wide, ~0.21in gutter
  - 4 columns: 2.83–3.03in wide, ~0.25in gutter

## Backgrounds by mood

| Layout family | Background |
|---|---|
| `Title - Black`, `Headline only - Blue bgr`'s sibling black moods | Black, white text |
| `Title - White` | White, black text |
| `Title - Yellow` | Yellow, black text |
| `Headline only - Blue bgr` | Blue, white text |
| `2_Headline only - Red bgr` | Orange/red accent4, white text |
| End-slide variants | White/black/green/image background with centered "Thank you!" — pick the mood that matches the deck's opening title mood |

See `references/examples/` for real rendered screenshots of these.
