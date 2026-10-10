#!/usr/bin/env python3
"""Generate a brand-accurate PPTX deck from a structured outline.

Builds on the trimmed base template (assets/base-template.pptx by default)
using named layouts from references/layouts.md. Prefers template picture
placeholders (crop-to-fill, never stretches), native python-pptx charts and
tables, and a simple timeline recipe for milestone sequences.

Usage:
    python3 generate_deck.py --inspect-template
    python3 generate_deck.py --outline outline.json --output outputs/deck.pptx
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

try:
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_LEGEND_POSITION
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Emu, Inches, Pt
except ImportError as exc:  # pragma: no cover - exercised by CLI users
    raise SystemExit(
        "Missing dependency: python-pptx. Install it with: "
        "python3 -m pip install -r requirements.txt"
    ) from exc


DEFAULT_TEMPLATE_PATH = Path(__file__).resolve().parents[1] / "assets" / "base-template.pptx"

# Brand palette (see references/design-system.md) -- never use colors outside this set.
PALETTE = {
    "blue": RGBColor(0x00, 0x58, 0xA3),
    "yellow": RGBColor(0xFF, 0xDB, 0x00),
    "black": RGBColor(0x00, 0x00, 0x00),
    "white": RGBColor(0xFF, 0xFF, 0xFF),
    "light_grey": RGBColor(0xF0, 0xF0, 0xF0),
    "grey": RGBColor(0x57, 0x57, 0x57),
    "orange": RGBColor(0xF5, 0x4C, 0x22),
    "green": RGBColor(0x36, 0xC3, 0x36),
}
CHART_SERIES_ORDER = ["blue", "grey", "orange", "green", "yellow"]

# Slide safe area (see references/design-system.md)
SAFE_LEFT = Inches(0.25)
SAFE_RIGHT_MARGIN = Inches(1.0)   # keep clear of the rotated copyright mark
SAFE_TOP = Inches(0.56)
SAFE_BOTTOM = Inches(7.2)
MIN_AUTHORED_FONT_PT = 14

# idx roles per real template layout name. Re-verify with --inspect-template
# whenever the source brand file changes (names/idx can shift between
# versions); this map is generated from assets/base-template.pptx.
LAYOUT_ROLES: dict[str, dict[str, int]] = {
    "Title - Black": {"title": 0, "subtitle": 10},
    "Title - White": {"title": 0, "subtitle": 10},
    "Title - Yellow": {"title": 0, "subtitle": 10},
    "Agenda - Simple": {"heading": 12, "items": 15},
    "1_Agenda - Simple": {"heading": 12, "items": 15},
    "Two columns": {"title": 0, "kicker": 14, "col1": 12, "col2": 15},
    "two columns - colorised": {"col1": 13, "col2": 14},
    "Three columns": {"title": 0, "kicker": 14, "col1": 13, "col2": 18, "col3": 19},
    "Four columns": {"title": 0, "kicker": 14, "col1": 13, "col2": 20, "col3": 21, "col4": 22},
    "1_Three columns - images": {
        "title": 0, "kicker": 14, "col1": 13, "col2": 18, "col3": 19,
        "img1": 20, "img2": 21, "img3": 22,
    },
    "Four column - images": {
        "title": 0, "kicker": 14, "col1": 13, "col2": 26, "col3": 27, "col4": 28,
        "img1": 20, "img2": 23, "img3": 24, "img4": 25,
    },
    "50/50 - image + large title": {"title": 0, "image": 10, "body": 13},
    "1_50/50 - image + medium title": {"title": 0, "image": 10, "body": 13},
    "1_50/50 - image + small title": {"kicker": 14, "image": 10, "body": 13},
    "50/50 - 2rd img+description": {"title": 0, "image": 10, "body": 13, "caption": 26},
    "50/50 - big title": {"title": 0, "image": 10},
    "three headliner": {
        "kicker": 20, "stat1": 17, "stat2": 18, "stat3": 19,
        "note1": 14, "note2": 15, "note3": 16,
    },
    "Statement title - image": {"image": 10, "kicker": 20, "headline": 21},
    "Statement - image + highlights": {"image": 10, "headline": 23, "h1": 20, "h2": 21, "h3": 22},
    "Statement - dual image": {"image1": 10, "image2": 11, "caption1": 21, "caption2": 22},
    "Full bleed image": {"image": 10},
    "Simple paragraph": {"body": 13},
    "Empty page": {},
    "Headline only - Blue bgr": {"headline": 0},
    "2_Headline only - Red bgr": {"headline": 0},
    "1_Headline only - white bgr": {"headline": 0},
    "1_End slide - Green": {"headline": 21, "caption": 20},
    "End slide - yellow": {"headline": 21, "caption": 20},
    "End slide - Black bgr": {"headline": 21, "caption": 20},
    "End slide - image bgr": {"headline": 21, "caption": 20},
}

# Named layout -> mood -> real template layout name. See references/layouts.md.
NAMED_LAYOUT_MOODS: dict[str, dict[str, str]] = {
    "title": {"black": "Title - Black", "white": "Title - White", "yellow": "Title - Yellow"},
    "section-divider": {"default": "Agenda - Simple"},
    "big-statement": {
        "image": "50/50 - big title",
        "photo-full": "Statement title - image",
        "blue": "Headline only - Blue bgr",
        "red": "2_Headline only - Red bgr",
        "white": "1_Headline only - white bgr",
    },
    "two-column": {"default": "Two columns", "colorised": "two columns - colorised"},
    "image-and-text": {
        "large-title": "50/50 - image + large title",
        "medium-title": "1_50/50 - image + medium title",
        "small-title": "1_50/50 - image + small title",
        "image-top": "50/50 - 2rd img+description",
    },
    "big-number-stat": {"rows": "three headliner", "photo": "Statement - image + highlights"},
    "comparison": {"dual-image": "Statement - dual image", "table": "Simple paragraph"},
    "timeline": {"default": "Empty page"},
    "closing": {
        "yellow": "End slide - yellow",
        "green": "1_End slide - Green",
        "black": "End slide - Black bgr",
        "image": "End slide - image bgr",
    },
}
DEFAULT_MOOD: dict[str, str] = {
    "title": "black",
    "section-divider": "default",
    "big-statement": "image",
    "two-column": "default",
    "image-and-text": "large-title",
    "big-number-stat": "rows",
    "comparison": "dual-image",
    "timeline": "default",
    "closing": "black",
}

WARNINGS: list[str] = []


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--template", help=f"Path to a .pptx template. Defaults to {DEFAULT_TEMPLATE_PATH}")
    parser.add_argument("--outline", help="Path to a JSON outline file")
    parser.add_argument("--output", help="Path to write the generated .pptx")
    parser.add_argument("--inspect-template", action="store_true", help="Print slide layouts and placeholders")
    args = parser.parse_args()

    template_path = Path(args.template).expanduser().resolve() if args.template else DEFAULT_TEMPLATE_PATH
    if not template_path.exists():
        raise SystemExit(
            f"Template not found: {template_path}\n"
            f"Run scripts/trim_template.py against your brand .pptx first, or pass --template."
        )

    presentation = Presentation(str(template_path))

    if args.inspect_template:
        inspect_template(presentation)
        return 0

    if not args.outline or not args.output:
        raise SystemExit("--outline and --output are required unless --inspect-template is used")

    outline = load_outline(Path(args.outline).expanduser().resolve())
    build_deck(presentation, outline)

    output_path = Path(args.output).expanduser().resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.suffix.lower() != ".pptx":
        raise SystemExit("Output path must end with .pptx")
    presentation.save(str(output_path))
    print(f"Wrote {output_path}")

    if WARNINGS:
        print("\nContent-rule warnings (see references/content-rules.md):")
        for warning in WARNINGS:
            print(f"  - {warning}")
    return 0


def inspect_template(presentation: Presentation) -> None:
    for layout in presentation.slide_masters[0].slide_layouts:
        print(f"=== {layout.name} ===")
        for shape in layout.placeholders:
            placeholder = shape.placeholder_format
            print(f"  idx={placeholder.idx} type={placeholder.type} name={shape.name!r}")


def load_outline(outline_path: Path) -> dict[str, Any]:
    if not outline_path.exists():
        raise SystemExit(f"Outline not found: {outline_path}")
    with outline_path.open("r", encoding="utf-8") as outline_file:
        outline = json.load(outline_file)
    if not isinstance(outline, dict) or not isinstance(outline.get("slides"), list):
        raise SystemExit("Outline must be a JSON object with a slides array")
    return outline


# ---------------------------------------------------------------------------
# Deck assembly
# ---------------------------------------------------------------------------

def build_deck(presentation: Presentation, outline: dict[str, Any]) -> None:
    layouts_by_name = {layout.name: layout for layout in presentation.slide_masters[0].slide_layouts}
    for index, slide_data in enumerate(outline["slides"]):
        if not isinstance(slide_data, dict):
            raise SystemExit(f"Slide {index + 1} must be a JSON object")
        layout = resolve_layout(layouts_by_name, slide_data, index)
        slide = presentation.slides.add_slide(layout)
        populate_slide(slide, layout.name, slide_data, index)


def resolve_layout(layouts_by_name: dict[str, Any], slide_data: dict[str, Any], index: int):
    named_layout = slide_data.get("layout")
    if not named_layout:
        raise SystemExit(f"Slide {index + 1} is missing required 'layout'")

    if named_layout in ("three-or-four-column",):
        columns = slide_data.get("columns") or []
        has_images = any(isinstance(col, dict) and col.get("image") for col in columns)
        if len(columns) >= 4:
            template_name = "Four column - images" if has_images else "Four columns"
        else:
            template_name = "1_Three columns - images" if has_images else "Three columns"
    else:
        moods = NAMED_LAYOUT_MOODS.get(named_layout)
        if not moods:
            raise SystemExit(
                f"Slide {index + 1}: unknown layout '{named_layout}'. "
                "See references/layouts.md for the allowed layout names."
            )
        mood = slide_data.get("mood") or DEFAULT_MOOD.get(named_layout) or next(iter(moods))
        template_name = moods.get(mood)
        if not template_name:
            raise SystemExit(
                f"Slide {index + 1}: unknown mood '{mood}' for layout '{named_layout}'. "
                f"Valid moods: {sorted(moods)}"
            )

    layout = layouts_by_name.get(template_name)
    if layout is None:
        raise SystemExit(f"Slide {index + 1}: template layout '{template_name}' not found in the base template")
    return layout


def populate_slide(slide, template_name: str, slide_data: dict[str, Any], index: int) -> None:
    roles = LAYOUT_ROLES.get(template_name, {})
    placeholders_by_idx = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

    def set_text(role: str, text: str | None, *, bullets: list | None = None) -> bool:
        if role not in roles or roles[role] not in placeholders_by_idx:
            return False
        shape = placeholders_by_idx[roles[role]]
        if not shape.has_text_frame:
            return False
        text_frame = shape.text_frame
        text_frame.clear()
        if bullets:
            for i, bullet in enumerate(bullets):
                paragraph = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
                paragraph.text = str(bullet)
        elif text:
            text_frame.text = str(text)
        return True

    def set_picture(role: str, image: dict | None) -> None:
        if not image or role not in roles or roles[role] not in placeholders_by_idx:
            if image:
                WARNINGS.append(f"Slide {index + 1}: no picture placeholder for role '{role}'")
            return
        shape = placeholders_by_idx[roles[role]]
        image_path = Path(image["path"]).expanduser().resolve()
        if not image_path.exists():
            raise SystemExit(f"Slide {index + 1}: image not found: {image_path}")
        # insert_picture crops to fill the placeholder's native aspect ratio
        # instead of stretching -- never use add_picture for a picture
        # placeholder (see "avoid: stretched images" in references/rubric.md).
        shape.insert_picture(str(image_path))

    # Title / headline / kicker roles. If the chosen layout has none of these
    # roles (e.g. "Empty page" for a timeline), fall back to a drawn textbox
    # at the standard title position rather than silently dropping the text.
    title_text = slide_data.get("title")
    placed_title = set_text("title", title_text)
    placed_title = set_text("headline", title_text or slide_data.get("headline")) or placed_title
    placed_title = set_text("heading", title_text or slide_data.get("heading")) or placed_title
    if title_text and not placed_title:
        add_fallback_title(slide, title_text)
    set_text("kicker", slide_data.get("kicker"))
    set_text("subtitle", slide_data.get("subtitle"))
    set_text("caption", slide_data.get("caption"))

    # Simple single body
    if slide_data.get("body") or slide_data.get("bullets"):
        check_content_rules(slide_data, index)
        set_text("body", slide_data.get("body"), bullets=slide_data.get("bullets"))

    # Agenda / section-divider items
    if slide_data.get("items"):
        set_text("items", None, bullets=slide_data["items"])

    # Two/three/four columns
    columns = slide_data.get("columns") or []
    for col_index, column in enumerate(columns, start=1):
        role = f"col{col_index}"
        if role not in roles:
            WARNINGS.append(f"Slide {index + 1}: layout has no column slot {col_index}, extra column ignored")
            continue
        if isinstance(column, dict):
            heading = column.get("heading")
            body = column.get("body")
            text = f"{heading}\n{body}" if heading and body else (heading or body)
            check_text_length(text, index, f"column {col_index}")
            set_text(role, text)
            set_picture(f"img{col_index}", column.get("image"))
        else:
            set_text(role, str(column))

    # Images (single or dual)
    set_picture("image", slide_data.get("image"))
    images = slide_data.get("images") or []
    if images:
        set_picture("image1", images[0] if len(images) > 0 else None)
        set_picture("image2", images[1] if len(images) > 1 else None)
        if len(images) > 1:
            set_text("caption1", images[0].get("caption") if isinstance(images[0], dict) else None)
            set_text("caption2", images[1].get("caption") if isinstance(images[1], dict) else None)

    # Big-number stats (three headliner / statement highlights).
    # The big stat slot gets the NUMBER ONLY -- it is already autosized to
    # ~70pt, and a full sentence there overflows into the next row (see
    # references/rubric.md, "clipped or overflowing text"). The label and any
    # note go into the small annotation slot instead.
    stats = slide_data.get("stats") or []
    stat_roles = ["stat1", "stat2", "stat3"]
    note_roles = ["note1", "note2", "note3"]
    highlight_roles = ["h1", "h2", "h3"]
    for i, stat in enumerate(stats[:3]):
        number = stat.get("number", "")
        label = stat.get("label", "")
        note = stat.get("note", "")
        annotation = " — ".join(part for part in (label, note) if part)
        if stat_roles[i] in roles:
            set_text(stat_roles[i], number)
            if note_roles[i] in roles:
                set_text(note_roles[i], annotation)
        elif highlight_roles[i] in roles:
            # "Statement - image + highlights" has one caption slot per
            # highlight (no separate number/label split), so combine them.
            set_text(highlight_roles[i], f"{number} — {label}".strip(" —") if label else number)

    # Native chart / table / timeline drawn into the safe content area
    if slide_data.get("chart"):
        add_chart(slide, slide_data["chart"])
    if slide_data.get("table"):
        add_table(slide, slide_data["table"])
    if slide_data.get("timeline"):
        add_timeline(slide, slide_data["timeline"]["milestones"])

    if slide_data.get("notes"):
        slide.notes_slide.notes_text_frame.text = str(slide_data["notes"])


# ---------------------------------------------------------------------------
# Content-rule checks (references/content-rules.md) -- warn, don't block
# ---------------------------------------------------------------------------

def check_content_rules(slide_data: dict[str, Any], index: int) -> None:
    bullets = slide_data.get("bullets") or []
    if len(bullets) > 5:
        WARNINGS.append(f"Slide {index + 1}: {len(bullets)} bullets exceeds the 5-bullet limit")
    for bullet in bullets:
        word_count = len(str(bullet).split())
        if word_count > 10:
            WARNINGS.append(f"Slide {index + 1}: bullet exceeds 10 words ({word_count}): {bullet!r}")
    body = slide_data.get("body")
    if body:
        word_count = len(str(body).split())
        if word_count > 40:
            WARNINGS.append(f"Slide {index + 1}: body text exceeds the ~40-word limit ({word_count} words)")


def check_text_length(text: str | None, index: int, label: str) -> None:
    if text and len(str(text).split()) > 20:
        WARNINGS.append(f"Slide {index + 1}: {label} is long for a narrow column ({text!r})")


# ---------------------------------------------------------------------------
# Fallback title (for layouts with no title/headline/heading placeholder,
# e.g. "Empty page" used by the timeline layout)
# ---------------------------------------------------------------------------

def add_fallback_title(slide, title_text: str) -> None:
    textbox = slide.shapes.add_textbox(
        SAFE_LEFT, SAFE_TOP, Inches(13.33) - SAFE_LEFT - SAFE_RIGHT_MARGIN, Inches(0.9)
    )
    text_frame = textbox.text_frame
    text_frame.word_wrap = True
    paragraph = text_frame.paragraphs[0]
    paragraph.text = title_text
    for run in paragraph.runs:
        run.font.size = Pt(32)
        run.font.name = "Noto IKEA Latin"
        run.font.color.rgb = PALETTE["black"]


# ---------------------------------------------------------------------------
# Native chart
# ---------------------------------------------------------------------------

def add_chart(slide, chart_data: dict[str, Any]) -> None:
    from pptx.enum.chart import XL_CHART_TYPE

    chart_type_map = {
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
    }
    xl_type = chart_type_map.get(chart_data.get("type", "bar"), XL_CHART_TYPE.COLUMN_CLUSTERED)

    data = CategoryChartData()
    data.categories = chart_data["categories"]
    for series in chart_data["series"]:
        data.add_series(series["name"], series["values"])

    left = SAFE_LEFT
    top = Inches(1.8)
    width = Inches(13.33) - SAFE_LEFT - SAFE_RIGHT_MARGIN
    height = SAFE_BOTTOM - top

    graphic_frame = slide.shapes.add_chart(xl_type, left, top, width, height, data)
    chart = graphic_frame.chart
    # The slide's own title placeholder already states the takeaway (see
    # references/content-rules.md); a second auto-generated chart title
    # duplicates it and reads as a bug (see references/rubric.md).
    chart.has_title = False
    chart.has_legend = len(chart_data["series"]) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False

    for i, series in enumerate(chart.plots[0].series if chart_data.get("type") != "pie" else [chart.plots[0].series[0]]):
        color_name = CHART_SERIES_ORDER[i % len(CHART_SERIES_ORDER)]
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = PALETTE[color_name]

    if chart_data.get("type") == "pie":
        for i, point in enumerate(chart.plots[0].series[0].points):
            color_name = CHART_SERIES_ORDER[i % len(CHART_SERIES_ORDER)]
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = PALETTE[color_name]


# ---------------------------------------------------------------------------
# Native table
# ---------------------------------------------------------------------------

def add_table(slide, table_data: dict[str, Any]) -> None:
    columns = table_data.get("columns") or []
    rows = table_data.get("rows") or []
    all_rows = [columns] + rows if columns else rows
    if not all_rows:
        return

    row_count = len(all_rows)
    column_count = max(len(row) for row in all_rows)
    left = SAFE_LEFT
    top = Inches(1.8)
    width = Inches(13.33) - SAFE_LEFT - SAFE_RIGHT_MARGIN
    height = min(Inches(0.5) * row_count, SAFE_BOTTOM - top)

    table_shape = slide.shapes.add_table(row_count, column_count, left, top, width, height)
    table = table_shape.table
    # Disable the default Office table style's automatic banding/coloring so
    # our explicit per-cell fill and font colors are what actually render --
    # otherwise the built-in style can set header text white-on-white against
    # our own header fill (see references/rubric.md, "contrast").
    table.horz_banding = False
    table.first_row = False

    for column_index in range(column_count):
        table.columns[column_index].width = int(width / column_count)

    header_offset = 1 if columns else 0
    for row_index, row in enumerate(all_rows):
        is_header = row_index == 0 and bool(columns)
        for column_index in range(column_count):
            value = row[column_index] if column_index < len(row) else ""
            cell = table.cell(row_index, column_index)
            cell.text = str(value)
            cell.fill.solid()
            if is_header:
                cell.fill.fore_color.rgb = PALETTE["blue"]
            elif (row_index - header_offset) % 2 == 0:
                cell.fill.fore_color.rgb = PALETTE["white"]
            else:
                cell.fill.fore_color.rgb = PALETTE["light_grey"]
            paragraph = cell.text_frame.paragraphs[0]
            for run in paragraph.runs:
                run.font.size = Pt(14)
                run.font.name = "Noto IKEA Latin"
                run.font.color.rgb = PALETTE["white"] if is_header else PALETTE["black"]
                run.font.bold = is_header


# ---------------------------------------------------------------------------
# Timeline (no dedicated template layout -- simple milestone row recipe)
# ---------------------------------------------------------------------------

def add_timeline(slide, milestones: list[dict[str, Any]]) -> None:
    if not milestones:
        return

    line_y = Inches(4.2)
    left = SAFE_LEFT + Inches(0.3)
    right = Inches(13.33) - SAFE_RIGHT_MARGIN - Inches(0.3)
    usable_width = right - left

    connector = slide.shapes.add_connector(1, left, line_y, right, line_y)  # 1 = straight connector
    connector.line.color.rgb = PALETTE["black"]
    connector.line.width = Pt(1.5)

    count = len(milestones)
    step = usable_width / max(count - 1, 1) if count > 1 else 0
    for i, milestone in enumerate(milestones):
        center_x = left + (step * i if count > 1 else usable_width // 2)
        size = Inches(0.35) if milestone.get("size") == "big" else Inches(0.2)
        circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, center_x - size // 2, line_y - size // 2, size, size
        )
        circle.fill.solid()
        color_name = "blue" if milestone.get("size") == "big" else "grey"
        circle.fill.fore_color.rgb = PALETTE[color_name]
        circle.line.fill.background()

        label = milestone.get("label", "")
        label_width = Inches(1.6)
        label_height = Inches(0.6)
        above = i % 2 == 0
        label_top = line_y - Inches(0.9) if above else line_y + Inches(0.35)
        textbox = slide.shapes.add_textbox(center_x - label_width // 2, label_top, label_width, label_height)
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        paragraph = text_frame.paragraphs[0]
        paragraph.text = label
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(MIN_AUTHORED_FONT_PT)
            run.font.name = "Noto IKEA Latin"


if __name__ == "__main__":
    sys.exit(main())
