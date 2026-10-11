#!/usr/bin/env python3
"""Rebuild the skill's base template from a source .pptx/.potx.

The source brand file typically ships with sample slides. They are retained by
default as the visual reference for the skill. Pass --strip-slides only when a
small, layout-only template is specifically needed.

Usage:
    python3 trim_template.py --source "/path/to/Brand Template.pptx" \
        --output ../assets/base-template.pptx

Run this again whenever the source brand file changes (new layouts, updated
theme colors/fonts, etc.) to refresh the packaged base template.
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
import zipfile
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, help="Path to the source .pptx/.potx brand file")
    parser.add_argument("--output", required=True, help="Path to write the trimmed base template")
    parser.add_argument(
        "--strip-slides",
        action="store_true",
        help="Remove source sample slides (not recommended; they are useful visual references)",
    )
    args = parser.parse_args()

    source_path = Path(args.source).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()
    if not source_path.exists():
        raise SystemExit(f"Source not found: {source_path}")

    with zipfile.ZipFile(source_path, "r") as zf:
        entries = {name: zf.read(name) for name in zf.namelist()}

    before_size = sum(len(data) for data in entries.values())

    if args.strip_slides:
        entries = drop_slides(entries)
        entries = drop_unreferenced_notes_masters(entries)
    entries = drop_unreferenced_media(entries)

    after_size = sum(len(data) for data in entries.values())

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in entries.items():
            zf.writestr(name, data)

    print(f"Wrote {output_path}")
    print(f"Uncompressed payload: {before_size / 1e6:.1f} MB -> {after_size / 1e6:.1f} MB")
    return 0


def drop_slides(entries: dict[str, bytes]) -> dict[str, bytes]:
    """Remove every slide part and all references to it."""
    slide_names = sorted(n for n in entries if re.fullmatch(r"ppt/slides/slide\d+\.xml", n))
    if not slide_names:
        return entries

    # Map slide part name -> its r:id in ppt/_rels/presentation.xml.rels
    pres_rels_name = "ppt/_rels/presentation.xml.rels"
    pres_rels = entries[pres_rels_name].decode("utf-8")
    slide_rids = []
    for slide_name in slide_names:
        target = slide_name.split("/", 1)[1]  # "slides/slideN.xml"
        m = re.search(rf'Id="([^"]+)"[^/]*Target="{re.escape(target)}"', pres_rels)
        if not m:
            m = re.search(rf'Target="{re.escape(target)}"[^/]*Id="([^"]+)"', pres_rels)
        if m:
            slide_rids.append(m.group(1))

    # Drop the relationship entries and the sldIdLst entries in presentation.xml
    for rid in slide_rids:
        pres_rels = re.sub(rf'<Relationship[^>]*Id="{re.escape(rid)}"[^>]*/>', "", pres_rels)
    entries[pres_rels_name] = pres_rels.encode("utf-8")

    pres_name = "ppt/presentation.xml"
    pres_xml = entries[pres_name].decode("utf-8")
    pres_xml = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>", "<p:sldIdLst/>", pres_xml, flags=re.S)
    entries[pres_name] = pres_xml.encode("utf-8")

    # Drop [Content_Types].xml overrides for the removed slides
    ct_name = "[Content_Types].xml"
    ct_xml = entries[ct_name].decode("utf-8")
    for slide_name in slide_names:
        part_name = "/" + slide_name
        ct_xml = re.sub(rf'<Override PartName="{re.escape(part_name)}"[^>]*/>', "", ct_xml)
    entries[ct_name] = ct_xml.encode("utf-8")

    # Remove the slide parts themselves and their rels/notes
    for slide_name in slide_names:
        entries.pop(slide_name, None)
        rels_name = slide_name.replace("slides/", "slides/_rels/") + ".rels"
        entries.pop(rels_name, None)

    return entries


def drop_unreferenced_notes_masters(entries: dict[str, bytes]) -> dict[str, bytes]:
    """Drop notesSlide parts, which only existed for the removed demo slides."""
    notes_names = [n for n in entries if re.fullmatch(r"ppt/notesSlides/notesSlide\d+\.xml", n)]
    if not notes_names:
        return entries

    ct_name = "[Content_Types].xml"
    ct_xml = entries[ct_name].decode("utf-8")
    for notes_name in notes_names:
        part_name = "/" + notes_name
        ct_xml = re.sub(rf'<Override PartName="{re.escape(part_name)}"[^>]*/>', "", ct_xml)
        entries.pop(notes_name, None)
        rels_name = notes_name.replace("notesSlides/", "notesSlides/_rels/") + ".rels"
        entries.pop(rels_name, None)
    entries[ct_name] = ct_xml.encode("utf-8")
    return entries


def drop_unreferenced_media(entries: dict[str, bytes]) -> dict[str, bytes]:
    """Keep media referenced by any retained slide, layout, or master."""
    media_names = {n for n in entries if n.startswith("ppt/media/")}
    if not media_names:
        return entries

    used_media: set[str] = set()
    for rels_name in (n for n in entries if n.endswith(".rels")):
        rels_xml = entries[rels_name].decode("utf-8")
        for target in re.findall(r'Target="([^"]*media/[^"]+)"', rels_xml):
            used_media.add("ppt/media/" + target.rsplit("media/", 1)[1])

    unused_media = media_names - used_media
    if not unused_media:
        return entries

    # Media content types are declared per-extension in [Content_Types].xml
    # Default entries, not per-part Overrides, so no content-type edit is
    # needed here -- only the unused parts themselves are removed.
    for name in unused_media:
        entries.pop(name, None)

    return entries


if __name__ == "__main__":
    sys.exit(main())
