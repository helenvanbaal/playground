#!/usr/bin/env python3
"""Render a .pptx to one PNG per slide for the required render-and-look pass.

Converts via headless LibreOffice to PDF, then rasterizes each page with
PyMuPDF. Requires LibreOffice (`soffice`) on PATH or at a common macOS
install location.

Usage:
    python3 render_preview.py --pptx outputs/deck.pptx --outdir outputs/deck_previews
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


def find_soffice() -> str:
    candidate = shutil.which("soffice") or shutil.which("libreoffice")
    if candidate:
        return candidate
    mac_path = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if mac_path.exists():
        return str(mac_path)
    raise SystemExit(
        "LibreOffice not found. Install it (e.g. `brew install --cask libreoffice`) "
        "or ensure `soffice` is on PATH -- the render-and-look pass is required "
        "before delivering a deck."
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--pptx", required=True, help="Path to the .pptx to render")
    parser.add_argument("--outdir", required=True, help="Directory to write slide PNGs into")
    parser.add_argument("--dpi", type=int, default=150, help="Rasterization DPI (default 150)")
    args = parser.parse_args()

    pptx_path = Path(args.pptx).expanduser().resolve()
    if not pptx_path.exists():
        raise SystemExit(f"File not found: {pptx_path}")

    outdir = Path(args.outdir).expanduser().resolve()
    outdir.mkdir(parents=True, exist_ok=True)

    soffice = find_soffice()
    pdf_path = outdir / (pptx_path.stem + ".pdf")

    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(pptx_path)],
        check=True,
        capture_output=True,
    )
    if not pdf_path.exists():
        raise SystemExit(f"LibreOffice did not produce the expected PDF at {pdf_path}")

    try:
        import pymupdf
    except ImportError as exc:
        raise SystemExit(
            "Missing dependency: pymupdf. Install it with: python3 -m pip install -r requirements.txt"
        ) from exc

    doc = pymupdf.open(str(pdf_path))
    written = []
    for i in range(doc.page_count):
        pixmap = doc[i].get_pixmap(dpi=args.dpi)
        png_path = outdir / f"slide-{i + 1:02d}.png"
        pixmap.save(str(png_path))
        written.append(png_path)
    doc.close()

    print(f"Rendered {len(written)} slide(s) to {outdir}")
    for path in written:
        print(f"  {path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
