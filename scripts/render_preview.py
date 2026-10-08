"""Render PNG / PDF previews from a generated PPTX.

Three-tier strategy, in order:

1. Windows + PowerPoint installed  -> COM automation (most faithful)
2. LibreOffice available           -> headless conversion
3. Neither                         -> print an explicit notice and exit non-zero

The PPTX itself is always produced by this skill; previews are optional.
Never claim a preview exists when the renderer failed.

Usage:
    python scripts/render_preview.py output/simple_workflow.pptx
    python scripts/render_preview.py output/simple_workflow.pptx --format png --width 1920
    python scripts/render_preview.py output/simple_workflow.pptx --format both
"""
from __future__ import annotations

import argparse
import platform
import shutil
import subprocess
import sys
from pathlib import Path

PP_SAVE_AS_PDF = 32


# --------------------------------------------------------------------------- #
# backend detection
# --------------------------------------------------------------------------- #
def powerpoint_available() -> bool:
    if platform.system() != "Windows":
        return False
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return False
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Quit()
        return True
    except Exception:
        return False


def find_soffice() -> str | None:
    """Locate a LibreOffice executable, or None."""
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    if platform.system() == "Windows":
        for base in (Path(r"C:\Program Files\LibreOffice\program"),
                     Path(r"C:\Program Files (x86)\LibreOffice\program")):
            candidate = base / "soffice.exe"
            if candidate.exists():
                return str(candidate)
    if platform.system() == "Darwin":
        candidate = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
        if candidate.exists():
            return str(candidate)
    return None


# --------------------------------------------------------------------------- #
# backends
# --------------------------------------------------------------------------- #
def render_with_powerpoint(pptx: Path, outdir: Path, fmt: str, width: int) -> list[Path]:
    import win32com.client  # type: ignore

    outdir.mkdir(parents=True, exist_ok=True)
    made: list[Path] = []
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = None
    try:
        pres = app.Presentations.Open(str(pptx.resolve()), WithWindow=False)
        if fmt in ("pdf", "both"):
            pdf = outdir / f"{pptx.stem}.pdf"
            pres.SaveAs(str(pdf), PP_SAVE_AS_PDF)
            made.append(pdf)
        if fmt in ("png", "both"):
            height = int(width * 9 / 16)
            for i, slide in enumerate(pres.Slides, start=1):
                png = outdir / f"{pptx.stem}_slide{i}.png"
                slide.Export(str(png), "PNG", width, height)
                made.append(png)
    finally:
        if pres is not None:
            pres.Close()
        app.Quit()
    return made


def render_with_libreoffice(pptx: Path, outdir: Path, fmt: str, width: int, soffice: str) -> list[Path]:
    outdir.mkdir(parents=True, exist_ok=True)
    made: list[Path] = []

    if fmt in ("pdf", "both"):
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(pptx.resolve())],
            check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=180,
        )
        pdf = outdir / f"{pptx.stem}.pdf"
        if pdf.exists():
            made.append(pdf)
        else:
            raise RuntimeError("LibreOffice reported success but no PDF was produced.")

    if fmt in ("png", "both"):
        pdf = outdir / f"{pptx.stem}.pdf"
        if not pdf.exists():
            render_with_libreoffice(pptx, outdir, "pdf", width, soffice)
        made.extend(_png_from_pdf(pdf, outdir, width))
    return made


def _png_from_pdf(pdf: Path, outdir: Path, width: int) -> list[Path]:
    """Rasterise a PDF. Tries PyMuPDF, then pypdfium2, then pdftoppm."""
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(pdf)
        out = []
        zoom = width / doc[0].rect.width
        for i, page in enumerate(doc, start=1):
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            png = outdir / f"{pdf.stem}_slide{i}.png"
            pix.save(png)
            out.append(png)
        return out
    except ImportError:
        pass

    try:
        import pypdfium2 as pdfium
        doc = pdfium.PdfDocument(pdf)
        out = []
        for i, page in enumerate(doc, start=1):
            png = outdir / f"{pdf.stem}_slide{i}.png"
            page.render(scale=width / page.get_width()).to_pil().save(png)
            out.append(png)
        return out
    except ImportError:
        pass

    if shutil.which("pdftoppm"):
        subprocess.run(
            ["pdftoppm", "-png", "-r", "150", str(pdf), str(outdir / pdf.stem)],
            check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=180,
        )
        return sorted(outdir.glob(f"{pdf.stem}*.png"))

    raise RuntimeError(
        "PNG rasterisation unavailable. Install one of: PyMuPDF, pypdfium2, poppler-utils (pdftoppm)."
    )


# --------------------------------------------------------------------------- #
# entrypoint
# --------------------------------------------------------------------------- #
def render(pptx: Path, outdir: Path, fmt: str, width: int) -> list[Path]:
    if not pptx.exists():
        raise SystemExit(f"Input not found: {pptx}")

    if powerpoint_available():
        print("[preview] backend: PowerPoint (COM)")
        return render_with_powerpoint(pptx, outdir, fmt, width)

    soffice = find_soffice()
    if soffice:
        print(f"[preview] backend: LibreOffice ({soffice})")
        return render_with_libreoffice(pptx, outdir, fmt, width, soffice)

    raise SystemExit(
        "\n[preview] UNAVAILABLE: no renderer found.\n"
        "  - PowerPoint was not detected via COM, and\n"
        "  - LibreOffice (soffice) was not found on PATH.\n\n"
        "The editable PPTX has still been generated and is fully usable.\n"
        "To obtain PNG/PDF previews, install either:\n"
        "  * Microsoft PowerPoint (Windows), or\n"
        "  * LibreOffice (https://www.libreoffice.org/)\n\n"
        "For PNG output you additionally need PyMuPDF / pypdfium2 / pdftoppm.\n"
        "Do not report a preview as delivered when this message appears."
    )


def main() -> None:
    ap = argparse.ArgumentParser(description="Render PNG/PDF previews from a PPTX.")
    ap.add_argument("pptx", type=Path, help="Path to the .pptx file")
    ap.add_argument("--format", choices=["png", "pdf", "both"], default="both")
    ap.add_argument("--outdir", type=Path, default=None, help="Default: <pptx dir>/preview")
    ap.add_argument("--width", type=int, default=1920, help="PNG width in pixels (default 1920)")
    args = ap.parse_args()

    outdir = args.outdir or (args.pptx.parent / "preview")
    files = render(args.pptx, outdir, args.format, args.width)
    print(f"[preview] wrote {len(files)} file(s) to {outdir}:")
    for f in files:
        print("  ", f)


if __name__ == "__main__":
    main()
