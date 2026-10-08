"""Structural validator for editable workflow PPTX files.

Checks, per slide:
  - native shapes / connectors / pictures / text objects (inventory)
  - connectors missing arrowheads (silent bug in naive python-pptx code)
  - CJK text rendered without an a:ea typeface (falls back to a default font)
  - text with too little contrast against its own fill (unset colour renders white
    on light module fills -- invisible, and completely silent at generation time)
  - possible text overflow (estimated text size vs shape size)
  - overlapping module shapes
  - a full-slide screenshot standing in for the workflow (forbidden)

Exit code 0 = clean, 1 = warnings found. Warnings are advisory; the counts tell
you whether the figure is genuinely editable.

Usage:
    python scripts/validate_pptx.py path/to/file.pptx
"""
from __future__ import annotations

import math
import sys
import unicodedata
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

EMU_PER_IN = 914400
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _is_cjk(char: str) -> bool:
    return unicodedata.east_asian_width(char) in ("W", "F")


def _char_width_em(char: str, font_pt: float) -> float:
    # rough advance width in points
    if _is_cjk(char):
        return font_pt
    if char == " ":
        return font_pt * 0.30
    return font_pt * 0.55


def _text_overflow_warnings(shape, issues: list[str]) -> None:
    if not getattr(shape, "has_text_frame", False) or not shape.text.strip():
        return
    width_in = (shape.width or 0) / EMU_PER_IN
    height_in = (shape.height or 0) / EMU_PER_IN
    if width_in <= 0.05 or height_in <= 0.05:
        return

    total_lines = 0.0
    max_font_pt = 0.0
    for para in shape.text_frame.paragraphs:
        text = "".join(run.text for run in para.runs)
        if not text:
            total_lines += 0.4
            continue
        font_pt = 18.0
        for run in para.runs:
            if run.font.size is not None:
                font_pt = run.font.size.pt
                break
        max_font_pt = max(max_font_pt, font_pt)
        line_pt = sum(_char_width_em(c, font_pt) for c in text)
        line_in = line_pt / 72.0
        total_lines += max(1.0, math.ceil(line_in / max(width_in - 0.1, 0.1) - 0.02))

    needed_in = total_lines * max(max_font_pt, 12) * 1.25 / 72.0
    if needed_in > height_in * 1.18 + 0.06:
        issues.append(
            f"possible text overflow in '{shape.text.strip()[:18]}…': "
            f"needs ~{needed_in:.2f}in, shape is {height_in:.2f}in"
        )


def _relative_luminance(rgb) -> float:
    def channel(value: float) -> float:
        value = value / 255.0
        return value / 12.92 if value <= 0.03928 else ((value + 0.055) / 1.055) ** 2.4

    r, g, b = rgb
    return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b)


def _srgb_of(element):
    if element is None:
        return None
    node = element.find(qn("a:srgbClr"))
    if node is None:
        return None
    value = node.get("val")
    try:
        return (int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))
    except (TypeError, ValueError):
        return None


def _text_rgb(run, shape):
    """Effective text colour, or None when PowerPoint's default cannot be resolved.

    This is the check that catches invisible labels: an autoshape created by
    python-pptx carries `fontRef -> lt1` (the LIGHT theme colour), so text whose
    colour was never set renders white. On a light module fill it disappears --
    with no exception raised anywhere during generation.
    """
    rPr = run.font._rPr
    if rPr is not None:
        explicit = _srgb_of(rPr.find(qn("a:solidFill")))
        if explicit is not None:
            return explicit
        scheme = rPr.find(qn("a:solidFill"))
        if scheme is not None:
            node = scheme.find(qn("a:schemeClr"))
            if node is not None and node.get("val") in ("lt1", "bg1"):
                return (255, 255, 255)

    style = getattr(shape, "_element", None)
    if style is not None:
        style_el = style.find(qn("p:style"))
        if style_el is not None:
            font_ref = style_el.find(qn("a:fontRef"))
            if font_ref is not None:
                node = font_ref.find(qn("a:schemeClr"))
                if node is not None and node.get("val") in ("lt1", "bg1"):
                    return (255, 255, 255)
                if node is not None and node.get("val") in ("dk1", "tx1"):
                    return (0, 0, 0)
    return (0, 0, 0)  # plain text box default


def _fill_rgb(shape):
    from pptx.enum.dml import MSO_FILL

    try:
        fill = shape.fill
        if fill.type == MSO_FILL.SOLID:
            return tuple(fill.fore_color.rgb)
    except Exception:  # noqa: BLE001 - unsupported fill types are simply not checked
        return None
    return None


def _contrast_warnings(shape, issues: list[str]) -> None:
    if not getattr(shape, "has_text_frame", False) or not shape.text.strip():
        return
    if shape.shape_type not in (MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.TEXT_BOX):
        return
    fill_rgb = _fill_rgb(shape)
    if fill_rgb is None:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if not run.text.strip():
                continue
            text_rgb = _text_rgb(run, shape)
            if text_rgb is None:
                continue
            high, low = sorted(
                (_relative_luminance(text_rgb), _relative_luminance(fill_rgb)), reverse=True
            )
            ratio = (high + 0.05) / (low + 0.05)
            if ratio < 2.5:
                issues.append(
                    f"low text contrast ({ratio:.1f}:1) for '{run.text.strip()[:18]}…' — "
                    "text colour was probably never set, so PowerPoint falls back to the "
                    "light theme colour and the label becomes unreadable"
                )
            return  # one warning per shape is enough


def _arrow_warnings(shape, issues: list[str]) -> None:
    if shape.shape_type != MSO_SHAPE_TYPE.LINE:
        return
    ln = shape.line._get_or_add_ln()
    if ln.find(qn("a:tailEnd")) is None and ln.find(qn("a:headEnd")) is None:
        issues.append("connector without arrowhead: relationship direction is not rendered")


def _cjk_font_warnings(shape, issues: list[str]) -> None:
    if not getattr(shape, "has_text_frame", False):
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if not any(_is_cjk(c) for c in run.text):
                continue
            rPr = run.font._rPr
            ea = rPr.find(qn("a:ea")) if rPr is not None else None
            if ea is None or not ea.get("typeface"):
                issues.append(
                    f"CJK text '{run.text[:12]}…' has no a:ea typeface; "
                    "PowerPoint will fall back to a default CJK font"
                )
            return  # one warning per shape is enough


def _overlap_warnings(shapes, issues: list[str]) -> None:
    boxes = []
    for shp in shapes:
        if shp.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and shp.width and shp.height:
            boxes.append((shp.left, shp.top, shp.left + shp.width, shp.top + shp.height, shp))
    for i in range(len(boxes)):
        for j in range(i + 1, len(boxes)):
            ax0, ay0, ax1, ay1, a = boxes[i]
            bx0, by0, bx1, by1, b = boxes[j]
            ow = min(ax1, bx1) - max(ax0, bx0)
            oh = min(ay1, by1) - max(ay0, by0)
            if ow > 0 and oh > 0:
                overlap = ow * oh
                smaller = min((ax1 - ax0) * (ay1 - ay0), (bx1 - bx0) * (by1 - by0))
                if overlap > smaller * 0.12:
                    ta = a.text.strip()[:12] if getattr(a, "has_text_frame", False) else a.shape_id
                    tb = b.text.strip()[:12] if getattr(b, "has_text_frame", False) else b.shape_id
                    issues.append(f"module shapes overlap: '{ta}' and '{tb}'")


def _screenshot_warnings(shapes, slide_w: int, slide_h: int, issues: list[str]) -> None:
    canvas = slide_w * slide_h
    for shp in shapes:
        if shp.shape_type == MSO_SHAPE_TYPE.PICTURE and shp.width and shp.height:
            if (shp.width * shp.height) > canvas * 0.85:
                issues.append(
                    "a picture covers most of the slide - looks like a full-slide screenshot "
                    "standing in for the workflow (forbidden by this skill)"
                )


def validate(path: Path) -> dict:
    prs = Presentation(path)
    report = {
        "slides": len(prs.slides), "native_shapes": 0, "pictures": 0,
        "connectors": 0, "text_objects": 0, "arrowed_connectors": 0, "warnings": [],
    }

    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                report["pictures"] += 1
            elif shape.shape_type == MSO_SHAPE_TYPE.LINE:
                report["connectors"] += 1
                ln = shape.line._get_or_add_ln()
                if ln.find(qn("a:tailEnd")) is not None or ln.find(qn("a:headEnd")) is not None:
                    report["arrowed_connectors"] += 1
            else:
                report["native_shapes"] += 1
            if getattr(shape, "has_text_frame", False) and shape.text.strip():
                report["text_objects"] += 1

            slide_issues: list[str] = []
            _arrow_warnings(shape, slide_issues)
            _cjk_font_warnings(shape, slide_issues)
            _text_overflow_warnings(shape, slide_issues)
            _contrast_warnings(shape, slide_issues)
            for issue in slide_issues:
                report["warnings"].append(f"slide {slide_idx}: {issue}")

        _overlap_warnings(slide.shapes, report["warnings"])
        _screenshot_warnings(slide.shapes, prs.slide_width, prs.slide_height, report["warnings"])

    return report


def main() -> None:
    if len(sys.argv) != 2:
        raise SystemExit("Usage: python scripts/validate_pptx.py path/to/file.pptx")
    path = Path(sys.argv[1])
    if not path.exists():
        raise SystemExit(f"File not found: {path}")

    report = validate(path)
    print(
        f"slides={report['slides']} native_shapes={report['native_shapes']} "
        f"connectors={report['connectors']} arrowed={report['arrowed_connectors']} "
        f"pictures={report['pictures']} text_objects={report['text_objects']}"
    )
    if report["pictures"] and report["connectors"] == 0 and report["native_shapes"] <= 1:
        report["warnings"].append(
            "picture-only slide: the workflow is not editable at all"
        )
    for warning in report["warnings"]:
        print("WARNING:", warning)
    if not report["warnings"]:
        print("OK: no structural issues found")
    sys.exit(1 if report["warnings"] else 0)


if __name__ == "__main__":
    main()
