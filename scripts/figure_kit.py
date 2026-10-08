"""Shared primitives for building editable PowerPoint figures.

This module exists because two things are easy to get silently wrong:

1. Arrowheads. `python-pptx` has no `line.end_arrowhead` API (upstream issue #375).
   Assigning it does NOT raise -- it just creates a dead attribute, and the PPTX
   ends up with connectors that have no arrowhead at all. Use `add_arrow()`.

2. CJK fonts. `run.font.name = "..."` only writes `a:latin`. Chinese glyphs are
   rendered with the `a:ea` (East Asian) typeface, so the font silently falls back
   to a default unless `a:ea` is written too. Use `set_font()`.

Both helpers are used by the examples, by `convert_image_to_pptx.py` and by the tests.
"""
from __future__ import annotations

from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.oxml import parse_xml
from pptx.util import Inches, Pt

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
EMU_PER_INCH = 914400

REPO_ROOT = Path(__file__).resolve().parents[1]
THEME_PATH = REPO_ROOT / "templates" / "theme.yaml"

DEFAULT_THEME = {
    "fonts": {"latin": "Aptos", "cjk": "Microsoft YaHei", "fallback": "Arial"},
    "colors": {
        "background": "FFFFFF", "text": "1F2937", "border": "64748B",
        "input": "DDEBF7", "process": "E2F0D9", "decision": "FCE4D6",
        "data": "E4DFEC", "output": "FFE699", "accent": "2F75B5", "connector": "475569",
    },
    "layout": {"margin_in": 0.45, "gap_in": 0.35, "node_width_in": 2.1, "node_height_in": 0.8},
}


def load_theme(path: Path | None = None) -> dict:
    """Load templates/theme.yaml, falling back to built-in defaults."""
    path = Path(path) if path else THEME_PATH
    if not path.exists():
        return DEFAULT_THEME
    try:
        import yaml
    except ImportError:
        return DEFAULT_THEME
    with path.open(encoding="utf-8") as fh:
        data = yaml.safe_load(fh) or {}
    merged = {k: dict(v) if isinstance(v, dict) else v for k, v in DEFAULT_THEME.items()}
    for key, value in data.items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            merged[key].update(value)
        else:
            merged[key] = value
    return merged


def hex_to_rgb(value) -> RGBColor:
    """Accept 'DDEBF7', '#DDEBF7' or a YAML-parsed int (theme.yaml without quotes)."""
    if isinstance(value, int):
        value = f"{value:06X}"
    value = str(value).lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


# --------------------------------------------------------------------------- #
# fonts
# --------------------------------------------------------------------------- #
def set_font(run, latin: str, cjk: str, size: float = 18, bold: bool = False,
             color: str | None = None) -> None:
    """Write text colour and latin / east-asian / complex-script typefaces.

    DrawingML requires the order
        a:ln -> fill -> effectLst -> highlight -> latin -> ea -> cs
    so elements are inserted, never appended blindly: a wrong order makes
    PowerPoint report the file as corrupt.

    `color` matters more than it looks. An autoshape created by python-pptx
    carries `<p:style><a:fontRef idx="minor"><a:schemeClr val="lt1"/>`, i.e. the
    *light* theme colour, so unset text renders **white**. On the light module
    fills used by scientific figures that is invisible text -- a silent failure
    that only shows up when the file is opened. Always pass an explicit colour.
    """
    run.font.size = Pt(size)
    run.font.bold = bold
    rPr = run.font._rPr

    if color is not None:
        rPr.insert_element_before(
            parse_xml(f'<a:solidFill xmlns:a="{A_NS}"><a:srgbClr val="{str(color).lstrip("#")}"/></a:solidFill>'),
            "a:effectLst", "a:highlight", "a:uLnTx", "a:uLn", "a:uFillTx", "a:uFill",
            "a:latin", "a:ea", "a:cs", "a:sym",
            "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst",
        )

    for tag in ("a:latin", "a:cs"):
        el = rPr.find(f"{{{A_NS}}}{tag.split(':')[1]}")
        if el is None:
            el = rPr.makeelement(f"{{{A_NS}}}{tag.split(':')[1]}", {})
            rPr.insert_element_before(
                el, "a:ea", "a:cs", "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst"
            )
        el.set("typeface", latin)

    ea = rPr.find(f"{{{A_NS}}}ea")
    if ea is None:
        ea = rPr.makeelement(f"{{{A_NS}}}ea", {})
        rPr.insert_element_before(
            ea, "a:cs", "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst"
        )
    ea.set("typeface", cjk)


def set_text(shape, text: str, theme: dict, size: float = 18, bold: bool = False,
             color: str | None = None) -> None:
    tf = shape.text_frame
    tf.clear()
    run = tf.paragraphs[0].add_run()
    run.text = text
    set_font(
        run, theme["fonts"]["latin"], theme["fonts"]["cjk"], size=size, bold=bold,
        color=color or theme["colors"]["text"],
    )


# --------------------------------------------------------------------------- #
# connectors
# --------------------------------------------------------------------------- #
def add_arrow(slide, x1_in: float, y1_in: float, x2_in: float, y2_in: float,
              color: str = "475569", width_pt: float = 1.5, bidirectional: bool = False):
    """Add a native connector that really renders an arrowhead."""
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x1_in), Inches(y1_in), Inches(x2_in), Inches(y2_in)
    )
    conn.line.color.rgb = hex_to_rgb(color)   # must set colour first: writes a:solidFill
    conn.line.width = Pt(width_pt)
    ln = conn.line._get_or_add_ln()
    # a:tailEnd must come AFTER the fill element
    ln.append(parse_xml(f'<a:tailEnd xmlns:a="{A_NS}" type="arrow" w="med" len="med"/>'))
    if bidirectional:
        ln.insert(0, parse_xml(f'<a:headEnd xmlns:a="{A_NS}" type="arrow" w="med" len="med"/>'))
    return conn


def add_node(slide, x_in: float, y_in: float, label: str, fill: str, theme: dict,
             width_in: float | None = None, height_in: float | None = None):
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN

    w = width_in if width_in is not None else theme["layout"]["node_width_in"]
    h = height_in if height_in is not None else theme["layout"]["node_height_in"]
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_in), Inches(y_in), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(fill)
    shape.line.color.rgb = hex_to_rgb(theme["colors"]["border"])
    set_text(shape, label, theme, size=14, bold=True)
    shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    return shape


def new_presentation(theme: dict):
    """Create a 16:9 presentation matching the theme page size."""
    from pptx import Presentation

    prs = Presentation()
    page = DEFAULT_THEME  # page size lives in theme.yaml under `page`
    try:
        import yaml
        if THEME_PATH.exists():
            with THEME_PATH.open(encoding="utf-8") as fh:
                page = (yaml.safe_load(fh) or {}).get("page", page)
    except ImportError:
        pass
    prs.slide_width = Inches(page.get("width_in", 13.333))
    prs.slide_height = Inches(page.get("height_in", 7.5))
    return prs, prs.slides.add_slide(prs.slide_layouts[6])
