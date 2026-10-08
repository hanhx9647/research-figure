# python-pptx: known traps

Do not re-derive these — use `scripts/figure_kit.py`, which handles all of them.

## 1. Arrowheads do not exist in the API

`connector.line.end_arrowhead = True` does **not** raise and does **not** work:
`LineFormat` has no such attribute (upstream issue #375), and the class has no
`__slots__`, so the assignment silently creates a dead attribute. The PPTX ends
up with arrow-less connectors.

Fix — append the XML directly (after setting line colour):

```python
from pptx.oxml import parse_xml
ln = connector.line._get_or_add_ln()
ln.append(parse_xml('<a:tailEnd xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" type="arrow" w="med" len="med"/>'))
```

## 2. `font.name` does not cover Chinese

`run.font.name = "Microsoft YaHei"` writes only `a:latin`. Chinese glyphs are
rendered with the `a:ea` (East Asian) typeface and silently fall back to a
default font. `a:ea` must be written too, and the element order in `rPr` must be
`a:latin -> a:ea -> a:cs`, or PowerPoint reports the file as corrupt.

Use `figure_kit.set_font()`.

## 3. YAML colours parse as integers

`connector: 475569` in YAML becomes the *integer* 475569. Quote colour values in
`theme.yaml`, and make colour helpers accept both (see `figure_kit.hex_to_rgb`).

## 4. Element order inside a:ln

`a:tailEnd` must come after the fill element; set the line colour first, then
append the arrowhead element.

## Core APIs

Shape: `add_shape()` · Text: `add_textbox()` · Connector: `add_connector()`
