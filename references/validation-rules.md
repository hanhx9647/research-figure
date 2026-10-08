# Validation rules

## Editability

- Every module is a native shape unless listed as an embedded region.
- Every label is an independent text box or inspectable shape text.
- Every relationship is a connector **with a rendered arrowhead**; a connector
  without `a:tailEnd`/`a:headEnd` fails.
- No full-slide screenshot replaces the workflow.
- CJK text must carry an explicit `a:ea` typeface (otherwise PowerPoint falls
  back to a default CJK font).
- Every image conversion ships an editability report listing native rebuild
  regions and embedded picture regions.

## Visual QA

- Check clipping, overlap, contrast, arrow direction, grayscale readability,
  and Google Slides portability.
- Estimated text height must fit the shape; flag likely overflow.

## How to check

`python scripts/validate_pptx.py path/to/file.pptx` runs the structural checks
above (arrowheads, CJK fonts, overflow, overlap, screenshot detection).

`python tests/run_tests.py` runs the generation test suite (Chinese workflow,
English framework, image conversion).
