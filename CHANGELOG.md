# Changelog

All notable changes to ResearchFigureGPT are documented here.
This project does **not** follow SemVer strictly while pre-1.0; breaking
behaviour changes are noted in each entry.

> **Note (2026-10-08):** this skill now lives inside the multi-skill collection
> **Research Artifact Lab** at `github.com/hanhx9647/research-artifact-skills`
> (`skills/research-figure/`). This standalone repository receives bug fixes only.

## v0.1.1 — 2026-10-08

### Fixed

- **Invisible text (silent).** Text colour was never written, so a shape inherited
  `<a:fontRef idx="minor"><a:schemeClr val="lt1"/>` — the *light* theme colour — and every
  label rendered **white** on the light module fills. No exception was raised anywhere during
  generation; the defect only appears when the file is opened.
  `figure_kit.set_font()` now accepts `color` and `set_text()` writes the theme text colour.

### Added

- `scripts/validate_pptx.py`: text/fill contrast check — anything below 2.5:1 is reported
  (the pre-fix example files measure 1.2:1).
- `tests/test_silent_failures.py`: regression tests for all three silent failure modes
  (missing arrowhead, missing `a:ea`, unset text colour).
- `tests/_common.py`: `assert_editable` now also fails on low-contrast and missing `a:ea`.
- `references/visual-guidelines.md`: explicit "always set text colour" and feedback-loop
  routing rules.

## v0.1.0 — 2026-10-08

Initial public-ready structure.

### Added

- Confirmation-first pipeline: seven-part confirmation draft with a hard
  `approved` gate (`references/confirmation-template.md`); no final PPTX is
  produced while `approved: false`.
- Image-to-PPTX conversion flow: declarative region classification
  (`rebuild` / `embed` / `drop` / `uncertain`), OCR confidence rules, and a
  mandatory editability report (`references/image-to-pptx.md`,
  `scripts/convert_image_to_pptx.py`).
- Reference-image analysis contract: ten analysis dimensions, explicit split
  between visual style and scientific content (`references/reference-image-analysis.md`).
- `scripts/render_preview.py`: PNG/PDF previews via PowerPoint COM or
  LibreOffice, with an explicit notice when no renderer is available.
- `scripts/figure_kit.py`: shared primitives that correctly handle connector
  arrowheads and CJK (a:ea) fonts.
- `scripts/validate_pptx.py`: arrowhead, CJK font, text-overflow, module-overlap
  and full-slide-screenshot checks.
- Test suite covering Chinese workflow, English framework, and image conversion
  (`tests/`); all runnable without PowerPoint.
- `.gitignore`, `CHANGELOG.md`, example region file, sample input image, and an
  expected-output guide for the image example.

### Fixed

- Example generated connectors that silently had **no arrowheads**
  (`line.end_arrowhead` is not a real python-pptx API; upstream issue #375).
- Chinese labels silently falling back to the default CJK font (`a:ea` was
  never written).
- YAML colour values in `theme.yaml` parsed as integers; now quoted and parsed
  defensively.
- Duplicate/stale top-level docs (`PROMPTS.md`, `WORKFLOW.md`, `VALIDATION.md`,
  `PYTHON_PPTX_GUIDE.md`) now point to the authoritative `references/` documents.

### Known limitations

- Image conversion is region-file driven; automatic detection only produces a
  draft for human classification.
- PNG/PDF previews require PowerPoint or LibreOffice; the skill states this
  explicitly instead of pretending previews exist.
