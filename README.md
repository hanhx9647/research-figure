# ResearchFigureGPT

> **This skill has moved** into the multi-skill collection **Research Artifact Lab**:
> `github.com/hanhx9647/research-artifact-skills` → `skills/research-figure/`.
> This standalone repository is kept for history and receives bug fixes only.

ResearchFigureGPT is a confirmation-first workflow for turning scientific logic into editable PowerPoint figures for researchers who may not program.

**AI Scientific Figure Architect** — a reusable workflow that converts scientific research logic into editable PowerPoint figures.

Pipeline:

Scientific content
-> Figure logic extraction
-> Confirmation draft (approved gate)
-> Visual architecture
-> PowerPoint object tree
-> python-pptx generation
-> Editable PPT figure

Features:

- Nature/Science style workflow design
- Native PowerPoint shapes, independent textboxes, real arrowed connectors
- No SVG as primary output, no screenshots replacing the workflow
- Confirmation-first: no final figure before the logic is approved
- Image-to-PPTX conversion with an explicit editability report
- Chinese and English labels with proper CJK font handling

## Quick start

```powershell
python -m pip install -r requirements.txt

# 1. From-scratch workflow figure (native shapes + arrowed connectors)
python examples/simple_workflow.py
python scripts/validate_pptx.py output/simple_workflow.pptx

# 2. Image -> editable PPTX
python examples/image_to_pptx.py

# 3. Full test suite (Chinese / English / image conversion)
python tests/run_tests.py
```

PowerPoint is **not** required to generate the PPTX. PNG/PDF previews depend on
locally available PowerPoint or LibreOffice — if neither exists,
`scripts/render_preview.py` says so explicitly instead of failing silently.

## Confirmation-first

For papers or research descriptions, the skill first outputs a seven-part
confirmation draft (goal, stages, inputs/outputs, relations, feedback loops,
ambiguities, questions) and only generates after approval
(`references/confirmation-template.md`). For existing figures, it proposes which
regions can be rebuilt and which must remain embedded; every image conversion
ships a "native rebuild regions" and "embedded picture regions" list
(`references/image-to-pptx.md`).

## Repository layout

```text
SKILL.md                        # agent entrypoint
references/                     # authoritative contracts and guidelines
  workflow-contract.md          #   state schema + approval gate
  confirmation-template.md      #   seven-part confirmation draft
  image-to-pptx.md              #   region classification + editability report
  reference-image-analysis.md   #   style-reference analysis (style ≠ logic)
  visual-guidelines.md          #   layout, colour, accessibility
  validation-rules.md           #   editability + visual QA rules
  reference-styles/             #   your style reference images (optional)
templates/theme.yaml            # default reusable theme
scripts/
  figure_kit.py                 # shapes / fonts / arrowed connectors (trap-free)
  convert_image_to_pptx.py      # image -> editable PPTX + report
  render_preview.py             # PNG/PDF previews (PowerPoint COM / LibreOffice)
  validate_pptx.py              # structural validator
examples/
  simple_workflow.py            # from-scratch example
  image_to_pptx.py              # image conversion example (+ input & expected output)
tests/                          # real generation tests, no PowerPoint needed
```

## License

**Custom Non-Commercial Attribution License** — not an OSI open-source license.
Use, copy, modify, and adapt for personal, academic, educational, and other
non-commercial purposes with credit to the original author **左手骑车**. No
commercial use; modified copies may not be redistributed without prior written
permission. See `LICENSE`.
