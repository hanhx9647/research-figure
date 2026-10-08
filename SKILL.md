---
name: research-figure-gpt
description: Design and generate editable scientific workflow figures from research text, papers, sketches, reference images, or existing figures. Use a confirmation-first process and produce native PowerPoint, source code, previews, and a reusable theme configuration.
version: 0.1.0
author: 左手骑车
metadata:
  short-description: Create editable scientific workflow figures
---

# ResearchFigureGPT

Use this skill for publication-oriented scientific workflow, framework, method, and architecture figures. It is intended for ordinary researchers, including users who do not program.

## Interaction contract

- Existing figure, sketch, or image: inspect it, identify editable and non-editable regions, propose a conversion plan, then generate an editable PPTX where practical. Follow `references/image-to-pptx.md`.
- Paper, abstract, method description, or raw research notes: do not generate a final figure immediately. First extract the scientific entities, stages, dependencies, loops, inputs, outputs, and uncertain points; present the confirmation draft (template below) for approval. Ask for a standard reference figure when visual style is not supplied, or use the bundled default theme.
- Mixed input: separate confirmed scientific content from visual references and ask for confirmation only on ambiguous parts.
- Reference images: analyse style only, per `references/reference-image-analysis.md`. A reference image never proves scientific logic.

Before generation, confirm the figure purpose, audience, reading direction, main stages, labels, required images, and uncertain scientific relationships. Keep a short decision log so the user can revise one item without restarting the whole design.

## Approval gate (hard rule)

Output the confirmation draft, then set `approved: false` in the figure state.
**While `approved` is `false`, do not generate the final PPTX or final images.**
See `references/confirmation-template.md` for the draft template and the exact
conditions for setting `approved: true`.

## Image-to-PPTX rules

Conversion never guarantees a 100% automatic reconstruction. Every conversion must
declare each region as `rebuild`, `embed`, `drop`, or `uncertain`, and must output
a "native rebuild regions" list and an "embedded picture regions" list.
See `references/image-to-pptx.md`.

Never use:
- SVG as the primary output format
- screenshots or flattened figures standing in for the workflow
- raster images to replace rebuildable workflow structure

Photos, microscopy images, and complex scientific plots may remain embedded pictures when rebuilding them as shapes would damage fidelity.

## Deliverables

Provide a figure design explanation, editable `.pptx`, reproducible Python source, PNG/PDF previews **when a renderer is available** (see `scripts/render_preview.py`; state clearly when it is not), a reusable theme configuration, and a validation report.

## Visual and interoperability rules

- Prefer native PowerPoint shapes for modules, independent text boxes for labels, and real connectors with arrowheads for relationships.
- Use `scripts/figure_kit.py` for shape/text/connector primitives — it handles the two known python-pptx traps (missing arrowhead API, missing CJK `a:ea` font).
- Use a restrained, color-blind-aware palette and common fonts with Chinese fallback. Avoid effects that break in Google Slides.
- Do not silently invent scientific relationships. Mark assumptions and request confirmation.

## References

- `references/workflow-contract.md`: confirmation-first decision flow and object schema.
- `references/confirmation-template.md`: the seven-part confirmation draft and approval gate.
- `references/image-to-pptx.md`: image conversion flow, region classification, editability report.
- `references/reference-image-analysis.md`: how to analyse style reference figures.
- `references/visual-guidelines.md`: layout, typography, color, accessibility, and Google Slides constraints.
- `references/validation-rules.md`: editability and visual QA checks.
- `templates/theme.yaml`: default reusable theme.

For runnable baselines, use `examples/simple_workflow.py` (from scratch) and
`examples/image_to_pptx.py` (image conversion). For structural checking, use
`scripts/validate_pptx.py`; for the full test suite, use `tests/run_tests.py`.
