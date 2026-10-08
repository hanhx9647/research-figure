# ResearchFigureGPT

**Turn research methods, sketches, and scientific logic into editable PowerPoint figures.**

ResearchFigureGPT is an open-source AI Skill for researchers, graduate students, and engineers who need publication-ready **editable scientific diagrams**—not another flat image. It converts a paper method, workflow description, hand-drawn sketch, or reference figure into a structured `.pptx` with native shapes, editable text, and real connectors.

> **ResearchFigureGPT is the first standalone Skill from [Research Artifact Lab](https://github.com/hanhx9647/research-artifact-skills).**

[![Install with Skills CLI](https://img.shields.io/badge/Skills%20CLI-install-blue)](https://github.com/agentskills/agentskills)
[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![PowerPoint](https://img.shields.io/badge/output-editable%20PPTX-B7472A?logo=microsoftpowerpoint&logoColor=white)](https://www.microsoft.com/microsoft-365/powerpoint)

## Why ResearchFigureGPT?

Most AI figure tools produce a PNG that researchers must redraw. ResearchFigureGPT is designed around the research deliverable:

- **Editable by design:** text, boxes, arrows, and layout elements are native PowerPoint objects.
- **Confirmation first:** the scientific logic is reviewed before final generation.
- **Image to PPTX:** rebuild selected regions as editable objects while clearly reporting embedded images.
- **Research-aware:** separates confirmed content, assumptions, ambiguities, and visual references.
- **Chinese and English:** supports CJK labels, scientific symbols, feedback loops, and multi-branch workflows.
- **Reproducible:** every example includes Python source, a theme, generated artifacts, and validation checks.

ResearchFigureGPT does not claim that every object is editable. Photos, microscopy images, and complex embedded artwork may remain pictures; the editability report identifies them explicitly.

## Install

Install the standalone Skill:

```bash
npx skills add hanhx9647/research-figure
```

Or install it from the collection:

```bash
npx skills add hanhx9647/research-artifact-skills --skill research-figure
```

For local generation, use Python 3.10+:

```bash
python -m pip install -r requirements.txt
```

PowerPoint is not required to generate `.pptx` files. PNG/PDF previews require PowerPoint or LibreOffice locally.

## How it works

```text
Research text / paper / sketch / reference image
                    ↓
        Scientific logic extraction
                    ↓
       Confirmation draft and approval gate
                    ↓
     Visual architecture and object planning
                    ↓
        Native PowerPoint generation
                    ↓
     PPTX + preview + source + validation report
```

For research descriptions, the Skill first identifies goals, stages, inputs, outputs, dependencies, feedback loops, and uncertain relationships. It generates the final figure only after approval. For image conversion, it classifies regions as native rebuilds or embedded pictures and records the result.

## Quick start

Generate a native workflow figure and validate its structure:

```bash
python examples/simple_workflow.py
python scripts/validate_pptx.py output/simple_workflow.pptx
```

Convert a reference image into an editable PPTX:

```bash
python examples/image_to_pptx.py
```

Run the complete test suite:

```bash
python tests/run_tests.py
```

## What you can create

- Scientific workflow and method diagrams
- Experimental pipelines and technical roadmaps
- Algorithm and system architecture figures
- Paper method sections converted to flowcharts
- Hand-drawn sketches rebuilt as editable PPTX
- Chinese-English research figures with feedback loops
- Figures containing both editable structure and clearly documented image regions

## Repository layout

```text
SKILL.md                 # Agent Skill entrypoint
references/              # workflow, approval, image conversion, and QA contracts
scripts/                 # generation, conversion, rendering, and validation tools
templates/theme.yaml     # reusable visual theme
examples/                # runnable examples and expected artifacts
tests/                   # regression and silent-failure tests
```

## Validation and limitations

The validator checks PPTX structure, native shapes, text objects, connectors, arrowheads, and picture regions. The project records checks that were not run, including human visual review and compatibility testing in applications that were not available in the local environment. Scientific correctness still requires researcher review.

## Documentation

- [Skill specification](SKILL.md)
- [Workflow contract](references/workflow-contract.md)
- [Confirmation template](references/confirmation-template.md)
- [Image-to-PPTX guide](references/image-to-pptx.md)
- [Validation rules](references/validation-rules.md)
- [中文说明](README.zh-CN.md)
- [Research Artifact Lab collection](https://github.com/hanhx9647/research-artifact-skills)

## License

This project is released under a custom non-commercial attribution license. Personal, academic, educational, and other non-commercial use is allowed with credit to the original author **左手骑车**. Commercial use and redistribution of modified copies require written permission. See [LICENSE](LICENSE).
