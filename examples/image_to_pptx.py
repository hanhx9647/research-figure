"""End-to-end minimal example: convert an existing figure image into an editable PPTX.

Run from the repository root:

    python examples/image_to_pptx.py

Pipeline:
    examples/image_to_pptx_input.png          (the original bitmap)
    examples/image_to_pptx_regions.yaml       (human-audited region description)
        -> output/image_to_pptx_converted.pptx
        -> output/image_to_pptx_converted_editability.md

What you get in the PPTX:
  * title, four workflow nodes, three arrowed connectors, one legend  -> NATIVE,
    fully editable shapes / text / connectors
  * the sample image thumbnail                                        -> EMBEDDED picture,
    cropped from the source, kept because rebuilding it would destroy fidelity

The split between "native" and "embedded" is declared in the region file and
restated in the editability report -- nothing here is guessed.
"""
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "scripts"))

import yaml  # noqa: E402

from convert_image_to_pptx import audit, build, write_report  # noqa: E402
from figure_kit import load_theme  # noqa: E402

IMAGE = REPO_ROOT / "examples" / "image_to_pptx_input.png"
REGIONS = REPO_ROOT / "examples" / "image_to_pptx_regions.yaml"
OUT_DIR = REPO_ROOT / "output"


def main() -> None:
    spec = yaml.safe_load(REGIONS.read_text(encoding="utf-8"))
    regions = spec["regions"]

    problems = audit(regions)
    if problems:
        print("[audit] blocking issues -- fix the region file first:")
        for p in problems:
            print(f"  - {p['id']}: {p['reason']}")
        raise SystemExit(2)

    theme = load_theme()
    out_pptx = OUT_DIR / "image_to_pptx_converted.pptx"
    result = build(IMAGE, regions, out_pptx, theme)
    write_report(result, out_pptx.with_name("image_to_pptx_converted_editability.md"))

    print(f"[ok] PPTX     -> {out_pptx}")
    print(f"[ok] report   -> {out_pptx.with_name('image_to_pptx_converted_editability.md')}")
    print(f"[ok] rebuild={len(result['rebuild'])}  embed={len(result['embed'])}  "
          f"drop={len(result['drop'])}  editability={result['score']:.0%} ({result['verdict']})")

    # Optional preview (PowerPoint COM or LibreOffice; skips silently if absent)
    render = REPO_ROOT / "scripts" / "render_preview.py"
    if render.exists():
        import subprocess
        subprocess.run([sys.executable, str(render), str(out_pptx)], check=False)


if __name__ == "__main__":
    main()
