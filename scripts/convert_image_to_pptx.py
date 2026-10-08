"""Convert an existing figure image into an editable PPTX.

This script does NOT pretend to be a perfect reverse-engineering tool.
It is driven by an explicit region description (YAML) so that the result is
auditable: every region is declared as `rebuild`, `embed`, `drop` or `uncertain`.

    python scripts/convert_image_to_pptx.py --image examples/image_to_pptx_input.png \
        --regions examples/image_to_pptx_regions.yaml

Without a region file, use --auto to produce a DRAFT region file whose regions are
all marked `uncertain`; fill it in and re-run. Auto-detection never guesses
scientific content.

Region file format:

    image: examples/image_to_pptx_input.png
    regions:
      - id: r1
        type: rebuild          # rebuild | embed | drop | uncertain
        kind: node             # node | arrow | text | title | legend
        bbox: [40, 30, 300, 110]
        text: "原始数据"
        fill: "DDEBF7"
        ocr_confidence: 0.96
      - id: r2
        type: rebuild
        kind: arrow
        from: r1
        to: r3
      - id: r5
        type: embed
        bbox: [640, 300, 900, 520]
        reason: "SEM micrograph; rebuilding would destroy fidelity"
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from figure_kit import (  # noqa: E402
    add_arrow, add_node, hex_to_rgb, load_theme, new_presentation, set_font, set_text,
)
from pptx.util import Inches  # noqa: E402

try:
    import yaml
except ImportError:
    raise SystemExit("PyYAML is required. Install with: python -m pip install -r requirements.txt")

try:
    from PIL import Image
except ImportError:
    raise SystemExit("Pillow is required. Install with: python -m pip install -r requirements.txt")

CANVAS_W, CANVAS_H = 13.333, 7.5
MARGIN = 0.45


# --------------------------------------------------------------------------- #
# auto-detection (draft only)
# --------------------------------------------------------------------------- #
def detect_candidates(image_path: Path, min_area: int = 2500) -> list[dict]:
    """Projection-based block detection. Produces a DRAFT only.

    Splits the non-white area into horizontal bands, then splits each band on
    vertical gaps. Every returned region is `uncertain` -- a human must decide.
    """
    im = Image.open(image_path).convert("L")
    w, h = im.size
    mask = im.point(lambda v: 255 if v < 240 else 0)
    data = mask.tobytes()

    row_ink = [any(data[y * w:(y + 1) * w]) for y in range(h)]
    boxes: list[list[int]] = []

    for y0, y1 in _runs(row_ink):
        if y1 - y0 < 8:
            continue
        col_ink = [False] * w
        for y in range(y0, y1):
            row = data[y * w:(y + 1) * w]
            for x in range(w):
                if row[x]:
                    col_ink[x] = True
        for x0, x1 in _runs(col_ink):
            if (x1 - x0) * (y1 - y0) < min_area:
                continue
            boxes.append([x0, y0, x1, y1])

    return [
        {
            "id": f"r{i}",
            "type": "uncertain",
            "kind": "node",
            "bbox": b,
            "text": "",
            "ocr_confidence": 0.0,
            "note": "auto-detected draft; classify as rebuild/embed/drop and fill in text",
        }
        for i, b in enumerate(boxes, start=1)
    ]


def _runs(flags: list[bool], min_len: int = 1):
    runs, start = [], None
    for i, flag in enumerate(flags):
        if flag and start is None:
            start = i
        elif not flag and start is not None:
            if i - start >= min_len:
                runs.append((start, i))
            start = None
    if start is not None and len(flags) - start >= min_len:
        runs.append((start, len(flags)))
    return runs


# --------------------------------------------------------------------------- #
# audit
# --------------------------------------------------------------------------- #
def audit(regions: list[dict]) -> list[dict]:
    """Return a list of blocking problems. Empty list means safe to generate."""
    problems = []
    ids = {r.get("id") for r in regions}

    for r in regions:
        rid = r.get("id", "<no id>")
        rtype = r.get("type")

        if rtype not in ("rebuild", "embed", "drop", "uncertain"):
            problems.append({"id": rid, "reason": f"unknown type '{rtype}'"})

        bbox = r.get("bbox")
        needs_bbox = rtype in ("rebuild", "embed", "drop") and r.get("kind", "node") != "arrow"
        if needs_bbox and (not bbox or len(bbox) != 4):
            problems.append({"id": rid, "reason": "missing or malformed bbox"})

        if rtype == "rebuild":
            kind = r.get("kind", "node")
            if kind in ("node", "text", "title", "legend") and not r.get("text"):
                problems.append({"id": rid, "reason": "rebuild region without text (OCR failed?)"})
            if kind == "arrow":
                for side in ("from", "to"):
                    if r.get(side) not in ids:
                        problems.append({"id": rid, "reason": f"arrow endpoint '{side}' does not reference a known id"})
            conf = r.get("ocr_confidence", 1.0)
            if conf < 0.8:
                problems.append({"id": rid, "reason": f"OCR confidence {conf} < 0.8; needs confirmation"})
        elif rtype == "embed" and not r.get("reason"):
            problems.append({"id": rid, "reason": "embed region without a reason"})

    return problems


# --------------------------------------------------------------------------- #
# geometry
# --------------------------------------------------------------------------- #
def placement(img_w: int, img_h: int):
    """Fit the image into the canvas; return (scale, offset_x_in, offset_y_in)."""
    avail_w, avail_h = CANVAS_W - 2 * MARGIN, CANVAS_H - 2 * MARGIN
    scale = min(avail_w / img_w, avail_h / img_h)
    off_x = (CANVAS_W - img_w * scale) / 2
    off_y = (CANVAS_H - img_h * scale) / 2
    return scale, off_x, off_y


def bbox_to_inches(bbox, scale, off_x, off_y):
    x0, y0, x1, y1 = bbox
    return (off_x + x0 * scale, off_y + y0 * scale,
            off_x + x1 * scale, off_y + y1 * scale)


# --------------------------------------------------------------------------- #
# build
# --------------------------------------------------------------------------- #
def build(image_path: Path, regions: list[dict], out_pptx: Path, theme: dict) -> dict:
    from pptx.enum.shapes import MSO_SHAPE

    img = Image.open(image_path)
    img_w, img_h = img.size
    scale, off_x, off_y = placement(img_w, img_h)

    prs, slide = new_presentation(theme)
    crops_dir = out_pptx.parent / "_crops"
    crops_dir.mkdir(parents=True, exist_ok=True)

    by_id = {r["id"]: r for r in regions}
    anchors: dict[str, tuple[float, float, float, float]] = {}

    # 1) nodes first, so arrows can reference their edges
    for r in regions:
        if r.get("type") != "rebuild" or r.get("kind", "node") == "arrow":
            continue
        x0i, y0i, x1i, y1i = bbox_to_inches(r["bbox"], scale, off_x, off_y)
        fill = r.get("fill", theme["colors"]["process"])
        shape = add_node(slide, x0i, y0i, r.get("text", ""), fill, theme,
                         width_in=max(x1i - x0i, 0.6), height_in=max(y1i - y0i, 0.4))
        anchors[r["id"]] = (x0i, y0i, x1i, y1i)
        _ = shape

    # 2) arrows
    for r in regions:
        if r.get("type") != "rebuild" or r.get("kind") != "arrow":
            continue
        src, dst = anchors.get(r.get("from")), anchors.get(r.get("to"))
        if not src or not dst:
            continue
        y = (src[1] + src[3]) / 2
        add_arrow(slide, src[2], y, dst[0], y,
                  color=theme["colors"]["connector"],
                  bidirectional=bool(r.get("bidirectional")))

    # 3) embedded regions (cropped from the original, keeping resolution)
    for r in regions:
        if r.get("type") not in ("embed", "uncertain"):
            continue
        x0i, y0i, x1i, y1i = bbox_to_inches(r["bbox"], scale, off_x, off_y)
        crop_path = crops_dir / f"{out_pptx.stem}_{r['id']}.png"
        img.crop(tuple(r["bbox"])).save(crop_path)
        slide.shapes.add_picture(str(crop_path), Inches(x0i), Inches(y0i),
                                 Inches(max(x1i - x0i, 0.2)), Inches(max(y1i - y0i, 0.2)))

    out_pptx.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_pptx)

    rebuild = [r for r in regions if r.get("type") == "rebuild"]
    embed = [r for r in regions if r.get("type") in ("embed", "uncertain")]
    dropped = [r for r in regions if r.get("type") == "drop"]
    total = len(rebuild) + len(embed) or 1
    score = len(rebuild) / total

    if score >= 0.8:
        verdict = "高度可编辑"
    elif score >= 0.5:
        verdict = "部分可编辑"
    else:
        verdict = "主要为嵌入图片，建议改用 re-draw 模式"

    return {
        "image": str(image_path), "image_size": [img_w, img_h],
        "rebuild": rebuild, "embed": embed, "drop": dropped,
        "score": score, "verdict": verdict, "pptx": out_pptx,
    }


def write_report(result: dict, report_path: Path) -> None:
    lines = [
        "# 图片转 PPT 编辑性报告",
        "",
        f"- 原图：`{result['image']}`（{result['image_size'][0]}×{result['image_size'][1]} px）",
        f"- 输出：`{result['pptx']}`",
        f"- 区域总数：{len(result['rebuild']) + len(result['embed']) + len(result['drop'])}",
        f"- 可编辑度：{result['score']:.0%} —— **{result['verdict']}**",
        "",
        f"> 本图共 {len(result['rebuild']) + len(result['embed'])} 个区域，"
        f"其中 {len(result['rebuild'])} 个已重建为可编辑形状，{len(result['embed'])} 个仍为嵌入图片。",
        "",
        "## 原生重建区域（可自由编辑）",
        "",
        "| id | 类型 | 文字 |",
        "|---|---|---|",
    ]
    lines += [f"| {r['id']} | {r.get('kind', 'node')} | {r.get('text', '')} |" for r in result["rebuild"]] \
        or ["| — | — | 无 |"]

    lines += ["", "## 嵌入图片区域（不可编辑）", "", "| id | 原因 |", "|---|---|"]
    lines += [f"| {r['id']} | {r.get('reason') or r.get('note', '未分类，默认嵌入')} |" for r in result["embed"]] \
        or ["| — | 无 |"]

    lines += ["", "## 丢弃区域", "", "| id | 原因 |", "|---|---|"]
    lines += [f"| {r['id']} | {r.get('reason', '')} |" for r in result["drop"]] or ["| — | 无 |"]

    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


# --------------------------------------------------------------------------- #
# entrypoint
# --------------------------------------------------------------------------- #
def main() -> None:
    ap = argparse.ArgumentParser(description="Convert a figure image into an editable PPTX.")
    ap.add_argument("--image", type=Path, required=True)
    ap.add_argument("--regions", type=Path, default=None, help="YAML region description")
    ap.add_argument("--out", type=Path, default=None, help="Output PPTX path")
    ap.add_argument("--auto", action="store_true", help="Write a DRAFT region file and exit")
    ap.add_argument("--strict", action="store_true", help="Refuse to generate if any region is uncertain")
    args = ap.parse_args()

    if not args.image.exists():
        raise SystemExit(f"Image not found: {args.image}")

    if args.auto:
        candidates = detect_candidates(args.image)
        draft = args.regions or args.image.with_suffix(".regions.yaml")
        draft.write_text(
            yaml.safe_dump({"image": str(args.image), "regions": candidates},
                           allow_unicode=True, sort_keys=False),
            encoding="utf-8",
        )
        print(f"[auto] wrote DRAFT with {len(candidates)} uncertain regions -> {draft}")
        print("[auto] classify every region (rebuild/embed/drop) and fill in text, then re-run without --auto.")
        return

    if not args.regions or not args.regions.exists():
        raise SystemExit("A region file is required (or pass --auto to generate a draft).")

    spec = yaml.safe_load(args.regions.read_text(encoding="utf-8")) or {}
    regions = spec.get("regions", [])
    if not regions:
        raise SystemExit("Region file contains no regions.")

    problems = audit(regions)
    if problems:
        print("[audit] blocking issues, NOT generating a final PPTX:")
        for p in problems:
            print(f"  - {p['id']}: {p['reason']}")
        if args.strict or any("confidence" in p["reason"] or "uncertain" in str(p) for p in problems):
            print("\nResolve these in the region file (or answer the questions) and re-run.")
            raise SystemExit(2)

    out = args.out or (Path("output") / f"{args.image.stem}_converted.pptx")
    theme = load_theme()
    result = build(args.image, regions, out, theme)
    report = out.with_name(f"{out.stem}_editability.md")
    write_report(result, report)

    print(f"[ok] PPTX   -> {out}")
    print(f"[ok] report -> {report}")
    print(f"[ok] editability {result['score']:.0%} ({result['verdict']})")


if __name__ == "__main__":
    main()
