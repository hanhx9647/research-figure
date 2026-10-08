"""Minimal editable workflow example. Run from the repository root.

Uses figure_kit primitives, which handle the two classic python-pptx traps:
connector arrowheads and CJK font fallback.
"""
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "scripts"))

from figure_kit import add_arrow, add_node, load_theme, new_presentation  # noqa: E402
from pptx.util import Inches  # noqa: E402

OUT = REPO_ROOT / "output"
OUT.mkdir(exist_ok=True)


def main():
    theme = load_theme()
    prs, slide = new_presentation(theme)

    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(12.3), Inches(0.5))
    from figure_kit import set_text
    set_text(title, "Scientific workflow example 科研流程示例", theme, size=26, bold=True)

    nodes = [
        add_node(slide, 0.7 + i * 3.05, 3.0, text, fill, theme)
        for i, (text, fill) in enumerate([
            ("Input data 输入数据", theme["colors"]["input"]),
            ("Data processing 数据处理", theme["colors"]["process"]),
            ("Model training 模型训练", theme["colors"]["decision"]),
            ("Validated output 验证输出", theme["colors"]["output"]),
        ])
    ]
    for left, right in zip(nodes, nodes[1:]):
        y = (left.top + left.height // 2) / 914400
        add_arrow(slide, (left.left + left.width) / 914400, y, right.left / 914400, y)

    path = OUT / "simple_workflow.pptx"
    prs.save(path)
    print(path)


if __name__ == "__main__":
    main()
