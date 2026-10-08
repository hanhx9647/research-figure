"""Shared helpers for the test suite."""
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "scripts"))
sys.path.insert(0, str(REPO_ROOT))

from figure_kit import add_arrow, add_node, load_theme, new_presentation, set_text  # noqa: E402
from validate_pptx import validate  # noqa: E402

OUTPUT = REPO_ROOT / "output" / "tests"


def build_chain(chain_spec, out_path: Path, feedback=None):
    """Build a left-to-right workflow from (label, fill) pairs.

    feedback: optional (i, j, bidirectional) adding an arrow from node j back to node i.
    """
    theme = load_theme()
    prs, slide = new_presentation(theme)
    gap = 3.05
    nodes = [
        add_node(slide, 0.7 + k * gap, 3.0, label, fill, theme)
        for k, (label, fill) in enumerate(chain_spec)
    ]
    for left, right in zip(nodes, nodes[1:]):
        y = (left.top + left.height // 2) / 914400
        add_arrow(slide, (left.left + left.width) / 914400, y, right.left / 914400, y)
    if feedback:
        i, j, bidirectional = feedback
        top_i = nodes[i].top / 914400
        add_arrow(
            slide,
            (nodes[j].left + nodes[j].width / 2) / 914400, nodes[j].top / 914400,
            (nodes[i].left + nodes[i].width / 2) / 914400, top_i,
            bidirectional=bidirectional,
        )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)
    return out_path


def assert_editable(report, min_shapes=4, min_connectors=2, min_text=3):
    assert report["slides"] >= 1
    assert report["native_shapes"] >= min_shapes, f"too few native shapes: {report}"
    assert report["connectors"] >= min_connectors, f"too few connectors: {report}"
    assert report["arrowed_connectors"] == report["connectors"], (
        f"some connectors lack arrowheads: {report}"
    )
    assert report["text_objects"] >= min_text, f"too few text objects: {report}"
    assert not any("screenshot" in w for w in report["warnings"]), report["warnings"]
    assert not any("overlap" in w for w in report["warnings"]), report["warnings"]
    assert not any("overflow" in w for w in report["warnings"]), report["warnings"]
    assert not any("low text contrast" in w for w in report["warnings"]), report["warnings"]
    assert not any("a:ea" in w for w in report["warnings"]), report["warnings"]
