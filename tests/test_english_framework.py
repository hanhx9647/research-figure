"""English research framework test: generate -> validate -> assert editability."""
import unittest

from _common import OUTPUT, assert_editable, build_chain
from validate_pptx import validate


class EnglishFrameworkTest(unittest.TestCase):
    def test_english_framework(self):
        out = build_chain(
            [
                ("Input data", "DDEBF7"),
                ("Preprocessing", "E2F0D9"),
                ("Model training", "FCE4D6"),
                ("Evaluation", "E4DFEC"),
                ("Validated output", "FFE699"),
            ],
            OUTPUT / "test_english_framework.pptx",
        )
        self.assertTrue(out.exists())

        report = validate(out)
        assert_editable(report, min_shapes=5, min_connectors=4, min_text=5)
        self.assertEqual(report["pictures"], 0, "a pure workflow figure must contain no pictures")


if __name__ == "__main__":
    unittest.main()
