# Validation

- Rules: `references/validation-rules.md`
- Runnable checker: `python scripts/validate_pptx.py path/to/file.pptx`
- Test suite: `python tests/run_tests.py`

The checker reports native shapes, arrowed connectors, pictures, and warns about
missing arrowheads, missing CJK font fallback, text overflow, overlapping modules,
and full-slide screenshots.
