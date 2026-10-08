# examples

Run from the repository root:

```powershell
python examples/simple_workflow.py    # from-scratch native-shape workflow
python examples/image_to_pptx.py      # image -> editable PPTX + editability report
```

- `simple_workflow.py` generates a native-shape, connector-with-arrowhead PPTX.
- `image_to_pptx.py` converts `image_to_pptx_input.png` using the human-audited
  regions in `image_to_pptx_regions.yaml`; see `image_to_pptx_expected.md` for
  what ends up editable vs embedded.
