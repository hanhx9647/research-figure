# Workflow contract

Classify input as `existing_figure`, `research_text`, or `mixed`. A reference image informs visual language unless the user confirms its scientific content.

Before generation maintain `output/figure_contract.yaml`:

```yaml
purpose: null
audience: null
orientation: left-to-right
stages: []
inputs: []
outputs: []
feedback_loops: []
assumptions: []
approved: false
decision_log: []
```

## Approval gate

Set `approved: true` only after the user confirms the logic and major labels.
While `approved` is `false`, do **not** write any final PPTX or final image to
`output/` — only the confirmation draft, the contract file itself, and draft
reports. The full draft template and the exact conditions for approval live in
`confirmation-template.md`.

For existing figures, add `conversion_notes` and `embedded_regions`; for image
conversion follow `image-to-pptx.md`.

Each node has `id`, `label`, `type`, and `position`; each edge has `source`,
`target`, and optional `label`.
