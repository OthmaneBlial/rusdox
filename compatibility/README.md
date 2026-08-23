# Compatibility evidence

This directory contains the small, public, non-sensitive artifacts used by the dated viewer scorecard.

- `fixtures/dual-output-contract.docx` and `.pdf` are generated from `examples/dual_output_contract.yaml` by `scripts/generate_compatibility_fixtures.sh`.
- `evidence/apple-quicklook-first-page.png` is the first-page thumbnail emitted by macOS Quick Look from the checked-in PDF fixture.
- `viewer-results.json` is the machine-readable record behind `docs/compatibility-scorecard.md`.

Fixture SHA-256 values for the 2026-08-24 run:

```text
5737623721ed3bf0c79b87f803bd5311e86dfffac691af5d8ef1b91cfac5aa52  fixtures/dual-output-contract.docx
39f5f283195fddadc223e444972490d096c6c637d68507dcb63ca050a24eea25  fixtures/dual-output-contract.pdf
83b22ac3b0cfec9029b4fea4f39a5464affeb0b454afcf2733c1074ec730ec2f  evidence/apple-quicklook-first-page.png
```

An application-open smoke test proves that a viewer accepted the file; it does not by itself prove pixel-perfect fidelity. The scorecard keeps those claims separate.
