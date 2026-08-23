# Rendered-page golden baselines

`linux-x86_64/` contains the exact deterministic geometry snapshots produced by the Ubuntu CI renderer for every top-level example.

CI compares new pages at a zero-pixel threshold. A baseline update therefore requires an intentional renderer or fixture change and reviewer inspection; it is not a way to hide a regression.

To refresh from a trusted Ubuntu parity artifact:

```bash
./scripts/update_visual_baselines.sh /path/to/artifact/reports
```

The baselines are platform-scoped because installed font files and fallback selection can legitimately differ between operating systems. Semantic DOCX/PDF checks remain cross-platform.
