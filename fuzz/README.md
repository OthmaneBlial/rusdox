# RusDox fuzz targets

The targets cover the untrusted-input boundaries that can amplify CPU or memory:

- `docx_open`: ZIP container, relationships, XML, and DOCX body parsing under reduced limits;
- `spec_parse`: YAML, JSON, and TOML deserialization;
- `visual_render`: PNG, JPEG, and SVG inspection/rasterization through the real DOCX save path.

Build every harness:

```bash
cargo +nightly fuzz build
```

Run one locally with a bounded duration:

```bash
cargo +nightly fuzz run docx_open -- -max_total_time=60
```

Crashes and minimized inputs belong under `fuzz/artifacts/` and `fuzz/corpus/`; do not commit documents containing private customer data.
