# CLI Guide

RusDox renders document specs into `.docx` and `.pdf`, validates them before render, can rebuild while you edit, and can benchmark the full pipeline.

## Most Common Commands

Render one document:

```bash
rusdox mydoc.yaml
```

Render every spec in a folder:

```bash
rusdox examples
```

Create a starter YAML document:

```bash
rusdox init-doc mydoc.yaml
```

Validate a file before render:

```bash
rusdox validate mydoc.yaml
```

Watch a file and rebuild on change:

```bash
rusdox watch mydoc.yaml
```

Benchmark a spec:

```bash
rusdox bench mydoc.yaml --iterations 5 --warmup 1
```

## Validation

Check one file:

```bash
rusdox validate mydoc.yaml
```

Check every spec in a folder and emit JSON for CI:

```bash
rusdox validate examples --format json
```

Validation catches semantic issues before render, including invalid colors, unknown style references, table shape mismatches, blank required values, and missing visual assets.

Render commands also run the same semantic validation before they write output.

## Watch Mode

Watch one file with the default poller:

```bash
rusdox watch mydoc.yaml
```

Watch a file with faster polling and DOCX-only output:

```bash
rusdox watch mydoc.yaml --docx-only --poll-interval-ms 250
```

Watch a whole folder and keep PDF generation enabled:

```bash
rusdox watch examples --with-pdf
```

Use `--max-builds 2` or another small number when you want the watcher to stop automatically, which is useful for tests and scripted workflows.

RusDox watches the spec input plus the active config path. Without `--config`, it tracks `./rusdox.toml` and the user config fallback automatically.

## Benchmarking

Benchmark one spec:

```bash
rusdox bench mydoc.yaml --iterations 5 --warmup 1
```

Benchmark a folder and emit machine-readable output:

```bash
rusdox bench examples --format json
```

Keep the generated artifacts instead of using a temporary output workspace:

```bash
rusdox bench mydoc.yaml --keep-output
```

Bench reports parse, validation, compose, DOCX write, PDF render, total runtime, and output byte sizes.

## Output Control

Write DOCX only:

```bash
rusdox mydoc.yaml --docx-only
```

Force PDF generation even if config disables it:

```bash
rusdox mydoc.yaml --with-pdf
```

Write a single input file to an explicit DOCX path:

```bash
rusdox mydoc.yaml --output ./out/custom-name.docx
```

## Config Commands

Create a config:

```bash
rusdox config init --template
```

Launch the simple wizard:

```bash
rusdox config wizard --level basic
```

Launch the full wizard:

```bash
rusdox config wizard --level advanced
```

Print the active user config path:

```bash
rusdox config path
```

Print the effective config:

```bash
rusdox config show
```

## Local Project Config

Create a local override:

```bash
rusdox config wizard --path ./rusdox.toml --level basic
```

That file overrides `~/rusdox/config.toml` for the current project.

## Other Supported Spec Formats

YAML is the recommended format, but these also work:

- `.yml`
- `.json`
- `.toml`

## Advanced Script Mode

RusDox also supports a `.rs` entrypoint for advanced workflows:

```bash
rusdox init-script mydoc.rs
rusdox mydoc.rs
```

This is useful when you need loops, conditional logic, API calls, or generated content that would be awkward in YAML.
