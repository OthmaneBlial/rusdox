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

Rebuild on change with a local PDF/status dashboard:

```bash
rusdox dev mydoc.yaml --open
```

Benchmark a spec:

```bash
rusdox bench mydoc.yaml --iterations 5 --warmup 1
```

Generate DOCX, PDF, and parity evidence:

```bash
rusdox verify mydoc.yaml
```

Generate the authoring schema or migrate a legacy spec:

    rusdox schema --output rusdox-spec-v1.schema.json
    rusdox migrate legacy.yaml --in-place
    rusdox migrate current.yaml --check

## Parity Verification

Verify one file:

```bash
rusdox verify mydoc.yaml
```

Verify every top-level spec in a directory and write all artifacts under a CI workspace:

```bash
rusdox verify examples --output-root target/parity --format json
```

Add a deterministic rendered-page comparison:

```bash
rusdox verify mydoc.yaml \
  --visual-baseline tests/visual-baselines/mydoc \
  --visual-threshold 0.002
```

The command always writes versioned JSON and standalone HTML reports. Exit code `0` means parity passed, `1` means verification could not complete, and `2` means generated outputs failed one or more parity checks. See [Parity verification](parity.md) for the complete contract and visual-diff boundary.

## Validation

Check one file:

```bash
rusdox validate mydoc.yaml
```

Check every spec in a folder and emit JSON for CI:

```bash
rusdox validate examples --format json
```

Validation catches semantic issues before render, including unsupported spec versions, invalid colors, unknown style references, table shape mismatches, blank required values, and missing visual assets. File-backed semantic issues include one-based line and column coordinates in text and JSON reports.

Render commands also run the same semantic validation before they write output.

## Schema and migration

The schema command prints the generated version 1 JSON Schema. Use --output for
an atomic file write. The same object schema powers YAML, JSON, and TOML
authoring and the bundled VS Code extension.

The migrate command prints a migrated spec by default. Use --in-place for an
atomic replacement, --output for a separate destination, and --check in CI to
reject legacy unversioned specs. RusDox rejects future versions instead of
silently downgrading them. See [Spec Versioning](spec-versioning.md).

## Development feedback loop

Start the recommended development loop:

```bash
rusdox dev mydoc.yaml --open
```

The loop binds only to `127.0.0.1`. Its script-free dashboard refreshes to show
the latest successful PDF, validation or parsing failure, per-stage timings,
and absolute DOCX/PDF paths. A failed rebuild does not replace the previous
successful artifact, so the preview remains useful while you repair the spec.

Tune polling and debounce for editors that save in several writes:

```bash
rusdox dev mydoc.yaml --poll-interval-ms 100 --debounce-ms 250
```

Use JSON Lines for tooling, or quiet mode for a bounded CI smoke check:

```bash
rusdox dev mydoc.yaml --json --port 0
rusdox dev mydoc.yaml --quiet --port 0 --max-builds 1
```

Each JSON build event includes status, trigger reason, changed paths, any error,
timings, artifact paths, warnings, and dashboard URL. `--port 0` asks the
operating system for a free local port. `--docx-only` retains status and DOCX
downloads while intentionally omitting PDF preview.

RusDox watches the input, active config, local includes, and asset paths. The
terminal and dashboard identify the trigger as input, config, or asset/include.

The original `rusdox watch` command remains available for compatibility and now
shares the debounced dependency watcher:

```bash
rusdox watch mydoc.yaml --docx-only --poll-interval-ms 250
rusdox watch examples --with-pdf
```

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

Isolate one pipeline for reproducible measurement:

```bash
rusdox bench mydoc.yaml --pipeline validation --iterations 5 --warmup 1
rusdox bench mydoc.yaml --pipeline docx --iterations 5 --warmup 1
rusdox bench mydoc.yaml --pipeline pdf --iterations 5 --warmup 1
rusdox bench mydoc.yaml --pipeline dual --iterations 5 --warmup 1
rusdox bench existing.docx --pipeline existing-docx --iterations 5 --warmup 1
```

Keep the generated artifacts instead of using a temporary output workspace:

```bash
rusdox bench mydoc.yaml --keep-output
```

Bench JSON reports include the pipeline, input SHA-256 and byte size, parse, validation, compose, DOCX write, PDF render, existing-DOCX open/save, total runtime, output byte sizes, and average/minimum/median/maximum summaries. The complete host and peak-memory protocol is documented in [Reproducible Performance](performance.md).

## Word-native templates

Inspect placeholders and structural errors:

    rusdox template inspect proposal.docx
    rusdox template inspect proposal.docx --format json

Render from JSON and generate the edited DOCX, native PDF, page snapshots, and parity report:

    rusdox template render proposal.docx data.json --strict
    rusdox template verify proposal.docx data.json --strict --format json

Use --name for the artifact stem, --output-root for the complete evidence root, and --config for PDF renderer configuration. See [Word-native Templates](word-templates.md) for syntax v1, strict behavior, loops, conditions, partials, preservation guarantees, and limitations.

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
