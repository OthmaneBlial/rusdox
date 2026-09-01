# Getting Started

This is the fastest way to go from nothing to a generated `.docx` and `.pdf`.

Already installed? Create a complete, self-contained result first:

```bash
rusdox demo
```

This creates a new `rusdox-demo/` directory containing an editable YAML spec,
a DOCX, a native PDF, deterministic page snapshots, and HTML/JSON parity
reports. It refuses to use an existing destination, so it cannot overwrite a
previous demo or your own files. Use `rusdox demo my-first-report` to choose a
different directory.

## 1. Install RusDox

macOS and Linux:

```bash
curl -fsSL https://raw.githubusercontent.com/OthmaneBlial/rusdox/2b3ca4eda4ab8389dc0e54198811bbaa3c368e44/scripts/install.sh | sh
```

Windows PowerShell:

```powershell
irm https://raw.githubusercontent.com/OthmaneBlial/rusdox/2b3ca4eda4ab8389dc0e54198811bbaa3c368e44/scripts/install.ps1 -OutFile install-rusdox.ps1
.\install-rusdox.ps1
```

The installer adds only the `rusdox` binary. Built-in defaults work immediately;
run `rusdox config init --template` only when you want a user config.

It does not clone the full GitHub repository.

## 2. Create a document file

```bash
mkdir my-rusdox-docs
cd my-rusdox-docs
rusdox init-doc mydoc.yaml
```

That creates a starter YAML document.

## 3. Edit the YAML

Example:

```yaml
version: 1
output_name: client-brief
blocks:
  - type: title
    text: Client Brief
  - type: subtitle
    text: Q2 rollout
  - type: section
    text: Summary
  - type: body
    text: Launch is approved pending final security FAQ wording.
  - type: bullets
    items:
      - Pricing is approved.
      - Support macros are in review.
      - Commercial release is planned for April 7.
```

Think of `blocks:` as the document itself, top to bottom.

## 4. Render the files

```bash
rusdox mydoc.yaml
```

By default you get:

- `generated/client-brief.docx`
- `rendered/client-brief.pdf`

If your YAML does not set `output_name`, RusDox uses the file name.

## 5. Change the style

The easiest way is the config wizard:

```bash
rusdox config wizard --level basic
```

For more control:

```bash
rusdox config wizard --level advanced
```

If you want settings only for the current project:

```bash
rusdox config wizard --path ./rusdox.toml --level basic
```

## 6. Render a whole folder

If a folder contains YAML specs, RusDox can render all of them:

```bash
rusdox examples
```

## 7. Validate before you render in CI

```bash
rusdox validate mydoc.yaml
```

That catches semantic issues such as invalid colors, table shape mismatches, missing assets, or unknown named styles before output is written.

## 8. Watch while you edit

```bash
rusdox watch mydoc.yaml
```

RusDox rebuilds when the spec or active config changes.

## 9. Measure the pipeline

```bash
rusdox bench mydoc.yaml --iterations 5 --warmup 1
```

This reports parse, validation, compose, DOCX, PDF, and total timings from the CLI.

## What Happens Behind The Scenes

When you run `rusdox mydoc.yaml`, RusDox:

1. Reads the YAML document spec
2. Applies your config
3. Builds a DOCX document in pure Rust
4. Writes a PDF preview in pure Rust

No Microsoft Word is required.

No LibreOffice is required.

## Next Pages

- [YAML Guide](yaml-guide.md)
- [Configuration](configuration.md)
- [CLI Guide](cli.md)
- [Gallery](gallery.md)
- [Zero-install playground](https://othmaneblial.github.io/rusdox/playground/)
- [Browser rendering feasibility and privacy boundary](wasm-feasibility.md)
