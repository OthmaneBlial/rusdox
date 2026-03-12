# Getting Started

This is the fastest way to go from nothing to a generated `.docx` and `.pdf`.

## 1. Install RusDox

macOS and Linux:

```bash
curl -fsSL https://raw.githubusercontent.com/OthmaneBlial/rusdox/main/scripts/install.sh | sh
```

Windows PowerShell:

```powershell
irm https://raw.githubusercontent.com/OthmaneBlial/rusdox/main/scripts/install.ps1 | iex
```

The installer adds the `rusdox` binary and creates `~/rusdox/config.toml` if it does not exist yet.

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
