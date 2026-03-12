# CLI Guide

RusDox has one main job: render document specs into `.docx` and `.pdf`.

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
