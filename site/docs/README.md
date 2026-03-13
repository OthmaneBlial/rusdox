# RusDox Docs

RusDox lets you write documents as YAML and generate real `.docx` and `.pdf` files with a Rust engine underneath.

If you are new, read these pages in order:

1. [Getting Started](getting-started.md)
2. [YAML Guide](yaml-guide.md)
3. [Configuration](configuration.md)
4. [CLI Guide](cli.md)
5. [Gallery](gallery.md)
6. [Rust API](rust-api.md)
7. [GitHub Setup](github-setup.md)

## Pick Your Path

If you want the simplest path:

- Start with [Getting Started](getting-started.md)
- Use YAML files
- Run `rusdox validate mydoc.yaml` before big renders or CI
- Use `rusdox config wizard --level basic`

If you want more control:

- Read [YAML Guide](yaml-guide.md)
- Read [Configuration](configuration.md)
- Use local `./rusdox.toml` overrides per project

If you are a developer and want full power:

- Read [Rust API](rust-api.md)
- Use `DocumentSpec` from Rust for programmatic YAML-style documents
- Drop down to `Document`, `Paragraph`, `Run`, and `Table` when needed

## What The Docs Cover

- how to install RusDox
- how to create your first YAML document
- how output folders work
- how DOCX and PDF generation works
- how to style documents with config and named styles
- how to render one file or a whole folder
- how to validate, watch, and benchmark from the CLI
- how to use the Rust layer for advanced cases
- how to present the repository well on GitHub

## Core Idea

RusDox works best when you keep these roles separate:

- content lives in YAML
- styling lives in config
- speed lives in Rust
