# Rust API

RusDox is YAML-first for everyday authoring, but the Rust API stays available for advanced and programmable workflows.

Use Rust when you need:

- document content generated from live data
- loops, conditions, or reusable functions
- integration inside a Rust service or CLI
- lower-level formatting beyond the YAML surface

## Choose The Right Layer

For most users:

- write YAML
- style with config
- run `rusdox mydoc.yaml`

For advanced users:

- use `spec::DocumentSpec` from Rust when you still want a data-shaped document model
- use `studio::Studio` helpers when you want config-driven paragraphs and tables
- use `Document`, `Paragraph`, `Run`, `Table`, and `Visual` directly when you need full control

## Install

```bash
cargo add rusdox
```

## Config

Most users should set config through the CLI wizard:

```bash
rusdox config path
rusdox config wizard --level basic
rusdox config wizard --level advanced
```

The installer creates `~/rusdox/config.toml` if it does not exist yet.

For per-project overrides:

```bash
rusdox config wizard --path ./rusdox.toml --level basic
```

Config load order is:

1. `./rusdox.toml`
2. `~/rusdox/config.toml`
3. built-in defaults

`Studio::from_default_file_or_default()` follows that same order.

## High-Level Rust: Compose From A Spec

If you like the YAML model but want to generate it programmatically, use `DocumentSpec`.

```rust
use rusdox::spec::{body, bullets, section, title, DocumentSpec};
use rusdox::studio::Studio;

fn main() -> rusdox::Result<()> {
    let studio = Studio::from_default_file_or_default()?;

    let mut spec = DocumentSpec::new();
    spec.output_name = Some("weekly-brief".to_string());
    spec.blocks = vec![
        title("Weekly Brief"),
        section("Summary"),
        body("Pipeline grew 14% week over week."),
        bullets([
            "Security review closed",
            "Support handoff approved",
            "Launch remains on schedule",
        ]),
    ];

    studio.save_spec_named(&spec, "weekly-brief")?;
    Ok(())
}
```

This is the best Rust path when:

- the document is mostly standard sections
- content comes from code, not a static YAML file
- you still want the document to stay easy to reason about

## Hybrid Rust: Start With A Spec, Then Add Custom Pieces

You can also compose a spec and then append lower-level content.

```rust
use rusdox::spec::{body, section, title, DocumentSpec};
use rusdox::studio::Studio;
use rusdox::{Paragraph, Run};

fn main() -> rusdox::Result<()> {
    let studio = Studio::from_default_file_or_default()?;

    let mut spec = DocumentSpec::new();
    spec.blocks = vec![
        title("Launch Packet"),
        section("Summary"),
        body("Core rollout is approved."),
    ];

    let mut document = studio.compose(&spec);
    document.push_paragraph(
        Paragraph::new()
            .add_run(studio.text_run("Custom note: ").bold())
            .add_run(studio.text_run("regional approvals still pending.")),
    );

    studio.save_named(&document, "launch-packet")?;
    Ok(())
}
```

This is a good middle ground when 90% of the document fits the high-level API and only a few sections need special handling.

## Config-Driven Builders With `Studio`

`Studio` is the main advanced entry point.

It gives you:

- config-aware text runs
- config-aware headings and body paragraphs
- config-aware table helpers
- document saving with DOCX and optional PDF output

Common helpers include:

- `studio.title(...)`
- `studio.subtitle(...)`
- `studio.section(...)`
- `studio.body(...)`
- `studio.cover_title(...)`
- `studio.page_heading(...)`
- `studio.tagline(...)`
- `studio.label_value(...)`
- `studio.metric_cell(...)`
- `studio.header_cell(...)`
- `studio.data_cell(...)`
- `studio.status_cell(...)`
- `studio.grid_borders()`
- `studio.card_borders()`

There are also convenience free functions in `rusdox::studio` such as `title(...)`, `body(...)`, and `save_with_pdf(...)` that use the configured default `Studio`.

## Low-Level Rust: Build The Document Yourself

When you need full control, use the core document model directly.

```rust
use rusdox::{
    Border, BorderStyle, Document, Paragraph, Run, Table, TableBorders, TableCell, TableRow,
    UnderlineStyle, Visual,
};

fn main() -> rusdox::Result<()> {
    let accent = TableBorders::new()
        .top(Border::new(BorderStyle::Single).size(8).color("1F2937"))
        .bottom(Border::new(BorderStyle::Single).size(8).color("1F2937"));

    let mut doc = Document::new();
    doc.push_paragraph(
        Paragraph::new()
            .add_run(Run::from_text("This is ").bold())
            .add_run(Run::from_text("blazing fast").italic().color("DC2626"))
            .add_run(Run::from_text(" and ").underline(UnderlineStyle::Single))
            .add_run(Run::from_text("typed.").small_caps()),
    );

    doc.push_table(
        Table::new()
            .width(9_360)
            .borders(accent)
            .add_row(
                TableRow::new()
                    .add_cell(TableCell::new().width(4_680).add_paragraph(
                        Paragraph::new().add_run(Run::from_text("Header A").bold()),
                    ))
                    .add_cell(TableCell::new().width(4_680).add_paragraph(
                        Paragraph::new().add_run(Run::from_text("Header B").bold()),
                    )),
            ),
    );

    doc.push_visual(
        Visual::logo("assets/rusdox-mark.svg")
            .alt_text_text("RusDox logo")
            .max_width_twips(2_200),
    );

    doc.save("output.docx")?;
    Ok(())
}
```

Use this layer when:

- you need exact run-level formatting
- you want to open and modify existing DOCX files
- you are building custom abstractions on top of RusDox

## Open Existing DOCX Files

RusDox can also read and preserve existing packages:

- `Document::open(...)`
- `Document::open_read_only(...)`

This is useful when you want to:

- inspect document text
- modify a document in place
- preserve package parts you are not touching

## Use YAML Specs From Rust

You can load and save specs in Rust too:

```rust
use rusdox::spec::DocumentSpec;

fn main() -> rusdox::Result<()> {
    let spec = DocumentSpec::load_from_path("examples/board_report.yaml")?;
    let yaml = spec.to_yaml_string()?;
    let json = spec.to_json_pretty()?;
    let toml = spec.to_toml_pretty()?;

    assert!(!yaml.is_empty());
    assert!(!json.is_empty());
    assert!(!toml.is_empty());
    Ok(())
}
```

That makes it easy to:

- generate YAML specs from application data
- validate specs before rendering
- convert between YAML, JSON, and TOML

## Script Mode

If you want a programmable entrypoint without creating a full Rust crate, RusDox still supports `.rs` scripts:

```bash
rusdox init-script mydoc.rs
rusdox mydoc.rs
```

Your script must expose:

```rust
pub fn build_document(studio: &rusdox::studio::Studio) -> rusdox::Result<rusdox::Document>
```

This is good for quick internal tools and local automation.

## Practical Recommendation

The best progression is:

1. Start with YAML
2. Move to `DocumentSpec` in Rust if content becomes dynamic
3. Drop to `Document` and `Run` only where the higher-level layers stop being enough

That keeps the product simple for most documents while preserving a real escape hatch for power users.
