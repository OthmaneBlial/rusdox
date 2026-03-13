# Rust API

RusDox is YAML-first for everyday authoring, but the Rust API stays available for advanced and programmable workflows.

Use Rust when you need:

- document content generated from live data
- loops, conditions, or reusable functions
- integration inside a Rust service or CLI
- direct control over document metadata and custom properties
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

`DocumentSpec` also exposes `metadata` and `styles`, so document properties and reusable named styles can be defined once and reused consistently.

```rust
use rusdox::spec::{body, section, title, DocumentSpec};
use rusdox::{DocumentMetadata};

let mut spec = DocumentSpec::new();
spec.metadata = DocumentMetadata::new()
    .title("Weekly Brief")
    .author("RusDox Studio")
    .subject("Executive update")
    .keyword("weekly")
    .custom_property("Audience", "Leadership");
spec.blocks = vec![
    title("Weekly Brief"),
    section("Summary"),
    body("Pipeline grew 14% week over week."),
];
```

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

## Reusable Named Styles

Named styles are available in both the spec layer and the low-level document model.

Built-in fallback ids:

- paragraph: `Normal`
- run: `DefaultParagraphFont`
- table: `TableNormal`

```rust
use rusdox::{
    Border, BorderStyle, Document, Paragraph, ParagraphAlignment, ParagraphStyle,
    ParagraphStyleProperties, Run, RunStyle, RunStyleProperties, Stylesheet, Table, TableBorders,
    TableCell, TableRow, TableStyle, TableStyleProperties,
};

fn main() -> rusdox::Result<()> {
    let border = Border::new(BorderStyle::Single).size(8).color("CBD5E1");
    let styles = Stylesheet::new()
        .add_paragraph_style(
            ParagraphStyle::new("lead")
                .based_on("Normal")
                .paragraph(
                    ParagraphStyleProperties::new()
                        .alignment(ParagraphAlignment::Center)
                        .spacing_after(180),
                )
                .run(RunStyleProperties::new().bold().color("0F172A")),
        )
        .add_run_style(
            RunStyle::new("accent")
                .based_on("DefaultParagraphFont")
                .properties(RunStyleProperties::new().italic().color("AA5500")),
        )
        .add_table_style(
            TableStyle::new("grid")
                .based_on("TableNormal")
                .properties(
                    TableStyleProperties::new()
                        .width(9_360)
                        .borders(TableBorders::new().top(border.clone()).bottom(border)),
                ),
        );

    let mut document = Document::new().with_styles(styles);
    document.push_paragraph(
        Paragraph::new()
            .with_style("lead")
            .add_run(Run::from_text("Quarterly ").with_style("accent"))
            .add_run(Run::from_text("review")),
    );
    document.push_table(
        Table::new().style("grid").add_row(
            TableRow::new().add_cell(
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("ARR"))),
            ),
        ),
    );

    document.save("styled-output.docx")?;
    Ok(())
}
```

Use these APIs when:

- multiple paragraphs should share the same typography and spacing rules
- run-level emphasis should stay stable across documents
- table framing should be reusable instead of copied as direct borders and widths

## First-Class Metadata

Use `DocumentMetadata` when the generated DOCX should carry clean package properties.

```rust
use rusdox::{Document, DocumentMetadata};

let metadata = DocumentMetadata::new()
    .title("Board Report")
    .author("Finance")
    .subject("Q4 review")
    .keyword("board")
    .custom_property("Client", "Northwind Health");

let document = Document::new().with_metadata(metadata);
```

Metadata works through both `DocumentSpec` and `Document`, and RusDox writes it into `docProps/core.xml` plus `docProps/custom.xml`.

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
- you want reusable styles through `Document::with_styles(...)`
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
