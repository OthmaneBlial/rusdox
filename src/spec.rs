use std::fs;
use std::path::{Path, PathBuf};

use serde::{Deserialize, Serialize};

use crate::{DocxError, HeaderFooter, PageNumbering, PageSetup, Result};

/// A high-level, serializable document specification.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(default)]
pub struct DocumentSpec {
    /// Optional logical output name. Falls back to the spec file stem when absent.
    pub output_name: Option<String>,
    /// Optional page size and margin overrides for the document section.
    pub page_setup: Option<PageSetup>,
    /// Optional default header template.
    pub header: Option<HeaderFooter>,
    /// Optional default footer template.
    pub footer: Option<HeaderFooter>,
    /// Optional page numbering format and restart control.
    pub page_numbering: Option<PageNumbering>,
    pub blocks: Vec<BlockSpec>,
    #[serde(skip)]
    asset_base_dir: Option<PathBuf>,
}

impl DocumentSpec {
    pub fn new() -> Self {
        Self::default()
    }

    /// Loads a document specification from a file path.
    ///
    /// `.yaml`, `.yml`, `.json`, and `.toml` are supported.
    pub fn load_from_path(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref();
        let content = fs::read_to_string(path)?;
        let extension = path
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();

        let mut spec = match extension.as_str() {
            "yaml" | "yml" | "" => Self::from_yaml_str(&content),
            "json" => Self::from_json_str(&content),
            "toml" => Self::from_toml_str(&content),
            other => Err(DocxError::parse(format!(
                "unsupported document spec extension '{other}', expected .yaml, .yml, .json, or .toml"
            ))),
        }?;
        spec.asset_base_dir = path.parent().map(Path::to_path_buf);
        Ok(spec)
    }

    /// Parses a YAML document specification string.
    pub fn from_yaml_str(content: &str) -> Result<Self> {
        serde_yaml::from_str(content)
            .map_err(|error| DocxError::parse(format!("invalid YAML document spec: {error}")))
    }

    /// Parses a JSON document specification string.
    pub fn from_json_str(content: &str) -> Result<Self> {
        serde_json::from_str(content)
            .map_err(|error| DocxError::parse(format!("invalid JSON document spec: {error}")))
    }

    /// Parses a TOML document specification string.
    pub fn from_toml_str(content: &str) -> Result<Self> {
        toml::from_str(content)
            .map_err(|error| DocxError::parse(format!("invalid TOML document spec: {error}")))
    }

    /// Serializes the document specification as YAML.
    pub fn to_yaml_string(&self) -> Result<String> {
        serde_yaml::to_string(self)
            .map_err(|error| DocxError::parse(format!("failed to serialize YAML spec: {error}")))
    }

    /// Serializes the document specification as JSON.
    pub fn to_json_pretty(&self) -> Result<String> {
        serde_json::to_string_pretty(self)
            .map_err(|error| DocxError::parse(format!("failed to serialize JSON spec: {error}")))
    }

    /// Serializes the document specification as TOML.
    pub fn to_toml_pretty(&self) -> Result<String> {
        toml::to_string_pretty(self)
            .map_err(|error| DocxError::parse(format!("failed to serialize TOML spec: {error}")))
    }

    /// Saves the current specification to disk.
    ///
    /// `.yaml`, `.yml`, `.json`, and `.toml` are supported.
    /// If no extension is provided, YAML is used by default.
    pub fn save_to_path(&self, path: impl AsRef<Path>) -> Result<()> {
        let path = path.as_ref();
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent)?;
        }

        let extension = path
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or("yaml")
            .to_ascii_lowercase();

        let content = match extension.as_str() {
            "yaml" | "yml" | "" => self.to_yaml_string()?,
            "json" => self.to_json_pretty()?,
            "toml" => self.to_toml_pretty()?,
            other => {
                return Err(DocxError::parse(format!(
                    "unsupported document spec extension '{other}', expected .yaml, .yml, .json, or .toml"
                )))
            }
        };

        fs::write(path, content)?;
        Ok(())
    }

    /// Writes a commented YAML starter document template to disk.
    pub fn write_yaml_template(path: impl AsRef<Path>) -> Result<()> {
        let path = path.as_ref();
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent)?;
        }
        fs::write(path, Self::default_yaml_template())?;
        Ok(())
    }

    /// Returns the default YAML starter document template.
    pub fn default_yaml_template() -> &'static str {
        DEFAULT_YAML_TEMPLATE
    }

    /// Returns the base directory used to resolve relative asset paths.
    pub fn asset_base_dir(&self) -> Option<&Path> {
        self.asset_base_dir.as_deref()
    }

    /// Sets the base directory used to resolve relative asset paths.
    pub fn set_asset_base_dir(&mut self, base_dir: Option<PathBuf>) -> &mut Self {
        self.asset_base_dir = base_dir;
        self
    }

    /// Sets the base directory used to resolve relative asset paths in builder style.
    pub fn with_asset_base_dir(mut self, base_dir: impl Into<PathBuf>) -> Self {
        self.asset_base_dir = Some(base_dir.into());
        self
    }
}

/// A high-level document block.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum BlockSpec {
    CoverTitle {
        text: String,
    },
    Title {
        text: String,
    },
    Subtitle {
        text: String,
    },
    Hero {
        text: String,
    },
    CenteredNote {
        text: String,
    },
    PageHeading {
        text: String,
    },
    Section {
        text: String,
    },
    Body {
        text: String,
    },
    Tagline {
        text: String,
    },
    Paragraph {
        spec: ParagraphSpec,
    },
    Bullets {
        items: Vec<String>,
    },
    Numbered {
        items: Vec<String>,
    },
    LabelValues {
        items: Vec<LabelValueSpec>,
    },
    Metrics {
        items: Vec<MetricSpec>,
    },
    Table {
        spec: TableSpec,
    },
    Image {
        #[serde(flatten)]
        spec: VisualSpec,
    },
    Logo {
        #[serde(flatten)]
        spec: VisualSpec,
    },
    Signature {
        #[serde(flatten)]
        spec: VisualSpec,
    },
    Chart {
        #[serde(flatten)]
        spec: VisualSpec,
    },
    Spacer,
}

/// A fully specified paragraph block.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(default)]
pub struct ParagraphSpec {
    pub runs: Vec<RunSpec>,
    pub alignment: Option<ParagraphAlignmentSpec>,
    pub spacing_before_twips: Option<u32>,
    pub spacing_after_twips: Option<u32>,
    pub page_break_before: bool,
}

impl ParagraphSpec {
    pub fn new<I>(runs: I) -> Self
    where
        I: IntoIterator<Item = RunSpec>,
    {
        Self {
            runs: runs.into_iter().collect(),
            ..Self::default()
        }
    }
}

/// A serializable paragraph alignment.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ParagraphAlignmentSpec {
    Left,
    Center,
    Right,
    Justified,
}

/// A fully specified text run.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(default)]
pub struct RunSpec {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: Option<UnderlineStyleSpec>,
    pub strikethrough: bool,
    pub small_caps: bool,
    pub shadow: bool,
    pub color: Option<String>,
    pub font_family: Option<String>,
    pub size_pt: Option<f32>,
    pub vertical_align: Option<VerticalAlignSpec>,
}

/// A fully specified visual/image block.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(default)]
pub struct VisualSpec {
    pub path: String,
    pub alt_text: Option<String>,
    pub alignment: Option<ParagraphAlignmentSpec>,
    pub width_twips: Option<u32>,
    pub height_twips: Option<u32>,
    pub max_width_twips: Option<u32>,
    pub max_height_twips: Option<u32>,
}

impl VisualSpec {
    pub fn new(path: impl Into<String>) -> Self {
        Self {
            path: path.into(),
            ..Self::default()
        }
    }
}

impl RunSpec {
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            ..Self::default()
        }
    }
}

/// A serializable underline style.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum UnderlineStyleSpec {
    Single,
    Double,
    Dotted,
    Dash,
    Wavy,
    Words,
    None,
}

/// A serializable run vertical alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum VerticalAlignSpec {
    Superscript,
    Subscript,
    Baseline,
}

/// A simple label-value pair block item.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct LabelValueSpec {
    pub label: String,
    pub value: String,
}

/// A metric card item.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct MetricSpec {
    pub label: String,
    pub value: String,
    pub tone: Tone,
}

/// Shared semantic color tone.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum Tone {
    Positive,
    Neutral,
    Warning,
    Risk,
}

/// A grid table specification.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct TableSpec {
    pub columns: Vec<ColumnSpec>,
    pub rows: Vec<RowSpec>,
}

/// A table column definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ColumnSpec {
    pub label: String,
    pub width: u32,
}

/// A table row definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct RowSpec {
    pub cells: Vec<CellSpec>,
}

/// A table cell definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum CellSpec {
    Text { text: String },
    Status(StatusSpec),
}

/// A status cell definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct StatusSpec {
    pub text: String,
    pub tone: Tone,
}

pub fn document<I>(blocks: I) -> DocumentSpec
where
    I: IntoIterator<Item = BlockSpec>,
{
    DocumentSpec {
        output_name: None,
        page_setup: None,
        header: None,
        footer: None,
        page_numbering: None,
        blocks: blocks.into_iter().collect(),
        asset_base_dir: None,
    }
}

pub fn cover_title(text: impl Into<String>) -> BlockSpec {
    BlockSpec::CoverTitle { text: text.into() }
}

pub fn title(text: impl Into<String>) -> BlockSpec {
    BlockSpec::Title { text: text.into() }
}

pub fn subtitle(text: impl Into<String>) -> BlockSpec {
    BlockSpec::Subtitle { text: text.into() }
}

pub fn hero(text: impl Into<String>) -> BlockSpec {
    BlockSpec::Hero { text: text.into() }
}

pub fn centered_note(text: impl Into<String>) -> BlockSpec {
    BlockSpec::CenteredNote { text: text.into() }
}

pub fn page_heading(text: impl Into<String>) -> BlockSpec {
    BlockSpec::PageHeading { text: text.into() }
}

pub fn section(text: impl Into<String>) -> BlockSpec {
    BlockSpec::Section { text: text.into() }
}

pub fn body(text: impl Into<String>) -> BlockSpec {
    BlockSpec::Body { text: text.into() }
}

pub fn tagline(text: impl Into<String>) -> BlockSpec {
    BlockSpec::Tagline { text: text.into() }
}

pub fn paragraph(spec: ParagraphSpec) -> BlockSpec {
    BlockSpec::Paragraph { spec }
}

pub fn bullets<I, S>(items: I) -> BlockSpec
where
    I: IntoIterator<Item = S>,
    S: Into<String>,
{
    BlockSpec::Bullets {
        items: items.into_iter().map(Into::into).collect(),
    }
}

pub fn numbered<I, S>(items: I) -> BlockSpec
where
    I: IntoIterator<Item = S>,
    S: Into<String>,
{
    BlockSpec::Numbered {
        items: items.into_iter().map(Into::into).collect(),
    }
}

pub fn label_values<I, L, V>(items: I) -> BlockSpec
where
    I: IntoIterator<Item = (L, V)>,
    L: Into<String>,
    V: Into<String>,
{
    BlockSpec::LabelValues {
        items: items
            .into_iter()
            .map(|(label, value)| LabelValueSpec {
                label: label.into(),
                value: value.into(),
            })
            .collect(),
    }
}

pub fn metric(label: impl Into<String>, value: impl Into<String>, tone: Tone) -> MetricSpec {
    MetricSpec {
        label: label.into(),
        value: value.into(),
        tone,
    }
}

pub fn metrics<I>(items: I) -> BlockSpec
where
    I: IntoIterator<Item = MetricSpec>,
{
    BlockSpec::Metrics {
        items: items.into_iter().collect(),
    }
}

pub fn col(label: impl Into<String>, width: u32) -> ColumnSpec {
    ColumnSpec {
        label: label.into(),
        width,
    }
}

pub fn text(text: impl Into<String>) -> CellSpec {
    CellSpec::Text { text: text.into() }
}

pub fn status(text: impl Into<String>, tone: Tone) -> StatusSpec {
    StatusSpec {
        text: text.into(),
        tone,
    }
}

pub fn row<T>(value: T) -> RowSpec
where
    T: IntoRowSpec,
{
    value.into_row_spec()
}

pub fn table<C, R>(columns: C, rows: R) -> BlockSpec
where
    C: IntoIterator<Item = ColumnSpec>,
    R: IntoIterator<Item = RowSpec>,
{
    BlockSpec::Table {
        spec: TableSpec {
            columns: columns.into_iter().collect(),
            rows: rows.into_iter().collect(),
        },
    }
}

pub fn image(path: impl Into<String>) -> BlockSpec {
    BlockSpec::Image {
        spec: VisualSpec::new(path),
    }
}

pub fn logo(path: impl Into<String>) -> BlockSpec {
    BlockSpec::Logo {
        spec: VisualSpec::new(path),
    }
}

pub fn signature(path: impl Into<String>) -> BlockSpec {
    BlockSpec::Signature {
        spec: VisualSpec::new(path),
    }
}

pub fn chart(path: impl Into<String>) -> BlockSpec {
    BlockSpec::Chart {
        spec: VisualSpec::new(path),
    }
}

pub fn spacer() -> BlockSpec {
    BlockSpec::Spacer
}

pub trait IntoRowSpec {
    fn into_row_spec(self) -> RowSpec;
}

impl From<&str> for CellSpec {
    fn from(value: &str) -> Self {
        text(value)
    }
}

impl From<String> for CellSpec {
    fn from(value: String) -> Self {
        text(value)
    }
}

impl From<StatusSpec> for CellSpec {
    fn from(value: StatusSpec) -> Self {
        CellSpec::Status(value)
    }
}

macro_rules! impl_into_row_spec {
    ($( $name:ident ),+ $(,)?) => {
        impl<$( $name ),+> IntoRowSpec for ($( $name, )+)
        where
            $( $name: Into<CellSpec>, )+
        {
            #[allow(non_snake_case)]
            fn into_row_spec(self) -> RowSpec {
                let ($( $name, )+) = self;
                RowSpec {
                    cells: vec![$( $name.into(), )+],
                }
            }
        }
    };
}

impl_into_row_spec!(A);
impl_into_row_spec!(A, B);
impl_into_row_spec!(A, B, C);
impl_into_row_spec!(A, B, C, D);
impl_into_row_spec!(A, B, C, D, E);

const DEFAULT_YAML_TEMPLATE: &str = r#"# RusDox document spec template
# Save this file as `mydoc.yaml` and run:
#   rusdox mydoc.yaml

output_name: my-document
# Optional layout controls:
# page_setup:
#   width_twips: 12240
#   height_twips: 15840
#   margin_top_twips: 1440
#   margin_right_twips: 1440
#   margin_bottom_twips: 1440
#   margin_left_twips: 1440
# header:
#   text: "Quarterly review"
#   alignment: center
# footer:
#   text: "Page {page} of {pages}"
#   alignment: right
# page_numbering:
#   start_at: 1
#   format: decimal
blocks:
  - type: title
    text: My Document
  - type: subtitle
    text: Written as data, rendered by Rust
  - type: section
    text: Summary
  - type: body
    text: Replace this with your real content.
  - type: bullets
    items:
      - Keep content in order.
      - Let config handle styling.
      - Render to DOCX and PDF with one command.
"#;

#[cfg(test)]
mod tests {
    use std::fs;

    use tempfile::tempdir;

    use super::{
        body, bullets, chart, col, document, image, label_values, logo, metric, metrics, numbered,
        paragraph, row, section, signature, status, table, title, BlockSpec,
        ParagraphAlignmentSpec, ParagraphSpec, RunSpec, Tone, UnderlineStyleSpec, VisualSpec,
    };
    use crate::{HeaderFooter, PageNumberFormat, PageNumbering, PageSetup, ParagraphAlignment};

    #[test]
    fn spec_round_trips_through_json() {
        let spec = document([
            title("Board Report"),
            section("Summary"),
            body("Everything is readable."),
            bullets(["Fast", "Configurable"]),
            numbered(["One", "Two"]),
            label_values([("Owner", "Finance")]),
            metrics([metric("ARR", "$18.7M", Tone::Positive)]),
            table(
                [col("Item", 4_000), col("Status", 2_000)],
                [row(("Pipeline", status("Watch", Tone::Warning)))],
            ),
            BlockSpec::Image {
                spec: VisualSpec {
                    path: "assets/template-gallery.png".to_string(),
                    alt_text: Some("Gallery".to_string()),
                    max_width_twips: Some(7_200),
                    ..VisualSpec::default()
                },
            },
            logo("assets/rusdox-mark.svg"),
            chart("assets/benchmark-stress-1000-pages.svg"),
            signature("assets/signature-demo.svg"),
        ]);

        let json = serde_json::to_string_pretty(&spec).expect("serialize spec");
        let round_trip: super::DocumentSpec =
            serde_json::from_str(&json).expect("deserialize spec");

        assert_eq!(round_trip, spec);
    }

    #[test]
    fn spec_round_trips_through_yaml() {
        let spec = super::DocumentSpec {
            output_name: Some("hello-world".to_string()),
            page_setup: Some(PageSetup::new(11_880, 16_380).margins(900, 1_000, 1_100, 1_200)),
            header: Some(
                HeaderFooter::new("Board Report").with_alignment(ParagraphAlignment::Center),
            ),
            footer: Some(
                HeaderFooter::new("Page {page} of {pages}")
                    .with_alignment(ParagraphAlignment::Right),
            ),
            page_numbering: Some(PageNumbering::new(PageNumberFormat::UpperRoman).start_at(3)),
            blocks: vec![
                title("Hello"),
                paragraph(ParagraphSpec {
                    runs: vec![
                        RunSpec {
                            text: "Bold".to_string(),
                            bold: true,
                            ..RunSpec::default()
                        },
                        RunSpec {
                            text: " | ".to_string(),
                            ..RunSpec::default()
                        },
                        RunSpec {
                            text: "Underline".to_string(),
                            underline: Some(UnderlineStyleSpec::Single),
                            ..RunSpec::default()
                        },
                    ],
                    alignment: Some(ParagraphAlignmentSpec::Center),
                    ..ParagraphSpec::default()
                }),
                image("assets/template-gallery.png"),
            ],
            asset_base_dir: None,
        };

        let yaml = spec.to_yaml_string().expect("serialize yaml");
        let round_trip = super::DocumentSpec::from_yaml_str(&yaml).expect("deserialize yaml");

        assert_eq!(round_trip, spec);
    }

    #[test]
    fn load_from_path_uses_extension_based_parser() {
        let temp = tempdir().expect("temp dir");
        let yaml_path = temp.path().join("spec.yaml");
        let json_path = temp.path().join("spec.json");

        fs::write(
            &yaml_path,
            r#"
output_name: hello-world
blocks:
  - type: title
    text: Hello
"#,
        )
        .expect("write yaml");
        fs::write(
            &json_path,
            r#"{"blocks":[{"type":"title","text":"Hello"}]}"#,
        )
        .expect("write json");

        let yaml_spec = super::DocumentSpec::load_from_path(&yaml_path).expect("load yaml");
        let json_spec = super::DocumentSpec::load_from_path(&json_path).expect("load json");

        assert_eq!(yaml_spec.output_name.as_deref(), Some("hello-world"));
        assert_eq!(yaml_spec.blocks.len(), 1);
        assert_eq!(json_spec.blocks.len(), 1);
    }

    #[test]
    fn tuple_rows_accept_plain_text_and_status_cells() {
        let row = row(("ARR", "$18.7M", status("Strong", Tone::Positive), "On plan"));
        assert_eq!(row.cells.len(), 4);
    }
}
