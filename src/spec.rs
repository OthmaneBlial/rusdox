use std::fs;
use std::io::Read;
use std::path::{Path, PathBuf};

use schemars::JsonSchema;
use serde::{Deserialize, Serialize};

use crate::spec_expand::{
    expand_document_spec_value_with_limits, expand_yaml_document_spec_with_limits,
};
use crate::DocumentMetadata;
use crate::{DocxError, HeaderFooter, InputLimits, PageNumbering, PageSetup, Result, Stylesheet};

/// Current supported document-spec contract.
pub const SPEC_VERSION: u32 = 1;

/// Returns the current document-spec version for serde defaults.
pub const fn default_spec_version() -> u32 {
    SPEC_VERSION
}

/// A high-level, serializable document specification.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, JsonSchema)]
#[serde(default)]
pub struct DocumentSpec {
    /// Version of the declarative document-spec contract.
    pub version: u32,
    /// Optional logical output name. Falls back to the spec file stem when absent.
    pub output_name: Option<String>,
    /// Optional core and custom document metadata written into the DOCX package.
    pub metadata: DocumentMetadata,
    /// Optional page size and margin overrides for the document section.
    pub page_setup: Option<PageSetup>,
    /// Optional default header template.
    pub header: Option<HeaderFooter>,
    /// Optional default footer template.
    pub footer: Option<HeaderFooter>,
    /// Optional page numbering format and restart control.
    pub page_numbering: Option<PageNumbering>,
    /// Reusable named styles available to blocks and runs in this document.
    pub styles: Stylesheet,
    /// Ordered blocks.
    pub blocks: Vec<BlockSpec>,
    #[serde(skip)]
    #[schemars(skip)]
    asset_base_dir: Option<PathBuf>,
}

impl Default for DocumentSpec {
    fn default() -> Self {
        Self {
            version: SPEC_VERSION,
            output_name: None,
            metadata: DocumentMetadata::default(),
            page_setup: None,
            header: None,
            footer: None,
            page_numbering: None,
            styles: Stylesheet::default(),
            blocks: Vec::new(),
            asset_base_dir: None,
        }
    }
}

impl DocumentSpec {
    /// Creates a value with default settings.
    pub fn new() -> Self {
        Self::default()
    }

    /// Loads a document specification from a file path.
    ///
    /// `.yaml`, `.yml`, `.json`, and `.toml` are supported.
    pub fn load_from_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::load_from_path_with_limits(path, InputLimits::default())
    }

    /// Loads a document specification with explicit resource ceilings.
    pub fn load_from_path_with_limits(path: impl AsRef<Path>, limits: InputLimits) -> Result<Self> {
        let path = path.as_ref();
        let content = read_utf8_with_limit(path, limits.max_spec_bytes, "document spec")?;
        let extension = path
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();

        let mut spec = match extension.as_str() {
            "yaml" | "yml" | "" => Self::from_yaml_path(&content, Some(path), limits),
            "json" => Self::from_json_path(&content, Some(path), limits),
            "toml" => Self::from_toml_path(&content, Some(path), limits),
            other => Err(DocxError::parse(format!(
                "unsupported document spec extension '{other}', expected .yaml, .yml, .json, or .toml"
            ))),
        }?;
        spec.asset_base_dir = path.parent().map(Path::to_path_buf);
        Ok(spec)
    }

    /// Parses a YAML document specification string.
    pub fn from_yaml_str(content: &str) -> Result<Self> {
        Self::from_yaml_str_with_limits(content, InputLimits::default())
    }

    /// Parses a YAML document specification with explicit resource ceilings.
    pub fn from_yaml_str_with_limits(content: &str, limits: InputLimits) -> Result<Self> {
        ensure_spec_size(content, limits.max_spec_bytes, "YAML document spec")?;
        Self::from_yaml_path(content, None, limits)
    }

    /// Parses a JSON document specification string.
    pub fn from_json_str(content: &str) -> Result<Self> {
        Self::from_json_str_with_limits(content, InputLimits::default())
    }

    /// Parses a JSON document specification with explicit resource ceilings.
    pub fn from_json_str_with_limits(content: &str, limits: InputLimits) -> Result<Self> {
        ensure_spec_size(content, limits.max_spec_bytes, "JSON document spec")?;
        Self::from_json_path(content, None, limits)
    }

    /// Parses a TOML document specification string.
    pub fn from_toml_str(content: &str) -> Result<Self> {
        Self::from_toml_str_with_limits(content, InputLimits::default())
    }

    /// Parses a TOML document specification with explicit resource ceilings.
    pub fn from_toml_str_with_limits(content: &str, limits: InputLimits) -> Result<Self> {
        ensure_spec_size(content, limits.max_spec_bytes, "TOML document spec")?;
        Self::from_toml_path(content, None, limits)
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

        crate::io_utils::atomic_write(path, content.as_bytes())
    }

    /// Writes a commented YAML starter document template to disk.
    pub fn write_yaml_template(path: impl AsRef<Path>) -> Result<()> {
        let path = path.as_ref();
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent)?;
        }
        crate::io_utils::atomic_write(path, Self::default_yaml_template().as_bytes())
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

    fn from_yaml_path(
        content: &str,
        source_path: Option<&Path>,
        limits: InputLimits,
    ) -> Result<Self> {
        let expanded = expand_yaml_document_spec_with_limits(content, source_path, limits)?;
        serde_yaml::from_value(expanded)
            .map_err(|error| DocxError::parse(format!("invalid YAML document spec: {error}")))
    }

    fn from_json_path(
        content: &str,
        source_path: Option<&Path>,
        limits: InputLimits,
    ) -> Result<Self> {
        let root: serde_json::Value = serde_json::from_str(content)
            .map_err(|error| DocxError::parse(format!("invalid JSON document spec: {error}")))?;
        let root = serde_yaml::to_value(root).map_err(|error| {
            DocxError::parse(format!("failed to normalize JSON document spec: {error}"))
        })?;
        let expanded = expand_document_spec_value_with_limits(root, source_path, limits)?;
        serde_yaml::from_value(expanded)
            .map_err(|error| DocxError::parse(format!("invalid JSON document spec: {error}")))
    }

    fn from_toml_path(
        content: &str,
        source_path: Option<&Path>,
        limits: InputLimits,
    ) -> Result<Self> {
        let root: toml::Value = toml::from_str(content)
            .map_err(|error| DocxError::parse(format!("invalid TOML document spec: {error}")))?;
        let root = serde_yaml::to_value(root).map_err(|error| {
            DocxError::parse(format!("failed to normalize TOML document spec: {error}"))
        })?;
        let expanded = expand_document_spec_value_with_limits(root, source_path, limits)?;
        serde_yaml::from_value(expanded)
            .map_err(|error| DocxError::parse(format!("invalid TOML document spec: {error}")))
    }
}

fn ensure_spec_size(content: &str, limit: u64, label: &str) -> Result<()> {
    let bytes = u64::try_from(content.len()).unwrap_or(u64::MAX);
    if bytes > limit {
        return Err(DocxError::resource_limit(format!(
            "{label} is {bytes} bytes; limit is {limit} bytes"
        )));
    }
    Ok(())
}

fn read_utf8_with_limit(path: &Path, limit: u64, label: &str) -> Result<String> {
    let declared = fs::metadata(path)?.len();
    if declared > limit {
        return Err(DocxError::resource_limit(format!(
            "{label} '{}' is {declared} bytes; limit is {limit} bytes",
            path.display()
        )));
    }
    let mut bytes = Vec::new();
    fs::File::open(path)?
        .take(limit.saturating_add(1))
        .read_to_end(&mut bytes)?;
    if u64::try_from(bytes.len()).unwrap_or(u64::MAX) > limit {
        return Err(DocxError::resource_limit(format!(
            "{label} '{}' grew beyond the {limit} byte limit while reading",
            path.display()
        )));
    }
    String::from_utf8(bytes).map_err(|error| {
        DocxError::parse(format!(
            "{label} '{}' is not valid UTF-8: {error}",
            path.display()
        ))
    })
}

/// A high-level document block.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, JsonSchema)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum BlockSpec {
    /// Selects the cover title form.
    CoverTitle {
        /// Cover-title text.
        text: String,
    },
    /// Selects the title form.
    Title {
        /// Title text.
        text: String,
    },
    /// Selects the subtitle form.
    Subtitle {
        /// Subtitle text.
        text: String,
    },
    /// Selects the hero form.
    Hero {
        /// Hero statement text.
        text: String,
    },
    /// Selects the centered note form.
    CenteredNote {
        /// Centered note text.
        text: String,
    },
    /// Selects the page heading form.
    PageHeading {
        /// Page-heading text.
        text: String,
    },
    /// Selects the section form.
    Section {
        /// Section-heading text.
        text: String,
    },
    /// Selects the body form.
    Body {
        /// Body text.
        text: String,
    },
    /// Selects the tagline form.
    Tagline {
        /// Tagline text.
        text: String,
    },
    /// Selects the paragraph form.
    Paragraph {
        /// Rich paragraph specification.
        spec: ParagraphSpec,
    },
    /// Starts the following content on a new page.
    PageBreak,
    /// Starts the following content in a new next-page section.
    SectionBreak,
    /// Emits a Word-updatable TOC field and a deterministic PDF heading list.
    TableOfContents {
        #[serde(default)]
        /// Optional heading shown above the generated table of contents.
        title: Option<String>,
    },
    /// Selects the bullets form.
    Bullets {
        /// Ordered bullet item text.
        items: Vec<String>,
    },
    /// Selects the numbered form.
    Numbered {
        /// Ordered numbered item text.
        items: Vec<String>,
    },
    /// Selects the label values form.
    LabelValues {
        /// Ordered label/value pairs.
        items: Vec<LabelValueSpec>,
    },
    /// Selects the metrics form.
    Metrics {
        /// Ordered metric cards.
        items: Vec<MetricSpec>,
    },
    /// Selects the table form.
    Table {
        /// Table specification.
        spec: TableSpec,
    },
    /// Selects the image form.
    Image {
        #[serde(flatten)]
        /// Image source, dimensions, alignment, and alternative text.
        spec: VisualSpec,
    },
    /// Selects the logo form.
    Logo {
        #[serde(flatten)]
        /// Logo source, dimensions, alignment, and alternative text.
        spec: VisualSpec,
    },
    /// Selects the signature form.
    Signature {
        #[serde(flatten)]
        /// Signature source, dimensions, alignment, and alternative text.
        spec: VisualSpec,
    },
    /// Selects the chart form.
    Chart {
        #[serde(flatten)]
        /// Chart source, dimensions, alignment, and alternative text.
        spec: VisualSpec,
    },
    /// Selects the spacer form.
    Spacer,
}

/// A fully specified paragraph block.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize, JsonSchema)]
#[serde(default)]
pub struct ParagraphSpec {
    /// Ordered runs.
    pub runs: Vec<RunSpec>,
    /// Optional style identifier.
    pub style_id: Option<String>,
    /// Optional alignment.
    pub alignment: Option<ParagraphAlignmentSpec>,
    /// Spacing before, in twentieths of a point.
    pub spacing_before_twips: Option<u32>,
    /// Spacing after, in twentieths of a point.
    pub spacing_after_twips: Option<u32>,
    /// Whether the paragraph starts on a new page.
    pub page_break_before: bool,
    /// Whether the paragraph starts a new section.
    pub section_break_before: bool,
}

impl ParagraphSpec {
    /// Creates a value with default settings.
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
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum ParagraphAlignmentSpec {
    /// Selects the left form.
    Left,
    /// Selects the center form.
    Center,
    /// Selects the right form.
    Right,
    /// Selects the justified form.
    Justified,
}

/// A fully specified text run.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize, JsonSchema)]
#[serde(default)]
pub struct RunSpec {
    /// Text.
    pub text: String,
    /// Optional style identifier.
    pub style_id: Option<String>,
    /// Whether the run is bold.
    pub bold: bool,
    /// Whether the run is italic.
    pub italic: bool,
    /// Optional underline.
    pub underline: Option<UnderlineStyleSpec>,
    /// Whether the run is struck through.
    pub strikethrough: bool,
    /// Whether the run uses small capitals.
    pub small_caps: bool,
    /// Whether the run has a shadow effect.
    pub shadow: bool,
    /// Optional color.
    pub color: Option<String>,
    /// Optional font family.
    pub font_family: Option<String>,
    /// Optional size points.
    pub size_pt: Option<f32>,
    /// Optional vertical align.
    pub vertical_align: Option<VerticalAlignSpec>,
    /// Optional hyperlink.
    pub hyperlink: Option<String>,
    /// Optional bookmark.
    pub bookmark: Option<String>,
    /// Optional field.
    pub field: Option<RunFieldSpec>,
    /// Optional footnote.
    pub footnote: Option<String>,
}

/// Serializable dynamic field kinds.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum RunFieldSpec {
    /// Selects the table of contents form.
    TableOfContents,
}

/// A fully specified visual/image block.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize, JsonSchema)]
#[serde(default)]
pub struct VisualSpec {
    /// Path.
    pub path: String,
    /// Optional alt text.
    pub alt_text: Option<String>,
    /// Optional alignment.
    pub alignment: Option<ParagraphAlignmentSpec>,
    /// Width, in twentieths of a point.
    pub width_twips: Option<u32>,
    /// Height, in twentieths of a point.
    pub height_twips: Option<u32>,
    /// Maximum allowed width twips.
    pub max_width_twips: Option<u32>,
    /// Maximum allowed height twips.
    pub max_height_twips: Option<u32>,
}

impl VisualSpec {
    /// Creates a value with default settings.
    pub fn new(path: impl Into<String>) -> Self {
        Self {
            path: path.into(),
            ..Self::default()
        }
    }
}

impl RunSpec {
    /// Creates a value with default settings.
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            ..Self::default()
        }
    }
}

/// A serializable underline style.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum UnderlineStyleSpec {
    /// Selects the single form.
    Single,
    /// Selects the double form.
    Double,
    /// Selects the dotted form.
    Dotted,
    /// Selects the dash form.
    Dash,
    /// Selects the wavy form.
    Wavy,
    /// Selects the words form.
    Words,
    /// Selects the none form.
    None,
}

/// A serializable run vertical alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum VerticalAlignSpec {
    /// Selects the superscript form.
    Superscript,
    /// Selects the subscript form.
    Subscript,
    /// Selects the baseline form.
    Baseline,
}

/// A simple label-value pair block item.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, JsonSchema)]
pub struct LabelValueSpec {
    /// Label.
    pub label: String,
    /// Value.
    pub value: String,
}

/// A metric card item.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, JsonSchema)]
pub struct MetricSpec {
    /// Label.
    pub label: String,
    /// Value.
    pub value: String,
    /// Tone.
    pub tone: Tone,
}

/// Shared semantic color tone.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum Tone {
    /// Selects the positive form.
    Positive,
    /// Selects the neutral form.
    Neutral,
    /// Selects the warning form.
    Warning,
    /// Selects the risk form.
    Risk,
}

/// A grid table specification.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, JsonSchema)]
pub struct TableSpec {
    /// Optional style identifier.
    pub style_id: Option<String>,
    /// Ordered columns.
    pub columns: Vec<ColumnSpec>,
    /// Ordered rows.
    pub rows: Vec<RowSpec>,
}

/// A table column definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, JsonSchema)]
pub struct ColumnSpec {
    /// Label.
    pub label: String,
    /// Width.
    pub width: u32,
}

/// A table row definition.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize, JsonSchema)]
#[serde(default)]
pub struct RowSpec {
    /// Ordered cells.
    pub cells: Vec<CellSpec>,
    /// Whether this row repeats as a header on later pages.
    pub repeat_as_header: bool,
    /// Whether this row may split across pages.
    pub allow_split_across_pages: Option<bool>,
}

/// A table cell definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, JsonSchema)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum CellSpec {
    /// Selects the text form.
    Text {
        /// Plain cell text.
        text: String,
    },
    /// Selects the status form.
    Status(StatusSpec),
    /// Selects the rich form.
    Rich {
        #[serde(default)]
        /// Ordered paragraphs contained by the cell.
        paragraphs: Vec<ParagraphSpec>,
        #[serde(default)]
        /// Number of grid columns spanned by the cell.
        grid_span: Option<u32>,
        #[serde(default)]
        /// Optional cell background color as a six-digit RGB value.
        background_color: Option<String>,
        #[serde(default)]
        /// Optional nested table contained by the cell.
        nested_table: Option<Box<TableSpec>>,
    },
}

/// A status cell definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, JsonSchema)]
pub struct StatusSpec {
    /// Text.
    pub text: String,
    /// Tone.
    pub tone: Tone,
}

/// Document.
pub fn document<I>(blocks: I) -> DocumentSpec
where
    I: IntoIterator<Item = BlockSpec>,
{
    DocumentSpec {
        version: SPEC_VERSION,
        output_name: None,
        metadata: DocumentMetadata::default(),
        page_setup: None,
        header: None,
        footer: None,
        page_numbering: None,
        styles: Stylesheet::default(),
        blocks: blocks.into_iter().collect(),
        asset_base_dir: None,
    }
}

/// Cover title.
pub fn cover_title(text: impl Into<String>) -> BlockSpec {
    BlockSpec::CoverTitle { text: text.into() }
}

/// Title.
pub fn title(text: impl Into<String>) -> BlockSpec {
    BlockSpec::Title { text: text.into() }
}

/// Subtitle.
pub fn subtitle(text: impl Into<String>) -> BlockSpec {
    BlockSpec::Subtitle { text: text.into() }
}

/// Hero.
pub fn hero(text: impl Into<String>) -> BlockSpec {
    BlockSpec::Hero { text: text.into() }
}

/// Centered note.
pub fn centered_note(text: impl Into<String>) -> BlockSpec {
    BlockSpec::CenteredNote { text: text.into() }
}

/// Page heading.
pub fn page_heading(text: impl Into<String>) -> BlockSpec {
    BlockSpec::PageHeading { text: text.into() }
}

/// Section.
pub fn section(text: impl Into<String>) -> BlockSpec {
    BlockSpec::Section { text: text.into() }
}

/// Body.
pub fn body(text: impl Into<String>) -> BlockSpec {
    BlockSpec::Body { text: text.into() }
}

/// Tagline.
pub fn tagline(text: impl Into<String>) -> BlockSpec {
    BlockSpec::Tagline { text: text.into() }
}

/// Paragraph.
pub fn paragraph(spec: ParagraphSpec) -> BlockSpec {
    BlockSpec::Paragraph { spec }
}

/// Bullets.
pub fn bullets<I, S>(items: I) -> BlockSpec
where
    I: IntoIterator<Item = S>,
    S: Into<String>,
{
    BlockSpec::Bullets {
        items: items.into_iter().map(Into::into).collect(),
    }
}

/// Numbered.
pub fn numbered<I, S>(items: I) -> BlockSpec
where
    I: IntoIterator<Item = S>,
    S: Into<String>,
{
    BlockSpec::Numbered {
        items: items.into_iter().map(Into::into).collect(),
    }
}

/// Label values.
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

/// Metric.
pub fn metric(label: impl Into<String>, value: impl Into<String>, tone: Tone) -> MetricSpec {
    MetricSpec {
        label: label.into(),
        value: value.into(),
        tone,
    }
}

/// Metrics.
pub fn metrics<I>(items: I) -> BlockSpec
where
    I: IntoIterator<Item = MetricSpec>,
{
    BlockSpec::Metrics {
        items: items.into_iter().collect(),
    }
}

/// Col.
pub fn col(label: impl Into<String>, width: u32) -> ColumnSpec {
    ColumnSpec {
        label: label.into(),
        width,
    }
}

/// Text.
pub fn text(text: impl Into<String>) -> CellSpec {
    CellSpec::Text { text: text.into() }
}

/// Status.
pub fn status(text: impl Into<String>, tone: Tone) -> StatusSpec {
    StatusSpec {
        text: text.into(),
        tone,
    }
}

/// Row.
pub fn row<T>(value: T) -> RowSpec
where
    T: IntoRowSpec,
{
    value.into_row_spec()
}

/// Table.
pub fn table<C, R>(columns: C, rows: R) -> BlockSpec
where
    C: IntoIterator<Item = ColumnSpec>,
    R: IntoIterator<Item = RowSpec>,
{
    BlockSpec::Table {
        spec: TableSpec {
            style_id: None,
            columns: columns.into_iter().collect(),
            rows: rows.into_iter().collect(),
        },
    }
}

/// Image.
pub fn image(path: impl Into<String>) -> BlockSpec {
    BlockSpec::Image {
        spec: VisualSpec::new(path),
    }
}

/// Logo.
pub fn logo(path: impl Into<String>) -> BlockSpec {
    BlockSpec::Logo {
        spec: VisualSpec::new(path),
    }
}

/// Signature.
pub fn signature(path: impl Into<String>) -> BlockSpec {
    BlockSpec::Signature {
        spec: VisualSpec::new(path),
    }
}

/// Chart.
pub fn chart(path: impl Into<String>) -> BlockSpec {
    BlockSpec::Chart {
        spec: VisualSpec::new(path),
    }
}

/// Spacer.
pub fn spacer() -> BlockSpec {
    BlockSpec::Spacer
}

/// Defines into row spec.
pub trait IntoRowSpec {
    /// Into row spec.
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
                    ..RowSpec::default()
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

version: 1
output_name: my-document
# Optional core and custom metadata:
# metadata:
#   title: My Document
#   author: RusDox
#   subject: Quarterly review
#   keywords:
#     - planning
#     - board
#   custom_properties:
#     Client: Acme Corp
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
# Optional reusable named styles:
# styles:
#   paragraph:
#     - id: lead
#       based_on: Normal
#       paragraph:
#         alignment: center
#         spacing_after: 180
#       run:
#         bold: true
#         color: "0F172A"
#   run:
#     - id: accent
#       based_on: DefaultParagraphFont
#       properties:
#         italic: true
#         color: "AA5500"
# Optional YAML composition helpers:
# variables:
#   company: Acme Corp
#   regions:
#     - name: North America
#       owner: Maya
#     - name: EMEA
#       owner: Leon
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
    use crate::{
        Border, BorderStyle, DocumentMetadata, HeaderFooter, PageNumberFormat, PageNumbering,
        PageSetup, ParagraphAlignment, ParagraphList, ParagraphStyle, ParagraphStyleProperties,
        RunStyle, RunStyleProperties, Stylesheet, TableBorders, TableStyle, TableStyleProperties,
    };

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
    fn json_and_toml_authoring_share_deterministic_expressions() {
        let json = r#"{
  "version": 1,
  "variables": {"customer": {"name": "Northstar", "active": true}},
  "blocks": [{
    "type": "when",
    "path": "customer.active",
    "blocks": [{"type": "body", "text": "{{ customer.name | upper }}"}],
    "otherwise": []
  }]
}"#;
        let json_spec = super::DocumentSpec::from_json_str(json).expect("JSON authoring");
        assert_eq!(json_spec.blocks, vec![body("NORTHSTAR")]);

        let toml = r#"version = 1

[variables]
name = "Northstar"

[[blocks]]
type = "body"
text = "{{ name | lower }}"
"#;
        let toml_spec = super::DocumentSpec::from_toml_str(toml).expect("TOML authoring");
        assert_eq!(toml_spec.blocks, vec![body("northstar")]);
    }

    #[test]
    fn spec_round_trips_through_yaml() {
        let spec = super::DocumentSpec {
            version: super::SPEC_VERSION,
            output_name: Some("hello-world".to_string()),
            metadata: DocumentMetadata::new()
                .title("Hello World")
                .author("RusDox")
                .subject("Round-trip"),
            page_setup: Some(PageSetup::new(11_880, 16_380).margins(900, 1_000, 1_100, 1_200)),
            header: Some(
                HeaderFooter::new("Board Report").with_alignment(ParagraphAlignment::Center),
            ),
            footer: Some(
                HeaderFooter::new("Page {page} of {pages}")
                    .with_alignment(ParagraphAlignment::Right),
            ),
            page_numbering: Some(PageNumbering::new(PageNumberFormat::UpperRoman).start_at(3)),
            styles: crate::Stylesheet::default(),
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
    fn spec_round_trips_named_styles_and_style_references() {
        let border = Border::new(BorderStyle::Single).size(8).color("CBD5E1");
        let spec = super::DocumentSpec {
            version: super::SPEC_VERSION,
            output_name: Some("styled-spec".to_string()),
            metadata: DocumentMetadata::default(),
            page_setup: None,
            header: None,
            footer: None,
            page_numbering: None,
            styles: Stylesheet::new()
                .add_paragraph_style(
                    ParagraphStyle::new("lead")
                        .based_on("Normal")
                        .next("body")
                        .paragraph(ParagraphStyleProperties {
                            list: Some(ParagraphList::bullet_with_id(7)),
                            alignment: Some(ParagraphAlignment::Center),
                            spacing_before: Some(120),
                            spacing_after: Some(240),
                            keep_next: Some(true),
                            page_break_before: Some(false),
                        })
                        .run(RunStyleProperties::new().bold().color("0F172A")),
                )
                .add_run_style(
                    RunStyle::new("accent")
                        .based_on("DefaultParagraphFont")
                        .properties(RunStyleProperties::new().italic().color("AA5500")),
                )
                .add_table_style(
                    TableStyle::new("grid").based_on("TableNormal").properties(
                        TableStyleProperties::new()
                            .width(9_360)
                            .borders(TableBorders::new().top(border)),
                    ),
                ),
            blocks: vec![
                paragraph(ParagraphSpec {
                    style_id: Some("lead".to_string()),
                    runs: vec![RunSpec {
                        text: "Styled".to_string(),
                        style_id: Some("accent".to_string()),
                        ..RunSpec::default()
                    }],
                    ..ParagraphSpec::default()
                }),
                BlockSpec::Table {
                    spec: super::TableSpec {
                        style_id: Some("grid".to_string()),
                        columns: vec![
                            super::ColumnSpec {
                                label: "Metric".to_string(),
                                width: 4_680,
                            },
                            super::ColumnSpec {
                                label: "Value".to_string(),
                                width: 4_680,
                            },
                        ],
                        rows: vec![super::RowSpec {
                            cells: vec![
                                super::CellSpec::Text {
                                    text: "ARR".to_string(),
                                },
                                super::CellSpec::Text {
                                    text: "$18.7M".to_string(),
                                },
                            ],
                            ..super::RowSpec::default()
                        }],
                    },
                },
            ],
            asset_base_dir: None,
        };

        let yaml = spec.to_yaml_string().expect("serialize yaml");
        assert!(yaml.contains("styles:"));
        assert!(yaml.contains("style_id: lead"));
        assert!(yaml.contains("style_id: accent"));
        assert!(yaml.contains("style_id: grid"));
        assert!(yaml.contains("based_on: Normal"));
        assert!(yaml.contains("based_on: DefaultParagraphFont"));
        assert!(yaml.contains("based_on: TableNormal"));

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
    fn load_from_path_expands_yaml_variables_includes_repeaters_and_metadata() {
        let temp = tempdir().expect("temp dir");
        let fragment_path = temp.path().join("summary.yaml");
        let spec_path = temp.path().join("spec.yaml");
        fs::write(
            &fragment_path,
            r#"variables:
  intro: Summary for {{client}}
blocks:
  - type: body
    text: "{{intro}}"
"#,
        )
        .expect("write fragment");
        fs::write(
            &spec_path,
            r#"output_name: regional-plan
metadata:
  title: "{{client}} Regional Plan"
  author: Strategy Team
  subject: "{{quarter}} rollout"
  keywords:
    - "{{quarter}}"
    - planning
  custom_properties:
    Client: "{{client}}"
variables:
  client: Acme
  quarter: Q2
  regions:
    - name: North America
      owner: Maya
    - name: EMEA
      owner: Leon
blocks:
  - type: title
    text: "{{client}} Regional Plan"
  - type: include
    path: summary.yaml
  - type: repeat
    variable: regions
    as: region
    blocks:
      - type: section
        text: "{{region.name}}"
      - type: body
        text: "Owner: {{region.owner}}"
"#,
        )
        .expect("write spec");

        let spec = super::DocumentSpec::load_from_path(&spec_path).expect("load expanded yaml");
        assert_eq!(spec.metadata.title.as_deref(), Some("Acme Regional Plan"));
        assert_eq!(spec.metadata.subject.as_deref(), Some("Q2 rollout"));
        assert_eq!(spec.metadata.keywords, vec!["Q2", "planning"]);
        assert_eq!(
            spec.metadata
                .custom_properties
                .get("Client")
                .map(String::as_str),
            Some("Acme")
        );
        assert_eq!(spec.blocks.len(), 6);
    }

    #[test]
    fn tuple_rows_accept_plain_text_and_status_cells() {
        let row = row(("ARR", "$18.7M", status("Strong", Tone::Positive), "On plan"));
        assert_eq!(row.cells.len(), 4);
    }
}
