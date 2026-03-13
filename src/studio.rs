//! Configurable document composition and PDF preview helpers.
//!
//! This module is intended to be the core orchestration layer for projects
//! that build rich `.docx` files programmatically and optionally emit PDF
//! previews.

use std::collections::{BTreeMap, HashMap};
use std::fmt::Write as _;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::OnceLock;
use std::time::{Duration, Instant};

use fontdb::{Database as FontDatabase, Family as FontFamily, Query as FontQuery};
use miniz_oxide::deflate::{compress_to_vec_zlib, CompressionLevel};
use pdf_writer::types::{CidFontType, FontFlags, SystemInfo, UnicodeCmap};
use pdf_writer::{Filter, Finish, Name, Pdf, Rect, Ref, Str};

use crate::{
    config::RusdoxConfig,
    spec::{
        BlockSpec, CellSpec, DocumentSpec, ParagraphAlignmentSpec,
        ParagraphSpec as BlockParagraphSpec, RunSpec as BlockRunSpec, TableSpec as BlockTableSpec,
        Tone, UnderlineStyleSpec, VerticalAlignSpec, VisualSpec as BlockVisualSpec,
    },
    Border, BorderStyle, Document, DocumentBlockRef, Paragraph, ParagraphAlignment, ParagraphList,
    Run, RunProperties, Stylesheet, Table, TableBorders, TableCell, TableRow, UnderlineStyle,
    VerticalAlign, Visual, VisualKind,
};

/// Default config file expected in the current working directory.
pub const DEFAULT_CONFIG_FILE: &str = "rusdox.toml";

const DEFAULT_PAGE_WIDTH: f32 = 612.0;
const DEFAULT_PAGE_HEIGHT: f32 = 792.0;
const DEFAULT_PAGE_MARGIN_X: f32 = 54.0;
const DEFAULT_PAGE_MARGIN_TOP: f32 = 54.0;
const DEFAULT_PAGE_MARGIN_BOTTOM: f32 = 54.0;
const DEFAULT_TEXT_SIZE: f32 = 11.0;
const DEFAULT_LINE_HEIGHT: f32 = 14.0;
const DEFAULT_LINE_HEIGHT_MULTIPLIER: f32 = 1.35;
const DEFAULT_BASELINE_FACTOR: f32 = 0.82;
const DEFAULT_TEXT_WIDTH_BIAS_REGULAR: f32 = 1.0;
const DEFAULT_TEXT_WIDTH_BIAS_BOLD: f32 = 1.03;
const DEFAULT_TABLE_ROW_PADDING_X: f32 = 7.0;
const DEFAULT_TABLE_ROW_PADDING_Y: f32 = 6.0;
const DEFAULT_TABLE_AFTER_SPACING: f32 = 12.0;
const DEFAULT_TABLE_GRID_STROKE_WIDTH: f32 = 0.75;
const MIN_CONTENT_WIDTH: f32 = 24.0;

/// Timings and output sizes captured while writing a companion DOCX and PDF.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct OutputStats {
    /// Time spent writing the `.docx` archive.
    pub docx_write: Duration,
    /// Time spent writing the `.pdf` preview.
    pub pdf_render: Duration,
    /// Final `.docx` size in bytes.
    pub docx_bytes: u64,
    /// Final `.pdf` size in bytes.
    pub pdf_bytes: u64,
}

/// Deep ink color for default presets.
pub const INK: &str = "0F172A";
/// Slate body-text color for default presets.
pub const SLATE: &str = "475569";
/// Muted caption color for default presets.
pub const MUTED: &str = "64748B";
/// Accent color for default presets.
pub const ACCENT: &str = "0F766E";
/// Gold accent for default presets.
pub const GOLD: &str = "B45309";
/// Red accent for default presets.
pub const RED: &str = "B91C1C";
/// Green accent for default presets.
pub const GREEN: &str = "166534";
/// Soft border/background color for default presets.
pub const SOFT: &str = "E2E8F0";
/// Pale surface color for default presets.
pub const PALE: &str = "F8FAFC";
/// Mint surface color for default presets.
pub const MINT: &str = "DCFCE7";
/// Amber surface color for default presets.
pub const AMBER: &str = "FEF3C7";
/// Rose surface color for default presets.
pub const ROSE: &str = "FEE2E2";

/// High-level composition helper built around a [`RusdoxConfig`].
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Studio {
    config: RusdoxConfig,
}

impl Studio {
    /// Creates a new studio with an explicit configuration.
    pub fn new(config: RusdoxConfig) -> Self {
        Self { config }
    }

    /// Returns the active configuration.
    pub fn config(&self) -> &RusdoxConfig {
        &self.config
    }

    /// Loads a studio from a config file.
    pub fn from_config_path(path: impl AsRef<Path>) -> crate::Result<Self> {
        let config = RusdoxConfig::load_from_path(path)?;
        Ok(Self::new(config))
    }

    /// Loads a studio from a config file when present, otherwise uses defaults.
    pub fn from_config_path_or_default(path: impl AsRef<Path>) -> crate::Result<Self> {
        let config = RusdoxConfig::load_from_path_or_default(path)?;
        Ok(Self::new(config))
    }

    /// Loads `./rusdox.toml` when present, otherwise the user-level config, otherwise defaults.
    pub fn from_default_file_or_default() -> crate::Result<Self> {
        let config = RusdoxConfig::load_local_or_user_default(DEFAULT_CONFIG_FILE)?;
        Ok(Self::new(config))
    }

    /// Writes the commented default template to a path.
    pub fn write_default_config(path: impl AsRef<Path>) -> crate::Result<()> {
        RusdoxConfig::write_toml_template(path)
    }

    /// Saves a document using the configured output directories and file stem.
    pub fn save_named(&self, document: &Document, name: &str) -> crate::Result<OutputStats> {
        let docx_path = self.docx_output_path(name);
        self.save_with_pdf_stats(document, docx_path)
    }

    /// Saves a document using the configured output directories and file stem without printing paths.
    pub fn save_named_quiet(&self, document: &Document, name: &str) -> crate::Result<OutputStats> {
        let docx_path = self.docx_output_path(name);
        self.save_with_pdf_stats_quiet(document, docx_path)
    }

    /// Renders a high-level document specification into a [`Document`].
    pub fn compose(&self, spec: &DocumentSpec) -> Document {
        let mut document = Document::new();
        if let Some(page_setup) = spec.page_setup.clone() {
            document.set_page_setup(page_setup);
        }
        document.set_metadata(spec.metadata.clone());
        document.set_header(spec.header.clone());
        document.set_footer(spec.footer.clone());
        document.set_page_numbering(spec.page_numbering.clone());
        document.set_styles(spec.styles.clone());
        self.append_spec(&mut document, spec);
        document
    }

    /// Appends a high-level document specification to an existing [`Document`].
    pub fn append_spec(&self, document: &mut Document, spec: &DocumentSpec) {
        let asset_base_dir = spec.asset_base_dir().map(Path::to_path_buf);
        for (index, block) in spec.blocks.iter().enumerate() {
            self.push_spec_block(
                document,
                block,
                spec.blocks.get(index + 1),
                asset_base_dir.as_deref(),
            );
        }
    }

    /// Renders and saves a high-level document specification by logical name.
    pub fn save_spec_named(&self, spec: &DocumentSpec, name: &str) -> crate::Result<OutputStats> {
        let document = self.compose(spec);
        self.save_named(&document, name)
    }

    /// Writes DOCX and optional PDF output.
    pub fn save_with_pdf(
        &self,
        document: &Document,
        docx_path: impl AsRef<Path>,
    ) -> crate::Result<()> {
        let _stats = self.save_with_pdf_stats(document, docx_path)?;
        Ok(())
    }

    /// Writes DOCX and optional PDF output without printing artifact paths.
    pub fn save_with_pdf_quiet(
        &self,
        document: &Document,
        docx_path: impl AsRef<Path>,
    ) -> crate::Result<()> {
        let _stats = self.save_with_pdf_stats_quiet(document, docx_path)?;
        Ok(())
    }

    /// Writes DOCX and optional PDF output and returns timing stats.
    pub fn save_with_pdf_stats(
        &self,
        document: &Document,
        docx_path: impl AsRef<Path>,
    ) -> crate::Result<OutputStats> {
        self.save_with_pdf_stats_impl(document, docx_path, true)
    }

    /// Writes DOCX and optional PDF output and returns timing stats without printing artifact paths.
    pub fn save_with_pdf_stats_quiet(
        &self,
        document: &Document,
        docx_path: impl AsRef<Path>,
    ) -> crate::Result<OutputStats> {
        self.save_with_pdf_stats_impl(document, docx_path, false)
    }

    fn save_with_pdf_stats_impl(
        &self,
        document: &Document,
        docx_path: impl AsRef<Path>,
        announce_outputs: bool,
    ) -> crate::Result<OutputStats> {
        let docx_path = docx_path.as_ref();
        if let Some(parent) = docx_path.parent() {
            fs::create_dir_all(parent)?;
        }

        let docx_start = Instant::now();
        document.save(docx_path)?;
        let docx_write = docx_start.elapsed();
        let docx_bytes = fs::metadata(docx_path)?.len();

        let mut pdf_render = Duration::ZERO;
        let mut pdf_bytes = 0_u64;
        if self.config.output.emit_pdf_preview {
            let rendered_dir = Path::new(&self.config.output.pdf_dir);
            fs::create_dir_all(rendered_dir)?;
            let stem = docx_path
                .file_stem()
                .ok_or_else(|| crate::DocxError::parse("invalid output file name"))?;
            let mut pdf_path = rendered_dir.join(stem);
            pdf_path.set_extension("pdf");
            let pdf_start = Instant::now();
            render_pdf(document, &pdf_path, &self.config)?;
            pdf_render = pdf_start.elapsed();
            pdf_bytes = fs::metadata(&pdf_path)?.len();
            if announce_outputs {
                println!("{}", docx_path.display());
                println!("{}", pdf_path.display());
            }
        } else if announce_outputs {
            println!("{}", docx_path.display());
        }

        Ok(OutputStats {
            docx_write,
            pdf_render,
            docx_bytes,
            pdf_bytes,
        })
    }

    /// Builds a base text run using the configured font family and body size.
    pub fn text_run(&self, text: impl Into<String>) -> Run {
        Run::from_text(text)
            .font(self.config.typography.font_family.clone())
            .size_points(points_to_u16(self.config.typography.body_size_pt))
    }

    /// Builds a centered cover title paragraph.
    pub fn cover_title(&self, text: &str) -> Paragraph {
        Paragraph::new()
            .with_alignment(ParagraphAlignment::Center)
            .spacing_before(self.config.spacing.cover_title_before_twips)
            .spacing_after(self.config.spacing.cover_title_after_twips)
            .add_run(
                self.text_run(text)
                    .size_points(points_to_u16(self.config.typography.cover_title_size_pt))
                    .bold()
                    .color(&self.config.colors.ink),
            )
    }

    /// Builds a centered title paragraph.
    pub fn title(&self, text: &str) -> Paragraph {
        Paragraph::new()
            .with_alignment(ParagraphAlignment::Center)
            .spacing_before(self.config.spacing.title_before_twips)
            .spacing_after(self.config.spacing.title_after_twips)
            .add_run(
                Run::from_text(text)
                    .font(self.config.typography.font_family.clone())
                    .size_points(points_to_u16(self.config.typography.title_size_pt))
                    .bold()
                    .color(&self.config.colors.ink),
            )
    }

    /// Builds a centered subtitle paragraph.
    pub fn subtitle(&self, text: &str) -> Paragraph {
        Paragraph::new()
            .with_alignment(ParagraphAlignment::Center)
            .spacing_after(self.config.spacing.subtitle_after_twips)
            .add_run(
                Run::from_text(text)
                    .font(self.config.typography.font_family.clone())
                    .size_points(points_to_u16(self.config.typography.subtitle_size_pt))
                    .color(&self.config.colors.muted),
            )
    }

    /// Builds a centered hero paragraph.
    pub fn hero(&self, text: &str) -> Paragraph {
        Paragraph::new()
            .with_alignment(ParagraphAlignment::Center)
            .spacing_after(self.config.spacing.hero_after_twips)
            .add_run(
                Run::from_text(text)
                    .font(self.config.typography.font_family.clone())
                    .size_points(points_to_u16(self.config.typography.hero_size_pt))
                    .bold()
                    .color(&self.config.colors.accent),
            )
    }

    /// Builds a section heading paragraph.
    pub fn section(&self, text: &str) -> Paragraph {
        Paragraph::new()
            .spacing_before(self.config.spacing.section_before_twips)
            .spacing_after(self.config.spacing.section_after_twips)
            .add_run(
                Run::from_text(text)
                    .font(self.config.typography.font_family.clone())
                    .size_points(points_to_u16(self.config.typography.section_size_pt))
                    .bold()
                    .underline(UnderlineStyle::Single)
                    .color(&self.config.colors.ink),
            )
    }

    /// Builds a body paragraph.
    pub fn body(&self, text: &str) -> Paragraph {
        Paragraph::new()
            .spacing_after(self.config.spacing.body_after_twips)
            .add_run(
                Run::from_text(text)
                    .font(self.config.typography.font_family.clone())
                    .size_points(points_to_u16(self.config.typography.body_size_pt))
                    .color(&self.config.colors.slate),
            )
    }

    fn list_item(&self, text: &str, list: ParagraphList) -> Paragraph {
        Paragraph::new()
            .with_list(list)
            .spacing_after(self.config.spacing.bullet_after_twips)
            .add_run(
                Run::from_text(text)
                    .font(self.config.typography.font_family.clone())
                    .size_points(points_to_u16(self.config.typography.body_size_pt))
                    .color(&self.config.colors.slate),
            )
    }

    /// Builds a bullet paragraph.
    pub fn bullet(&self, text: &str) -> Paragraph {
        self.list_item(text, ParagraphList::bullet())
    }

    /// Builds a numbered list paragraph.
    pub fn numbered(&self, text: &str) -> Paragraph {
        self.list_item(text, ParagraphList::numbered())
    }

    fn list_with_id(&self, text: &str, list: ParagraphList) -> Paragraph {
        self.list_item(text, list)
    }

    fn next_list_id(document: &Document) -> u32 {
        document
            .paragraphs()
            .filter_map(|paragraph| paragraph.list().map(|list| list.id()))
            .chain(
                document
                    .styles()
                    .paragraph_styles()
                    .filter_map(|style| style.paragraph.list.map(|list| list.id())),
            )
            .max()
            .unwrap_or(0)
            + 1
    }

    /// Builds a label-value paragraph.
    pub fn label_value(&self, label: &str, value: &str) -> Paragraph {
        Paragraph::new()
            .spacing_after(self.config.spacing.label_value_after_twips)
            .add_run(
                Run::from_text(format!("{label}: "))
                    .font(self.config.typography.font_family.clone())
                    .size_points(points_to_u16(self.config.typography.body_size_pt))
                    .bold()
                    .color(&self.config.colors.ink),
            )
            .add_run(
                Run::from_text(value)
                    .font(self.config.typography.font_family.clone())
                    .size_points(points_to_u16(self.config.typography.body_size_pt))
                    .color(&self.config.colors.slate),
            )
    }

    /// Builds a spacer paragraph.
    pub fn spacer(&self) -> Paragraph {
        Paragraph::new().spacing_after(self.config.spacing.spacer_after_twips)
    }

    /// Builds a centered note paragraph.
    pub fn centered_note(&self, text: &str) -> Paragraph {
        Paragraph::new()
            .with_alignment(ParagraphAlignment::Center)
            .spacing_after(self.config.spacing.note_after_twips)
            .add_run(
                Run::from_text(text)
                    .font(self.config.typography.font_family.clone())
                    .size_points(points_to_u16(self.config.typography.note_size_pt))
                    .italic()
                    .color(&self.config.colors.muted),
            )
    }

    /// Builds a page-heading paragraph that starts on a new page.
    pub fn page_heading(&self, text: &str) -> Paragraph {
        Paragraph::new()
            .page_break_before()
            .spacing_after(self.config.spacing.page_heading_after_twips)
            .add_run(
                self.text_run(text)
                    .size_points(points_to_u16(self.config.typography.page_heading_size_pt))
                    .bold()
                    .color(&self.config.colors.ink),
            )
    }

    /// Builds a centered accent tagline paragraph.
    pub fn tagline(&self, text: &str) -> Paragraph {
        Paragraph::new()
            .with_alignment(ParagraphAlignment::Center)
            .spacing_after(self.config.spacing.tagline_after_twips)
            .add_run(
                self.text_run(text)
                    .size_points(points_to_u16(self.config.typography.tagline_size_pt))
                    .italic()
                    .color(&self.config.colors.accent),
            )
    }

    /// Builds a metric cell.
    pub fn metric_cell(&self, label: &str, value: &str, background: &str) -> TableCell {
        TableCell::new()
            .width(self.config.table.metric_cell_width_twips)
            .background(background)
            .borders(self.card_borders())
            .add_paragraph(
                Paragraph::new()
                    .with_alignment(ParagraphAlignment::Center)
                    .spacing_before(self.config.spacing.metric_label_before_twips)
                    .spacing_after(self.config.spacing.metric_label_after_twips)
                    .add_run(
                        Run::from_text(label)
                            .font(self.config.typography.font_family.clone())
                            .size_points(points_to_u16(self.config.typography.metric_label_size_pt))
                            .bold()
                            .color(&self.config.colors.muted),
                    ),
            )
            .add_paragraph(
                Paragraph::new()
                    .with_alignment(ParagraphAlignment::Center)
                    .spacing_after(self.config.spacing.metric_value_after_twips)
                    .add_run(
                        Run::from_text(value)
                            .font(self.config.typography.font_family.clone())
                            .size_points(points_to_u16(self.config.typography.metric_value_size_pt))
                            .bold()
                            .color(&self.config.colors.ink),
                    ),
            )
    }

    /// Builds a header cell.
    pub fn header_cell(&self, text: &str, width: u32) -> TableCell {
        TableCell::new()
            .width(width)
            .background(&self.config.colors.soft)
            .borders(self.grid_borders())
            .add_paragraph(
                Paragraph::new()
                    .spacing_before(self.config.spacing.table_header_before_twips)
                    .spacing_after(self.config.spacing.table_header_after_twips)
                    .add_run(
                        Run::from_text(text)
                            .font(self.config.typography.font_family.clone())
                            .size_points(points_to_u16(self.config.typography.table_size_pt))
                            .bold()
                            .color(&self.config.colors.ink),
                    ),
            )
    }

    /// Builds a table data cell.
    pub fn data_cell(&self, text: &str, width: u32) -> TableCell {
        TableCell::new()
            .width(width)
            .background(&self.config.colors.pale)
            .borders(self.grid_borders())
            .add_paragraph(
                Paragraph::new()
                    .spacing_before(self.config.spacing.table_data_before_twips)
                    .spacing_after(self.config.spacing.table_data_after_twips)
                    .add_run(
                        Run::from_text(text)
                            .font(self.config.typography.font_family.clone())
                            .size_points(points_to_u16(self.config.typography.table_size_pt))
                            .color(&self.config.colors.slate),
                    ),
            )
    }

    /// Builds a status cell.
    pub fn status_cell(
        &self,
        text: &str,
        width: u32,
        background: &str,
        foreground: &str,
    ) -> TableCell {
        TableCell::new()
            .width(width)
            .background(background)
            .borders(self.grid_borders())
            .add_paragraph(
                Paragraph::new()
                    .with_alignment(ParagraphAlignment::Center)
                    .spacing_before(self.config.spacing.table_status_before_twips)
                    .spacing_after(self.config.spacing.table_status_after_twips)
                    .add_run(
                        Run::from_text(text)
                            .font(self.config.typography.font_family.clone())
                            .size_points(points_to_u16(self.config.typography.table_size_pt))
                            .bold()
                            .color(foreground),
                    ),
            )
    }

    /// Returns standard grid borders.
    pub fn grid_borders(&self) -> TableBorders {
        let border = Border::new(BorderStyle::Single)
            .size(clamp_u32_to_u16(
                self.config.table.grid_border_size_eighth_pt,
            ))
            .color(&self.config.colors.table_border);
        TableBorders::new()
            .top(border.clone())
            .bottom(border.clone())
            .left(border.clone())
            .right(border.clone())
            .inside_horizontal(border.clone())
            .inside_vertical(border)
    }

    /// Returns card borders.
    pub fn card_borders(&self) -> TableBorders {
        let border = Border::new(BorderStyle::Single)
            .size(clamp_u32_to_u16(
                self.config.table.card_border_size_eighth_pt,
            ))
            .color(&self.config.colors.table_border);
        TableBorders::new()
            .top(border.clone())
            .bottom(border.clone())
            .left(border.clone())
            .right(border)
    }

    fn docx_output_path(&self, name: &str) -> PathBuf {
        let stem = name.strip_suffix(".docx").unwrap_or(name);
        let mut path = PathBuf::from(&self.config.output.docx_dir);
        path.push(stem);
        path.set_extension("docx");
        path
    }

    fn push_spec_block(
        &self,
        document: &mut Document,
        block: &BlockSpec,
        next_block: Option<&BlockSpec>,
        asset_base_dir: Option<&Path>,
    ) {
        match block {
            BlockSpec::CoverTitle { text } => {
                document.push_paragraph(self.decorate_spec_paragraph(
                    self.cover_title(text),
                    block,
                    next_block,
                ));
            }
            BlockSpec::Title { text } => {
                document.push_paragraph(self.decorate_spec_paragraph(
                    self.title(text),
                    block,
                    next_block,
                ));
            }
            BlockSpec::Subtitle { text } => {
                document.push_paragraph(self.decorate_spec_paragraph(
                    self.subtitle(text),
                    block,
                    next_block,
                ));
            }
            BlockSpec::Hero { text } => {
                document.push_paragraph(self.decorate_spec_paragraph(
                    self.hero(text),
                    block,
                    next_block,
                ));
            }
            BlockSpec::CenteredNote { text } => {
                document.push_paragraph(self.decorate_spec_paragraph(
                    self.centered_note(text),
                    block,
                    next_block,
                ));
            }
            BlockSpec::PageHeading { text } => {
                document.push_paragraph(self.decorate_spec_paragraph(
                    self.page_heading(text),
                    block,
                    next_block,
                ));
            }
            BlockSpec::Section { text } => {
                document.push_paragraph(self.decorate_spec_paragraph(
                    self.section(text),
                    block,
                    next_block,
                ));
            }
            BlockSpec::Body { text } => {
                document.push_paragraph(self.decorate_spec_paragraph(
                    self.body(text),
                    block,
                    next_block,
                ));
            }
            BlockSpec::Tagline { text } => {
                document.push_paragraph(self.decorate_spec_paragraph(
                    self.tagline(text),
                    block,
                    next_block,
                ));
            }
            BlockSpec::Paragraph { spec } => {
                document.push_paragraph(self.decorate_spec_paragraph(
                    self.paragraph_from_spec(spec),
                    block,
                    next_block,
                ));
            }
            BlockSpec::Bullets { items } => {
                let list_id = Self::next_list_id(document);
                for item in items {
                    document.push_paragraph(
                        self.list_with_id(item, ParagraphList::bullet_with_id(list_id)),
                    );
                }
            }
            BlockSpec::Numbered { items } => {
                let list_id = Self::next_list_id(document);
                for item in items {
                    document.push_paragraph(
                        self.list_with_id(item, ParagraphList::numbered_with_id(list_id)),
                    );
                }
            }
            BlockSpec::LabelValues { items } => {
                for item in items {
                    document.push_paragraph(self.label_value(&item.label, &item.value));
                }
            }
            BlockSpec::Metrics { items } => {
                let mut row = TableRow::new();
                for item in items {
                    row = row.add_cell(self.metric_cell(
                        &item.label,
                        &item.value,
                        self.tone_background(item.tone),
                    ));
                }
                document.push_table(
                    Table::new()
                        .width(self.config.table.default_width_twips)
                        .add_row(row),
                );
            }
            BlockSpec::Table { spec } => {
                document.push_table(self.table_from_spec(spec));
            }
            BlockSpec::Image { spec } => {
                document.push_visual(self.visual_from_spec(
                    spec,
                    VisualKind::Image,
                    asset_base_dir,
                ));
            }
            BlockSpec::Logo { spec } => {
                document.push_visual(self.visual_from_spec(spec, VisualKind::Logo, asset_base_dir));
            }
            BlockSpec::Signature { spec } => {
                document.push_visual(self.visual_from_spec(
                    spec,
                    VisualKind::Signature,
                    asset_base_dir,
                ));
            }
            BlockSpec::Chart { spec } => {
                document.push_visual(self.visual_from_spec(
                    spec,
                    VisualKind::Chart,
                    asset_base_dir,
                ));
            }
            BlockSpec::Spacer => {
                document.push_paragraph(self.spacer());
            }
        }
    }

    fn decorate_spec_paragraph(
        &self,
        paragraph: Paragraph,
        block: &BlockSpec,
        next_block: Option<&BlockSpec>,
    ) -> Paragraph {
        if should_keep_next(block, next_block) {
            paragraph.keep_next()
        } else {
            paragraph
        }
    }

    fn paragraph_from_spec(&self, spec: &BlockParagraphSpec) -> Paragraph {
        let mut paragraph = Paragraph::new();

        if let Some(style_id) = &spec.style_id {
            paragraph = paragraph.with_style(style_id.clone());
        }

        if let Some(alignment) = &spec.alignment {
            paragraph = paragraph.with_alignment(match alignment {
                ParagraphAlignmentSpec::Left => ParagraphAlignment::Left,
                ParagraphAlignmentSpec::Center => ParagraphAlignment::Center,
                ParagraphAlignmentSpec::Right => ParagraphAlignment::Right,
                ParagraphAlignmentSpec::Justified => ParagraphAlignment::Justified,
            });
        }

        if let Some(spacing_before) = spec.spacing_before_twips {
            paragraph = paragraph.spacing_before(spacing_before);
        }

        if let Some(spacing_after) = spec.spacing_after_twips {
            paragraph = paragraph.spacing_after(spacing_after);
        }

        if spec.page_break_before {
            paragraph = paragraph.page_break_before();
        }

        for run in &spec.runs {
            paragraph = paragraph.add_run(self.run_from_spec(run));
        }

        paragraph
    }

    fn run_from_spec(&self, spec: &BlockRunSpec) -> Run {
        let mut run = self.text_run(&spec.text);

        if let Some(style_id) = &spec.style_id {
            run = run.with_style(style_id.clone());
        }

        if spec.bold {
            run = run.bold();
        }
        if spec.italic {
            run = run.italic();
        }
        if let Some(underline) = &spec.underline {
            run = run.underline(match underline {
                UnderlineStyleSpec::Single => UnderlineStyle::Single,
                UnderlineStyleSpec::Double => UnderlineStyle::Double,
                UnderlineStyleSpec::Dotted => UnderlineStyle::Dotted,
                UnderlineStyleSpec::Dash => UnderlineStyle::Dash,
                UnderlineStyleSpec::Wavy => UnderlineStyle::Wavy,
                UnderlineStyleSpec::Words => UnderlineStyle::Words,
                UnderlineStyleSpec::None => UnderlineStyle::None,
            });
        }
        if spec.strikethrough {
            run = run.strikethrough();
        }
        if spec.small_caps {
            run = run.small_caps();
        }
        if spec.shadow {
            run = run.shadow();
        }
        if let Some(color) = &spec.color {
            run = run.color(color);
        }
        if let Some(font_family) = &spec.font_family {
            run = run.font(font_family.clone());
        }
        if let Some(size_pt) = spec.size_pt {
            run = run.size_points(points_to_u16(size_pt));
        }
        if let Some(vertical_align) = spec.vertical_align {
            run = match vertical_align {
                VerticalAlignSpec::Superscript => run.superscript(),
                VerticalAlignSpec::Subscript => run.subscript(),
                VerticalAlignSpec::Baseline => {
                    let mut baseline = run;
                    baseline.properties_mut().vertical_align = Some(VerticalAlign::Baseline);
                    baseline
                }
            };
        }

        run
    }

    fn table_from_spec(&self, spec: &BlockTableSpec) -> Table {
        let total_width = spec.columns.iter().map(|column| column.width).sum::<u32>();
        let mut table = Table::new().add_row(
            spec.columns
                .iter()
                .fold(TableRow::new().repeat_as_header(), |row, column| {
                    row.add_cell(self.header_cell(&column.label, column.width))
                }),
        );

        if let Some(style_id) = &spec.style_id {
            table = table.style(style_id.clone());
        } else {
            table = table
                .width(if total_width == 0 {
                    self.config.table.default_width_twips
                } else {
                    total_width
                })
                .borders(self.grid_borders());
        }

        for row_spec in &spec.rows {
            let mut row = TableRow::new();

            for (index, column) in spec.columns.iter().enumerate() {
                let cell = row_spec.cells.get(index);
                row = row.add_cell(match cell {
                    Some(CellSpec::Text { text }) => self.data_cell(text, column.width),
                    Some(CellSpec::Status(status)) => self.status_cell(
                        &status.text,
                        column.width,
                        self.tone_background(status.tone),
                        self.tone_foreground(status.tone),
                    ),
                    None => self.data_cell("", column.width),
                });
            }

            table = table.add_row(row);
        }

        table
    }

    fn tone_background(&self, tone: Tone) -> &str {
        match tone {
            Tone::Positive => &self.config.colors.mint,
            Tone::Neutral => &self.config.colors.pale,
            Tone::Warning => &self.config.colors.amber,
            Tone::Risk => &self.config.colors.rose,
        }
    }

    fn visual_from_spec(
        &self,
        spec: &BlockVisualSpec,
        kind: VisualKind,
        asset_base_dir: Option<&Path>,
    ) -> Visual {
        let mut source_path = PathBuf::from(&spec.path);
        if source_path.is_relative() {
            if let Some(base_dir) = asset_base_dir {
                source_path = base_dir.join(source_path);
            }
        }

        let mut visual = Visual::from_path(source_path).with_kind(kind);

        if let Some(alt_text) = &spec.alt_text {
            visual = visual.alt_text_text(alt_text);
        }
        if let Some(alignment) = &spec.alignment {
            visual = visual.with_alignment(match alignment {
                ParagraphAlignmentSpec::Left => ParagraphAlignment::Left,
                ParagraphAlignmentSpec::Center => ParagraphAlignment::Center,
                ParagraphAlignmentSpec::Right => ParagraphAlignment::Right,
                ParagraphAlignmentSpec::Justified => ParagraphAlignment::Justified,
            });
        }
        if let Some(width_twips) = spec.width_twips {
            visual = visual.width_twips(width_twips);
        }
        if let Some(height_twips) = spec.height_twips {
            visual = visual.height_twips(height_twips);
        }
        if let Some(max_width_twips) = spec.max_width_twips {
            visual = visual.max_width_twips(max_width_twips);
        }
        if let Some(max_height_twips) = spec.max_height_twips {
            visual = visual.max_height_twips(max_height_twips);
        }

        visual
    }

    fn tone_foreground(&self, tone: Tone) -> &str {
        match tone {
            Tone::Positive => &self.config.colors.green,
            Tone::Risk => &self.config.colors.red,
            Tone::Neutral | Tone::Warning => &self.config.colors.ink,
        }
    }
}

fn should_keep_next(block: &BlockSpec, next_block: Option<&BlockSpec>) -> bool {
    let Some(next_block) = next_block else {
        return false;
    };

    if matches!(next_block, BlockSpec::Spacer) {
        return false;
    }

    match block {
        BlockSpec::CoverTitle { .. }
        | BlockSpec::Title { .. }
        | BlockSpec::Subtitle { .. }
        | BlockSpec::Hero { .. }
        | BlockSpec::PageHeading { .. }
        | BlockSpec::Section { .. }
        | BlockSpec::Tagline { .. } => true,
        BlockSpec::Body { .. } | BlockSpec::Paragraph { .. } => matches!(
            next_block,
            BlockSpec::Image { .. }
                | BlockSpec::Logo { .. }
                | BlockSpec::Signature { .. }
                | BlockSpec::Chart { .. }
        ),
        _ => false,
    }
}

/// Builds a centered title paragraph using default config values.
pub fn title(text: &str) -> Paragraph {
    configured_studio().title(text)
}

/// Builds a base text run using default config values.
pub fn text_run(text: impl Into<String>) -> Run {
    configured_studio().text_run(text)
}

/// Builds a centered cover title paragraph using default config values.
pub fn cover_title(text: &str) -> Paragraph {
    configured_studio().cover_title(text)
}

/// Builds a centered subtitle paragraph using default config values.
pub fn subtitle(text: &str) -> Paragraph {
    configured_studio().subtitle(text)
}

/// Builds a centered hero paragraph using default config values.
pub fn hero(text: &str) -> Paragraph {
    configured_studio().hero(text)
}

/// Builds a section heading paragraph using default config values.
pub fn section(text: &str) -> Paragraph {
    configured_studio().section(text)
}

/// Builds a body paragraph using default config values.
pub fn body(text: &str) -> Paragraph {
    configured_studio().body(text)
}

/// Builds a bullet paragraph using default config values.
pub fn bullet(text: &str) -> Paragraph {
    configured_studio().bullet(text)
}

/// Builds a numbered list paragraph using default config values.
pub fn numbered(text: &str) -> Paragraph {
    configured_studio().numbered(text)
}

/// Builds a label-value paragraph using default config values.
pub fn label_value(label: &str, value: &str) -> Paragraph {
    configured_studio().label_value(label, value)
}

/// Builds a spacer paragraph using default config values.
pub fn spacer() -> Paragraph {
    configured_studio().spacer()
}

/// Builds a centered note paragraph using default config values.
pub fn centered_note(text: &str) -> Paragraph {
    configured_studio().centered_note(text)
}

/// Builds a page-heading paragraph using default config values.
pub fn page_heading(text: &str) -> Paragraph {
    configured_studio().page_heading(text)
}

/// Builds a centered accent tagline paragraph using default config values.
pub fn tagline(text: &str) -> Paragraph {
    configured_studio().tagline(text)
}

/// Builds a metric cell using default config values.
pub fn metric_cell(label: &str, value: &str, background: &str) -> TableCell {
    configured_studio().metric_cell(label, value, background)
}

/// Builds a header cell using default config values.
pub fn header_cell(text: &str, width: u32) -> TableCell {
    configured_studio().header_cell(text, width)
}

/// Builds a data cell using default config values.
pub fn data_cell(text: &str, width: u32) -> TableCell {
    configured_studio().data_cell(text, width)
}

/// Builds a status cell using default config values.
pub fn status_cell(text: &str, width: u32, background: &str, foreground: &str) -> TableCell {
    configured_studio().status_cell(text, width, background, foreground)
}

/// Returns standard grid borders using default config values.
pub fn grid_borders() -> TableBorders {
    configured_studio().grid_borders()
}

/// Returns card borders using default config values.
pub fn card_borders() -> TableBorders {
    configured_studio().card_borders()
}

/// Writes DOCX and optional PDF output using default config values.
pub fn save_with_pdf(document: &Document, docx_path: impl AsRef<Path>) -> crate::Result<()> {
    let docx_path = docx_path.as_ref();
    let studio = configured_studio();
    if uses_default_generated_folder(docx_path) {
        let stem = output_stem(docx_path)?;
        let _stats = studio.save_named(document, stem)?;
        Ok(())
    } else {
        studio.save_with_pdf(document, docx_path)
    }
}

/// Writes DOCX and optional PDF output using default config values and returns timing stats.
pub fn save_with_pdf_stats(
    document: &Document,
    docx_path: impl AsRef<Path>,
) -> crate::Result<OutputStats> {
    let docx_path = docx_path.as_ref();
    let studio = configured_studio();
    if uses_default_generated_folder(docx_path) {
        let stem = output_stem(docx_path)?;
        studio.save_named(document, stem)
    } else {
        studio.save_with_pdf_stats(document, docx_path)
    }
}

fn configured_studio() -> &'static Studio {
    static CONFIGURED: OnceLock<Studio> = OnceLock::new();
    CONFIGURED.get_or_init(|| Studio::from_default_file_or_default().unwrap_or_default())
}

fn uses_default_generated_folder(path: &Path) -> bool {
    path.parent() == Some(Path::new("generated"))
}

fn output_stem(path: &Path) -> crate::Result<&str> {
    path.file_stem()
        .and_then(|stem| stem.to_str())
        .ok_or_else(|| crate::DocxError::parse("invalid output file name"))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Rgb(u8, u8, u8);

#[derive(Clone, Copy)]
struct PdfRenderSettings {
    page_width: f32,
    page_height: f32,
    margin_x: f32,
    margin_top: f32,
    margin_bottom: f32,
    content_width: f32,
    default_text_size: f32,
    default_line_height: f32,
    line_height_multiplier: f32,
    baseline_factor: f32,
    text_width_bias_regular: f32,
    text_width_bias_bold: f32,
    table_cell_padding_x: f32,
    table_cell_padding_y: f32,
    table_after_spacing: f32,
    table_grid_stroke_width: f32,
    table_grid_stroke_color: Rgb,
    default_text_color: Rgb,
}

impl PdfRenderSettings {
    fn from_config(config: &RusdoxConfig) -> Self {
        let default_text_size =
            normalize_positive(config.pdf.default_text_size_pt, DEFAULT_TEXT_SIZE);
        let default_line_height =
            normalize_positive(config.pdf.default_line_height_pt, DEFAULT_LINE_HEIGHT);
        let page_width =
            normalize_positive(config.pdf.page_width_pt, DEFAULT_PAGE_WIDTH).max(MIN_CONTENT_WIDTH);
        let page_height = normalize_positive(config.pdf.page_height_pt, DEFAULT_PAGE_HEIGHT)
            .max(default_line_height);
        let margin_x = normalize_non_negative(config.pdf.margin_x_pt, DEFAULT_PAGE_MARGIN_X)
            .min(((page_width - MIN_CONTENT_WIDTH).max(0.0)) / 2.0);
        let margin_top = normalize_non_negative(config.pdf.margin_top_pt, DEFAULT_PAGE_MARGIN_TOP)
            .min((page_height - default_line_height).max(0.0));
        let margin_bottom =
            normalize_non_negative(config.pdf.margin_bottom_pt, DEFAULT_PAGE_MARGIN_BOTTOM)
                .min((page_height - margin_top - default_line_height).max(0.0));

        Self {
            page_width,
            page_height,
            margin_x,
            margin_top,
            margin_bottom,
            content_width: (page_width - (margin_x * 2.0)).max(MIN_CONTENT_WIDTH),
            default_text_size,
            default_line_height,
            line_height_multiplier: normalize_positive(
                config.pdf.line_height_multiplier,
                DEFAULT_LINE_HEIGHT_MULTIPLIER,
            ),
            baseline_factor: normalize_positive(
                config.pdf.baseline_factor,
                DEFAULT_BASELINE_FACTOR,
            ),
            text_width_bias_regular: normalize_positive(
                config.pdf.text_width_bias_regular,
                DEFAULT_TEXT_WIDTH_BIAS_REGULAR,
            ),
            text_width_bias_bold: normalize_positive(
                config.pdf.text_width_bias_bold,
                DEFAULT_TEXT_WIDTH_BIAS_BOLD,
            ),
            table_cell_padding_x: normalize_non_negative(
                config.table.pdf_cell_padding_x_pt,
                DEFAULT_TABLE_ROW_PADDING_X,
            ),
            table_cell_padding_y: normalize_non_negative(
                config.table.pdf_cell_padding_y_pt,
                DEFAULT_TABLE_ROW_PADDING_Y,
            ),
            table_after_spacing: normalize_non_negative(
                config.table.pdf_after_spacing_pt,
                DEFAULT_TABLE_AFTER_SPACING,
            ),
            table_grid_stroke_width: normalize_non_negative(
                config.table.pdf_grid_stroke_width_pt,
                DEFAULT_TABLE_GRID_STROKE_WIDTH,
            ),
            table_grid_stroke_color: parse_hex_color(&config.colors.table_border)
                .unwrap_or(Rgb(203, 213, 225)),
            default_text_color: parse_hex_color(&config.colors.ink).unwrap_or(Rgb(15, 23, 42)),
        }
    }

    fn effective_line_height(self, size: f32) -> f32 {
        self.default_line_height
            .max(size * self.line_height_multiplier)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum PdfFont {
    Regular,
    Bold,
    Oblique,
    BoldOblique,
}

impl PdfFont {
    fn width_bias(self, settings: PdfRenderSettings) -> f32 {
        match self {
            Self::Regular | Self::Oblique => settings.text_width_bias_regular,
            Self::Bold | Self::BoldOblique => settings.text_width_bias_bold,
        }
    }

    fn fontdb_weight(self) -> fontdb::Weight {
        match self {
            Self::Bold | Self::BoldOblique => fontdb::Weight::BOLD,
            Self::Regular | Self::Oblique => fontdb::Weight::NORMAL,
        }
    }

    fn fontdb_style(self) -> fontdb::Style {
        match self {
            Self::Oblique | Self::BoldOblique => fontdb::Style::Italic,
            Self::Regular | Self::Bold => fontdb::Style::Normal,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct PdfFontRequest {
    family: String,
    font: PdfFont,
}

impl PdfFontRequest {
    fn new(family: impl Into<String>, font: PdfFont) -> Self {
        Self {
            family: family.into(),
            font,
        }
    }
}

#[derive(Clone)]
struct RequestedTextStyle {
    font_request: PdfFontRequest,
    size: f32,
    color: Rgb,
}

#[derive(Clone, Copy)]
struct TextPaintStyle {
    font_id: usize,
    size: f32,
    color: Rgb,
}

#[derive(Clone)]
struct TextSpan {
    text: String,
    glyphs: Vec<u16>,
    style: TextPaintStyle,
    width: f32,
}

#[derive(Clone, Copy)]
struct UsedCharacter {
    cid: u16,
    glyph_id: u16,
    pdf_width: f32,
}

struct ShapedToken {
    spans: Vec<TextSpan>,
    width: f32,
}

#[derive(Clone)]
struct LineLayout {
    spans: Vec<TextSpan>,
    width: f32,
    line_height: f32,
    alignment: ParagraphAlignment,
}

#[derive(Default)]
struct Page {
    ops: Vec<DrawOp>,
}

enum DrawOp {
    TextLine {
        x: f32,
        y_top: f32,
        line: LineLayout,
        max_width: f32,
    },
    Rect {
        x: f32,
        y_top: f32,
        width: f32,
        height: f32,
        fill: Option<Rgb>,
        stroke: Option<(Rgb, f32)>,
    },
    Image {
        x: f32,
        y_top: f32,
        width: f32,
        height: f32,
        image_id: usize,
    },
}

struct PdfDocumentLayout {
    pages: Vec<Page>,
    images: Vec<PdfImageAsset>,
}

struct PdfImageAsset {
    resource_name: String,
    width_px: u32,
    height_px: u32,
    encoded_rgb: Vec<u8>,
    encoded_alpha: Option<Vec<u8>>,
}

struct PdfLayout {
    pages: Vec<Page>,
    images: Vec<PdfImageAsset>,
    cursor_y: f32,
    settings: PdfRenderSettings,
}

impl PdfLayout {
    fn new(settings: PdfRenderSettings) -> Self {
        Self {
            pages: vec![Page::default()],
            images: Vec::new(),
            cursor_y: settings.margin_top,
            settings,
        }
    }

    fn current_page_mut(&mut self) -> &mut Page {
        if self.pages.is_empty() {
            self.pages.push(Page::default());
            self.cursor_y = self.settings.margin_top;
        }
        let index = self.pages.len() - 1;
        &mut self.pages[index]
    }

    fn push_page(&mut self) {
        self.pages.push(Page::default());
        self.cursor_y = self.settings.margin_top;
    }

    fn ensure_space(&mut self, height: f32) {
        if self.cursor_y + height > self.settings.page_height - self.settings.margin_bottom {
            self.push_page();
        }
    }

    fn push_op(&mut self, op: DrawOp) {
        self.current_page_mut().ops.push(op);
    }
}

const PDF_FONT_SYSTEM_INFO: SystemInfo<'static> = SystemInfo {
    registry: Str(b"Adobe"),
    ordering: Str(b"Identity"),
    supplement: 0,
};

const PDF_SANS_FALLBACKS: &[&str] = &[
    "Noto Sans",
    "Liberation Sans",
    "DejaVu Sans",
    "Arial Unicode MS",
    "Droid Sans Fallback",
];

const PDF_SERIF_FALLBACKS: &[&str] = &[
    "Noto Serif",
    "Liberation Serif",
    "DejaVu Serif",
    "Times New Roman",
];

const PDF_MONO_FALLBACKS: &[&str] = &[
    "Noto Sans Mono",
    "Liberation Mono",
    "DejaVu Sans Mono",
    "Courier New",
];

const PDF_CJK_FALLBACKS: &[&str] = &[
    "Droid Sans Fallback",
    "Noto Sans CJK SC",
    "WenQuanYi Zen Hei",
    "Noto Serif CJK SC",
];

const PDF_ARABIC_FALLBACKS: &[&str] = &["Noto Sans Arabic", "Noto Naskh Arabic", "Amiri"];

const PDF_HEBREW_FALLBACKS: &[&str] = &["Noto Sans Hebrew", "Noto Serif Hebrew"];

const PDF_DEVANAGARI_FALLBACKS: &[&str] = &["Noto Sans Devanagari", "Noto Serif Devanagari"];

struct PdfFontSystem {
    db: &'static FontDatabase,
    settings: PdfRenderSettings,
    default_family: String,
    face_cache: HashMap<(fontdb::ID, PdfFont), usize>,
    family_cache: HashMap<PdfFontRequest, Option<usize>>,
    request_cache: HashMap<PdfFontRequest, Vec<usize>>,
    faces: Vec<ResolvedPdfFace>,
}

impl PdfFontSystem {
    fn new(config: &RusdoxConfig, settings: PdfRenderSettings) -> Self {
        let default_family = config.typography.font_family.trim();
        Self {
            db: system_font_db(),
            settings,
            default_family: if default_family.is_empty() {
                "Arial".to_string()
            } else {
                default_family.to_string()
            },
            face_cache: HashMap::new(),
            family_cache: HashMap::new(),
            request_cache: HashMap::new(),
            faces: Vec::new(),
        }
    }

    fn default_family(&self) -> &str {
        &self.default_family
    }

    fn used_face_ids(&self) -> Vec<usize> {
        self.faces
            .iter()
            .enumerate()
            .filter_map(|(index, face)| face.is_used().then_some(index))
            .collect()
    }

    fn face(&self, font_id: usize) -> &ResolvedPdfFace {
        &self.faces[font_id]
    }

    fn shape_text(&mut self, style: &RequestedTextStyle, text: &str) -> crate::Result<ShapedToken> {
        let normalized = normalize_pdf_text(text);
        if normalized.is_empty() {
            return Ok(ShapedToken {
                spans: Vec::new(),
                width: 0.0,
            });
        }

        let base_faces = self.resolve_base_faces(&style.font_request)?;
        let mut spans = Vec::new();
        let mut total_width = 0.0;
        let mut current_font_id = None;
        let mut current_text = String::new();
        let mut current_glyphs = Vec::new();
        let mut current_width = 0.0;

        for ch in normalized.chars() {
            let font_id = self.resolve_face_for_char(&style.font_request, &base_faces, ch);
            let used = self.faces[font_id].ensure_char(ch);
            let advance = self.faces[font_id].advance_points(used, style.size);

            if current_font_id != Some(font_id) && !current_glyphs.is_empty() {
                spans.push(TextSpan {
                    text: std::mem::take(&mut current_text),
                    glyphs: std::mem::take(&mut current_glyphs),
                    style: TextPaintStyle {
                        font_id: current_font_id.expect("font id"),
                        size: style.size,
                        color: style.color,
                    },
                    width: current_width,
                });
                current_width = 0.0;
            }

            current_font_id = Some(font_id);
            current_text.push(ch);
            current_glyphs.push(used.cid);
            current_width += advance;
            total_width += advance;
        }

        if let Some(font_id) = current_font_id {
            spans.push(TextSpan {
                text: current_text,
                glyphs: current_glyphs,
                style: TextPaintStyle {
                    font_id,
                    size: style.size,
                    color: style.color,
                },
                width: current_width,
            });
        }

        Ok(ShapedToken {
            spans,
            width: total_width,
        })
    }

    fn resolve_base_faces(&mut self, request: &PdfFontRequest) -> crate::Result<Vec<usize>> {
        if let Some(cached) = self.request_cache.get(request) {
            return Ok(cached.clone());
        }

        let mut faces = Vec::new();
        for family in base_font_family_candidates(&request.family) {
            if let Some(font_id) = self.resolve_face_by_family(&family, request.font) {
                if !faces.contains(&font_id) {
                    faces.push(font_id);
                }
            }
        }

        if faces.is_empty() {
            return Err(crate::DocxError::parse(format!(
                "could not find an embeddable TrueType font for '{}'",
                request.family
            )));
        }

        self.request_cache.insert(request.clone(), faces.clone());
        Ok(faces)
    }

    fn resolve_face_for_char(
        &mut self,
        request: &PdfFontRequest,
        base_faces: &[usize],
        ch: char,
    ) -> usize {
        for &font_id in base_faces {
            if self.faces[font_id].supports_char(ch) {
                return font_id;
            }
        }

        for &family in script_specific_fallback_families(ch) {
            if let Some(font_id) = self.resolve_face_by_family(family, request.font) {
                if self.faces[font_id].supports_char(ch) {
                    return font_id;
                }
            }
        }

        base_faces[0]
    }

    fn resolve_face_by_family(&mut self, family: &str, font: PdfFont) -> Option<usize> {
        let request = PdfFontRequest::new(family, font);
        if let Some(cached) = self.family_cache.get(&request) {
            return *cached;
        }

        let families = [FontFamily::Name(family)];
        let query = FontQuery {
            families: &families,
            weight: font.fontdb_weight(),
            stretch: fontdb::Stretch::Normal,
            style: font.fontdb_style(),
        };

        let resolved = self
            .db
            .query(&query)
            .and_then(|db_id| self.load_face(db_id, font));
        self.family_cache.insert(request, resolved);
        resolved
    }

    fn load_face(&mut self, db_id: fontdb::ID, font: PdfFont) -> Option<usize> {
        if let Some(&cached) = self.face_cache.get(&(db_id, font)) {
            return Some(cached);
        }

        let resource_index = self.faces.len();
        let face = ResolvedPdfFace::load(self.db, db_id, font, self.settings, resource_index)?;
        self.faces.push(face);
        self.face_cache.insert((db_id, font), resource_index);
        Some(resource_index)
    }
}

struct ResolvedPdfFace {
    family_name: String,
    base_font_name: String,
    type0_font_name: String,
    cmap_name: String,
    resource_name: String,
    font_bytes: Vec<u8>,
    face_index: u32,
    units_per_em: f32,
    width_bias: f32,
    default_width: f32,
    missing_width: f32,
    missing_width_units: u16,
    avg_width: f32,
    max_width: f32,
    bbox: Rect,
    ascent: f32,
    descent: f32,
    cap_height: f32,
    italic_angle: f32,
    font_flags: FontFlags,
    stem_v: f32,
    used_chars: BTreeMap<char, UsedCharacter>,
    next_cid: u16,
}

impl ResolvedPdfFace {
    fn load(
        db: &FontDatabase,
        db_id: fontdb::ID,
        font: PdfFont,
        settings: PdfRenderSettings,
        resource_index: usize,
    ) -> Option<Self> {
        let info = db.face(db_id)?;
        let family_name = info
            .families
            .first()
            .map(|(family, _)| family.clone())
            .unwrap_or_else(|| "Rusdox Sans".to_string());
        let post_script_name = if info.post_script_name.is_empty() {
            family_name.clone()
        } else {
            info.post_script_name.clone()
        };

        db.with_face_data(db_id, |font_data, face_index| {
            if !supports_embedded_truetype(font_data) {
                return None;
            }

            let face = ttf_parser::Face::parse(font_data, face_index).ok()?;
            if matches!(
                face.permissions(),
                Some(ttf_parser::Permissions::Restricted)
            ) {
                return None;
            }

            let units_per_em = f32::from(face.units_per_em());
            let missing_width_units = face
                .glyph_hor_advance(ttf_parser::GlyphId(0))
                .or_else(|| {
                    face.glyph_index(' ')
                        .and_then(|glyph_id| face.glyph_hor_advance(glyph_id))
                })
                .unwrap_or(face.units_per_em() / 2);
            let width_bias = font.width_bias(settings);
            let default_width =
                scale_font_units(f32::from(missing_width_units), units_per_em) * width_bias;
            let bbox = face.global_bounding_box();
            let bbox = Rect::new(
                scale_font_units(f32::from(bbox.x_min), units_per_em),
                scale_font_units(f32::from(bbox.y_min), units_per_em),
                scale_font_units(f32::from(bbox.x_max), units_per_em),
                scale_font_units(f32::from(bbox.y_max), units_per_em),
            );
            let cap_height = scale_font_units(
                f32::from(face.capital_height().unwrap_or(face.ascender())),
                units_per_em,
            );
            let ascent = scale_font_units(f32::from(face.ascender()), units_per_em);
            let descent = scale_font_units(f32::from(face.descender()), units_per_em);
            let bbox_width = (bbox.x2 - bbox.x1).abs().max(default_width);

            let mut font_flags = FontFlags::SYMBOLIC;
            if face.is_monospaced() {
                font_flags |= FontFlags::FIXED_PITCH;
            }
            if face.italic_angle() != 0.0 || face.is_italic() {
                font_flags |= FontFlags::ITALIC;
            }
            if face.is_bold() {
                font_flags |= FontFlags::FORCE_BOLD;
            }

            Some(Self {
                family_name,
                base_font_name: sanitize_pdf_name_component(&post_script_name),
                type0_font_name: format!(
                    "{}-Identity-H",
                    sanitize_pdf_name_component(&post_script_name)
                ),
                cmap_name: format!("{}-UTF16", sanitize_pdf_name_component(&post_script_name)),
                resource_name: format!("F{}", resource_index + 1),
                font_bytes: font_data.to_vec(),
                face_index,
                units_per_em,
                width_bias,
                default_width,
                missing_width: default_width,
                missing_width_units,
                avg_width: default_width,
                max_width: bbox_width,
                bbox,
                ascent,
                descent,
                cap_height,
                italic_angle: face.italic_angle(),
                font_flags,
                stem_v: if face.is_bold() { 120.0 } else { 80.0 },
                used_chars: BTreeMap::new(),
                next_cid: 1,
            })
        })?
    }

    fn is_used(&self) -> bool {
        !self.used_chars.is_empty()
    }

    fn supports_char(&self, ch: char) -> bool {
        self.lookup_glyph(ch).is_some()
    }

    fn ensure_char(&mut self, ch: char) -> UsedCharacter {
        if let Some(&used) = self.used_chars.get(&ch) {
            return used;
        }

        let (glyph_id, advance_units) = self
            .lookup_glyph(ch)
            .unwrap_or((0, self.missing_width_units));
        let used = UsedCharacter {
            cid: self.next_cid,
            glyph_id,
            pdf_width: scale_font_units(f32::from(advance_units), self.units_per_em)
                * self.width_bias,
        };
        self.next_cid = self.next_cid.saturating_add(1);
        self.used_chars.insert(ch, used);
        self.avg_width = self.avg_width.max(used.pdf_width);
        self.max_width = self.max_width.max(used.pdf_width);
        used
    }

    fn advance_points(&self, used: UsedCharacter, size: f32) -> f32 {
        (used.pdf_width * size) / 1000.0
    }

    fn cid_to_gid_map_data(&self) -> Vec<u8> {
        let max_cid = self
            .used_chars
            .values()
            .map(|used| used.cid)
            .max()
            .unwrap_or(0);
        let mut data = vec![0; (usize::from(max_cid) + 1) * 2];
        for used in self.used_chars.values() {
            let index = usize::from(used.cid) * 2;
            data[index] = (used.glyph_id >> 8) as u8;
            data[index + 1] = (used.glyph_id & 0x00FF) as u8;
        }
        data
    }

    fn ordered_used_chars(&self) -> Vec<(char, UsedCharacter)> {
        let mut characters = self
            .used_chars
            .iter()
            .map(|(&ch, &used)| (ch, used))
            .collect::<Vec<_>>();
        characters.sort_by_key(|(_, used)| used.cid);
        characters
    }

    fn lookup_glyph(&self, ch: char) -> Option<(u16, u16)> {
        let face = ttf_parser::Face::parse(&self.font_bytes, self.face_index).ok()?;
        let glyph_id = face.glyph_index(ch)?;
        let advance_units = face
            .glyph_hor_advance(glyph_id)
            .unwrap_or(self.missing_width_units);
        Some((glyph_id.0, advance_units))
    }
}

#[derive(Clone, Copy)]
struct PdfFaceObjectIds {
    type0_id: Ref,
    cid_font_id: Ref,
    font_descriptor_id: Ref,
    font_file_id: Ref,
    to_unicode_id: Ref,
    cid_to_gid_map_id: Ref,
}

#[derive(Clone, Copy)]
struct PdfImageObjectIds {
    image_id: Ref,
    alpha_id: Option<Ref>,
}

fn render_pdf(document: &Document, pdf_path: &Path, config: &RusdoxConfig) -> crate::Result<()> {
    let settings = PdfRenderSettings::from_config(config);
    let mut font_system = PdfFontSystem::new(config, settings);
    let layout = layout_document(document, settings, &mut font_system)?;
    let catalog_id = Ref::new(1);
    let page_tree_id = Ref::new(2);
    let mut next_id = 3;
    let mut page_ids = Vec::with_capacity(layout.pages.len());
    let mut content_ids = Vec::with_capacity(layout.pages.len());

    for _ in &layout.pages {
        page_ids.push(Ref::new(next_id));
        next_id += 1;
        content_ids.push(Ref::new(next_id));
        next_id += 1;
    }

    let used_face_ids = font_system.used_face_ids();
    let mut font_object_ids = HashMap::new();
    for &face_id in &used_face_ids {
        let object_ids = PdfFaceObjectIds {
            type0_id: Ref::new(next_id),
            cid_font_id: Ref::new(next_id + 1),
            font_descriptor_id: Ref::new(next_id + 2),
            font_file_id: Ref::new(next_id + 3),
            to_unicode_id: Ref::new(next_id + 4),
            cid_to_gid_map_id: Ref::new(next_id + 5),
        };
        next_id += 6;
        font_object_ids.insert(face_id, object_ids);
    }

    let mut image_object_ids = Vec::with_capacity(layout.images.len());
    for image in &layout.images {
        let image_id = Ref::new(next_id);
        next_id += 1;
        let alpha_id = image.encoded_alpha.as_ref().map(|_| {
            let id = Ref::new(next_id);
            next_id += 1;
            id
        });
        image_object_ids.push(PdfImageObjectIds { image_id, alpha_id });
    }

    let estimated_capacity = layout
        .pages
        .iter()
        .map(|page| page.ops.len() * 96)
        .sum::<usize>()
        .max(64 * 1024);
    let mut pdf = Pdf::with_capacity(estimated_capacity);
    pdf.catalog(catalog_id).pages(page_tree_id);
    pdf.pages(page_tree_id)
        .kids(page_ids.iter().copied())
        .count(i32::try_from(page_ids.len()).unwrap_or(i32::MAX));

    for &face_id in &used_face_ids {
        let face = font_system.face(face_id);
        let object_ids = font_object_ids[&face_id];
        let cid_to_gid_map = face.cid_to_gid_map_data();
        let unicode_cmap_bytes = {
            let mut cmap = UnicodeCmap::new(Name(face.cmap_name.as_bytes()), PDF_FONT_SYSTEM_INFO);
            for (ch, used) in face.ordered_used_chars() {
                cmap.pair(used.cid, ch);
            }
            cmap.finish().into_vec()
        };

        pdf.stream(object_ids.font_file_id, &face.font_bytes);
        pdf.stream(object_ids.cid_to_gid_map_id, &cid_to_gid_map);
        pdf.cmap(object_ids.to_unicode_id, &unicode_cmap_bytes)
            .name(Name(face.cmap_name.as_bytes()))
            .system_info(PDF_FONT_SYSTEM_INFO);

        pdf.font_descriptor(object_ids.font_descriptor_id)
            .name(Name(face.base_font_name.as_bytes()))
            .family(Str(face.family_name.as_bytes()))
            .flags(FontFlags::from_bits_retain(face.font_flags.bits()))
            .bbox(face.bbox)
            .italic_angle(face.italic_angle)
            .ascent(face.ascent)
            .descent(face.descent)
            .cap_height(face.cap_height)
            .stem_v(face.stem_v)
            .avg_width(face.avg_width)
            .max_width(face.max_width)
            .missing_width(face.missing_width)
            .font_file2(object_ids.font_file_id);

        {
            let mut cid_font = pdf.cid_font(object_ids.cid_font_id);
            cid_font
                .subtype(CidFontType::Type2)
                .base_font(Name(face.base_font_name.as_bytes()))
                .system_info(PDF_FONT_SYSTEM_INFO)
                .font_descriptor(object_ids.font_descriptor_id)
                .default_width(face.default_width)
                .cid_to_gid_map_stream(object_ids.cid_to_gid_map_id);
            let mut widths = cid_font.widths();
            for (_, used) in face.ordered_used_chars() {
                widths.consecutive(used.cid, [used.pdf_width]);
            }
        }

        pdf.type0_font(object_ids.type0_id)
            .base_font(Name(face.type0_font_name.as_bytes()))
            .encoding_predefined(Name(b"Identity-H"))
            .descendant_font(object_ids.cid_font_id)
            .to_unicode(object_ids.to_unicode_id);
    }

    for (index, image) in layout.images.iter().enumerate() {
        let object_ids = image_object_ids[index];
        let mut image_stream = pdf.image_xobject(object_ids.image_id, &image.encoded_rgb);
        image_stream.filter(Filter::FlateDecode);
        image_stream.width(image.width_px as i32);
        image_stream.height(image.height_px as i32);
        image_stream.color_space().device_rgb();
        image_stream.bits_per_component(8);
        if let Some(alpha_id) = object_ids.alpha_id {
            image_stream.s_mask(alpha_id);
        }
        image_stream.finish();

        if let (Some(alpha_id), Some(alpha)) = (object_ids.alpha_id, image.encoded_alpha.as_ref()) {
            let mut alpha_stream = pdf.image_xobject(alpha_id, alpha);
            alpha_stream.filter(Filter::FlateDecode);
            alpha_stream.width(image.width_px as i32);
            alpha_stream.height(image.height_px as i32);
            alpha_stream.color_space().device_gray();
            alpha_stream.bits_per_component(8);
            alpha_stream.finish();
        }
    }

    for ((page, page_id), content_id) in layout.pages.iter().zip(&page_ids).zip(&content_ids) {
        let content = render_page_content(page, settings, &font_system, &layout.images);
        pdf.stream(*content_id, &content);

        let mut page_writer = pdf.page(*page_id);
        page_writer
            .parent(page_tree_id)
            .media_box(Rect::new(
                0.0,
                0.0,
                settings.page_width,
                settings.page_height,
            ))
            .contents(*content_id);

        let mut resources = page_writer.resources();
        {
            let mut fonts = resources.fonts();
            for &face_id in &used_face_ids {
                let face = font_system.face(face_id);
                fonts.pair(
                    Name(face.resource_name.as_bytes()),
                    font_object_ids[&face_id].type0_id,
                );
            }
        }
        let mut x_objects = resources.x_objects();
        for (image, object_ids) in layout.images.iter().zip(&image_object_ids) {
            x_objects.pair(Name(image.resource_name.as_bytes()), object_ids.image_id);
        }
    }

    fs::write(pdf_path, pdf.finish())?;
    Ok(())
}

fn layout_document(
    document: &Document,
    settings: PdfRenderSettings,
    font_system: &mut PdfFontSystem,
) -> crate::Result<PdfDocumentLayout> {
    let mut layout = PdfLayout::new(settings);
    let blocks = document.blocks().collect::<Vec<_>>();
    let styles = document.styles();

    for (index, block) in blocks.iter().copied().enumerate() {
        maybe_push_page_for_keep_next_group(&mut layout, styles, &blocks, index, font_system)?;
        match block {
            DocumentBlockRef::Paragraph(paragraph) => {
                layout_paragraph_block(&mut layout, styles, paragraph, font_system)?
            }
            DocumentBlockRef::Table(table) => {
                layout_table_block(&mut layout, styles, table, font_system)?
            }
            DocumentBlockRef::Visual(visual) => layout_visual_block(&mut layout, visual)?,
        }
    }

    Ok(PdfDocumentLayout {
        pages: layout.pages,
        images: layout.images,
    })
}

fn maybe_push_page_for_keep_next_group(
    layout: &mut PdfLayout,
    styles: &Stylesheet,
    blocks: &[DocumentBlockRef<'_>],
    start_index: usize,
    font_system: &mut PdfFontSystem,
) -> crate::Result<()> {
    let DocumentBlockRef::Paragraph(paragraph) = blocks[start_index] else {
        return Ok(());
    };
    if !paragraph.has_keep_next() {
        return Ok(());
    }

    let usable_height =
        layout.settings.page_height - layout.settings.margin_top - layout.settings.margin_bottom;
    let mut group_height = 0.0;
    let mut index = start_index;

    loop {
        let Some(block_height) =
            estimated_keep_next_block_height(layout, styles, blocks[index], font_system)?
        else {
            return Ok(());
        };
        group_height += block_height;

        match blocks[index] {
            DocumentBlockRef::Paragraph(paragraph) if paragraph.has_keep_next() => {
                index += 1;
                if index >= blocks.len() {
                    break;
                }
            }
            _ => break,
        }
    }

    if group_height > usable_height || layout.cursor_y <= layout.settings.margin_top {
        return Ok(());
    }

    let start_y = if paragraph.has_page_break_before() {
        layout.settings.margin_top
    } else {
        layout.cursor_y
    };
    let remaining_height = layout.settings.page_height - layout.settings.margin_bottom - start_y;
    if group_height > remaining_height {
        layout.push_page();
    }

    Ok(())
}

fn estimated_keep_next_block_height(
    layout: &PdfLayout,
    styles: &Stylesheet,
    block: DocumentBlockRef<'_>,
    font_system: &mut PdfFontSystem,
) -> crate::Result<Option<f32>> {
    match block {
        DocumentBlockRef::Paragraph(paragraph) => {
            let lines = layout_paragraph_lines(
                styles,
                paragraph,
                layout.settings.content_width,
                layout.settings,
                font_system,
            )?;
            let content_height = if lines.is_empty() {
                layout.settings.default_line_height
            } else {
                lines.iter().map(|line| line.line_height).sum()
            };
            let resolved = styles.resolve_paragraph(paragraph)?;
            Ok(Some(
                twips_to_points(resolved.spacing_before.unwrap_or(0))
                    + content_height
                    + twips_to_points(resolved.spacing_after.unwrap_or(0)),
            ))
        }
        DocumentBlockRef::Visual(visual) => {
            let content_width_twips = points_to_twips(layout.settings.content_width);
            let content_height_twips = points_to_twips(
                layout.settings.page_height
                    - layout.settings.margin_top
                    - layout.settings.margin_bottom,
            );
            let (_, height_twips) =
                visual.resolved_dimensions_twips(content_width_twips, content_height_twips)?;
            Ok(Some(twips_to_points(height_twips)))
        }
        DocumentBlockRef::Table(_) => Ok(None),
    }
}

fn layout_paragraph_block(
    layout: &mut PdfLayout,
    styles: &Stylesheet,
    paragraph: &Paragraph,
    font_system: &mut PdfFontSystem,
) -> crate::Result<()> {
    let resolved = styles.resolve_paragraph(paragraph)?;
    if resolved.page_break_before && layout.cursor_y > layout.settings.margin_top {
        layout.push_page();
    }

    layout.cursor_y += twips_to_points(resolved.spacing_before.unwrap_or(0));
    let lines = layout_paragraph_lines(
        styles,
        paragraph,
        layout.settings.content_width,
        layout.settings,
        font_system,
    )?;

    if lines.is_empty() {
        layout.ensure_space(layout.settings.default_line_height);
        layout.cursor_y += layout.settings.default_line_height;
    } else {
        for line in lines {
            let line_height = line.line_height;
            layout.ensure_space(line_height);
            let y_top = layout.cursor_y;
            layout.push_op(DrawOp::TextLine {
                x: layout.settings.margin_x,
                y_top,
                line,
                max_width: layout.settings.content_width,
            });
            layout.cursor_y += line_height;
        }
    }

    layout.cursor_y += twips_to_points(resolved.spacing_after.unwrap_or(0));
    Ok(())
}

fn layout_table_block(
    layout: &mut PdfLayout,
    styles: &Stylesheet,
    table: &Table,
    font_system: &mut PdfFontSystem,
) -> crate::Result<()> {
    let resolved_table = styles.resolve_table(table)?;
    let total_width = resolved_table
        .width
        .map(twips_to_points)
        .unwrap_or(layout.settings.content_width)
        .min(layout.settings.content_width);
    let column_widths = resolve_table_column_widths(table, total_width);
    let row_layouts = table
        .rows()
        .map(|row| layout_row(row, styles, &column_widths, layout.settings, font_system))
        .collect::<crate::Result<Vec<_>>>()?;
    let repeated_headers = row_layouts
        .iter()
        .take_while(|row| row.repeat_as_header)
        .cloned()
        .collect::<Vec<_>>();

    for row_layout in row_layouts {
        place_table_row(layout, row_layout, &repeated_headers);
    }

    layout.cursor_y += layout.settings.table_after_spacing;
    Ok(())
}

fn layout_visual_block(layout: &mut PdfLayout, visual: &Visual) -> crate::Result<()> {
    let content_width_twips = points_to_twips(layout.settings.content_width);
    let content_height_twips = points_to_twips(
        layout.settings.page_height - layout.settings.margin_top - layout.settings.margin_bottom,
    );
    let (width_twips, height_twips) =
        visual.resolved_dimensions_twips(content_width_twips, content_height_twips)?;
    let width = twips_to_points(width_twips);
    let height = twips_to_points(height_twips);
    let image_id = prepare_pdf_image_asset(layout, visual, width_twips, height_twips)?;

    layout.ensure_space(height);
    let x = match visual.alignment() {
        ParagraphAlignment::Center => {
            layout.settings.margin_x + ((layout.settings.content_width - width).max(0.0) / 2.0)
        }
        ParagraphAlignment::Right => {
            layout.settings.margin_x + (layout.settings.content_width - width).max(0.0)
        }
        _ => layout.settings.margin_x,
    };
    let y_top = layout.cursor_y;
    layout.push_op(DrawOp::Image {
        x,
        y_top,
        width,
        height,
        image_id,
    });
    layout.cursor_y += height;
    Ok(())
}

fn prepare_pdf_image_asset(
    layout: &mut PdfLayout,
    visual: &Visual,
    width_twips: u32,
    height_twips: u32,
) -> crate::Result<usize> {
    let raster = visual.pdf_raster(width_twips, height_twips)?;
    let level = CompressionLevel::DefaultLevel as u8;
    let mut rgb = Vec::with_capacity((raster.width_px * raster.height_px * 3) as usize);
    let mut alpha = Vec::with_capacity((raster.width_px * raster.height_px) as usize);
    let mut has_alpha = false;

    for pixel in raster.rgba.chunks_exact(4) {
        rgb.extend_from_slice(&pixel[..3]);
        alpha.push(pixel[3]);
        has_alpha |= pixel[3] < 255;
    }

    let image_id = layout.images.len();
    layout.images.push(PdfImageAsset {
        resource_name: format!("Im{}", image_id + 1),
        width_px: raster.width_px,
        height_px: raster.height_px,
        encoded_rgb: compress_to_vec_zlib(&rgb, level),
        encoded_alpha: has_alpha.then(|| compress_to_vec_zlib(&alpha, level)),
    });
    Ok(image_id)
}

#[derive(Clone)]
struct RowLayout {
    cells: Vec<CellLayout>,
    height: f32,
    allow_split_across_pages: bool,
    repeat_as_header: bool,
}

#[derive(Clone)]
struct CellLayout {
    x_offset: f32,
    width: f32,
    background: Option<Rgb>,
    lines: Vec<CellLine>,
}

#[derive(Clone)]
struct CellLine {
    y_offset: f32,
    layout: LineLayout,
}

struct ResolvedTableCell<'a> {
    cell: &'a TableCell,
    x_offset: f32,
    width: f32,
}

fn layout_row(
    row: &TableRow,
    styles: &Stylesheet,
    column_widths: &[f32],
    settings: PdfRenderSettings,
    font_system: &mut PdfFontSystem,
) -> crate::Result<RowLayout> {
    let mut cells = Vec::new();
    let mut row_height: f32 = 0.0;

    for resolved_cell in resolve_row_cells(row, column_widths) {
        let width = resolved_cell.width;
        let content_width = (width - (settings.table_cell_padding_x * 2.0)).max(MIN_CONTENT_WIDTH);
        let mut lines = Vec::new();
        let mut y_offset = settings.table_cell_padding_y;

        for paragraph in resolved_cell.cell.paragraphs() {
            let resolved_paragraph = styles.resolve_paragraph(paragraph)?;
            y_offset += twips_to_points(resolved_paragraph.spacing_before.unwrap_or(0));
            let paragraph_lines =
                layout_paragraph_lines(styles, paragraph, content_width, settings, font_system)?;
            if paragraph_lines.is_empty() {
                lines.push(CellLine {
                    y_offset,
                    layout: blank_line_layout(styles, paragraph, settings)?,
                });
                y_offset += settings.default_line_height;
            } else {
                for line in paragraph_lines {
                    let line_height = line.line_height;
                    lines.push(CellLine {
                        y_offset,
                        layout: line,
                    });
                    y_offset += line_height;
                }
            }
            y_offset += twips_to_points(resolved_paragraph.spacing_after.unwrap_or(0));
        }

        let height = cell_height(&lines, settings);
        row_height = row_height.max(height);
        cells.push(CellLayout {
            x_offset: resolved_cell.x_offset,
            width,
            background: resolved_cell
                .cell
                .properties()
                .background_color
                .as_deref()
                .and_then(parse_hex_color),
            lines,
        });
    }

    Ok(RowLayout {
        cells,
        height: row_height.max(MIN_CONTENT_WIDTH),
        allow_split_across_pages: row.properties().allow_split_across_pages,
        repeat_as_header: row.properties().repeat_as_header,
    })
}

fn blank_line_layout(
    styles: &Stylesheet,
    paragraph: &Paragraph,
    settings: PdfRenderSettings,
) -> crate::Result<LineLayout> {
    Ok(LineLayout {
        spans: Vec::new(),
        width: 0.0,
        line_height: settings.default_line_height,
        alignment: styles
            .resolve_paragraph(paragraph)?
            .alignment
            .unwrap_or(ParagraphAlignment::Left),
    })
}

fn resolve_table_column_widths(table: &Table, total_width: f32) -> Vec<f32> {
    let column_count = table
        .rows()
        .map(|row| {
            row.cells()
                .map(|cell| cell.properties().grid_span.unwrap_or(1).max(1) as usize)
                .sum::<usize>()
        })
        .max()
        .unwrap_or(1)
        .max(1);
    let mut column_widths = vec![0.0_f32; column_count];

    for row in table.rows() {
        let mut column_index = 0usize;
        for cell in row.cells() {
            let remaining_columns = column_count.saturating_sub(column_index);
            if remaining_columns == 0 {
                break;
            }
            let span = normalized_grid_span(cell.properties().grid_span, remaining_columns);
            if let Some(width) = cell.properties().width.map(twips_to_points) {
                let per_column = width / span as f32;
                for column_width in &mut column_widths[column_index..column_index + span] {
                    *column_width = (*column_width).max(per_column);
                }
            }
            column_index += span;
        }
    }

    let mut assigned_width = column_widths.iter().sum::<f32>();
    let zero_columns = column_widths.iter().filter(|&&width| width <= 0.0).count();
    if zero_columns > 0 {
        let fallback = if total_width > assigned_width {
            (total_width - assigned_width) / zero_columns as f32
        } else {
            total_width / column_count as f32
        };
        for column_width in &mut column_widths {
            if *column_width <= 0.0 {
                *column_width = fallback;
            }
        }
        assigned_width = column_widths.iter().sum::<f32>();
    }

    if assigned_width <= 0.0 {
        return vec![total_width / column_count as f32; column_count];
    }

    let scale = total_width / assigned_width;
    for column_width in &mut column_widths {
        *column_width *= scale;
    }

    column_widths
}

fn resolve_row_cells<'a>(row: &'a TableRow, column_widths: &[f32]) -> Vec<ResolvedTableCell<'a>> {
    let mut column_index = 0usize;
    let mut x_offset = 0.0;
    let mut cells = Vec::new();

    for cell in row.cells() {
        let remaining_columns = column_widths.len().saturating_sub(column_index);
        if remaining_columns == 0 {
            break;
        }
        let span = normalized_grid_span(cell.properties().grid_span, remaining_columns);
        let width = column_widths[column_index..column_index + span]
            .iter()
            .sum::<f32>();
        cells.push(ResolvedTableCell {
            cell,
            x_offset,
            width,
        });
        column_index += span;
        x_offset += width;
    }

    cells
}

fn normalized_grid_span(grid_span: Option<u32>, remaining_columns: usize) -> usize {
    grid_span.unwrap_or(1).max(1).min(remaining_columns as u32) as usize
}

fn place_table_row(layout: &mut PdfLayout, mut row: RowLayout, repeated_headers: &[RowLayout]) {
    let mut page_is_fresh_for_row = layout.cursor_y <= layout.settings.margin_top + 0.01;

    loop {
        let available_height = page_remaining_height(layout);
        if row.height <= available_height {
            render_row_layout(layout, &row);
            break;
        }

        if row.allow_split_across_pages {
            if let Some((current_page, remaining)) =
                split_row_layout(&row, available_height, layout.settings)
            {
                render_row_layout(layout, &current_page);
                start_new_table_page(layout, repeated_headers);
                row = remaining;
                page_is_fresh_for_row = true;
                continue;
            }
        }

        if !page_is_fresh_for_row {
            start_new_table_page(layout, repeated_headers);
            page_is_fresh_for_row = true;
            continue;
        }

        render_row_layout(layout, &row);
        break;
    }
}

fn start_new_table_page(layout: &mut PdfLayout, repeated_headers: &[RowLayout]) {
    layout.push_page();
    for header in repeated_headers {
        render_row_layout(layout, header);
    }
}

fn render_row_layout(layout: &mut PdfLayout, row: &RowLayout) {
    let y_top = layout.cursor_y;

    for cell in &row.cells {
        layout.push_op(DrawOp::Rect {
            x: layout.settings.margin_x + cell.x_offset,
            y_top,
            width: cell.width,
            height: row.height,
            fill: cell.background,
            stroke: Some((
                layout.settings.table_grid_stroke_color,
                layout.settings.table_grid_stroke_width,
            )),
        });

        for line in &cell.lines {
            layout.push_op(DrawOp::TextLine {
                x: layout.settings.margin_x + cell.x_offset + layout.settings.table_cell_padding_x,
                y_top: y_top + line.y_offset,
                line: line.layout.clone(),
                max_width: cell.width - (layout.settings.table_cell_padding_x * 2.0),
            });
        }
    }

    layout.cursor_y += row.height;
}

fn page_remaining_height(layout: &PdfLayout) -> f32 {
    (layout.settings.page_height - layout.settings.margin_bottom - layout.cursor_y).max(0.0)
}

fn split_row_layout(
    row: &RowLayout,
    available_height: f32,
    settings: PdfRenderSettings,
) -> Option<(RowLayout, RowLayout)> {
    if available_height <= (settings.table_cell_padding_y * 2.0) {
        return None;
    }

    let mut current_cells = Vec::with_capacity(row.cells.len());
    let mut remaining_cells = Vec::with_capacity(row.cells.len());
    let mut current_height = settings.table_cell_padding_y * 2.0;
    let mut remaining_height = settings.table_cell_padding_y * 2.0;
    let mut has_current_lines = false;
    let mut has_remaining_lines = false;

    for cell in &row.cells {
        let fitting_lines = cell
            .lines
            .iter()
            .take_while(|line| {
                line.y_offset + line.layout.line_height + settings.table_cell_padding_y
                    <= available_height + 0.01
            })
            .count();
        if fitting_lines > 0 {
            has_current_lines = true;
        }
        if fitting_lines < cell.lines.len() {
            has_remaining_lines = true;
        }

        let current_lines = cell.lines[..fitting_lines].to_vec();
        let remaining_lines =
            normalize_continued_cell_lines(&cell.lines[fitting_lines..], settings);
        current_height = current_height.max(cell_height(&current_lines, settings));
        remaining_height = remaining_height.max(cell_height(&remaining_lines, settings));
        current_cells.push(CellLayout {
            x_offset: cell.x_offset,
            width: cell.width,
            background: cell.background,
            lines: current_lines,
        });
        remaining_cells.push(CellLayout {
            x_offset: cell.x_offset,
            width: cell.width,
            background: cell.background,
            lines: remaining_lines,
        });
    }

    if !has_current_lines || !has_remaining_lines {
        return None;
    }

    Some((
        RowLayout {
            cells: current_cells,
            height: current_height.max(MIN_CONTENT_WIDTH),
            allow_split_across_pages: row.allow_split_across_pages,
            repeat_as_header: row.repeat_as_header,
        },
        RowLayout {
            cells: remaining_cells,
            height: remaining_height.max(MIN_CONTENT_WIDTH),
            allow_split_across_pages: row.allow_split_across_pages,
            repeat_as_header: false,
        },
    ))
}

fn normalize_continued_cell_lines(
    lines: &[CellLine],
    settings: PdfRenderSettings,
) -> Vec<CellLine> {
    if lines.is_empty() {
        return Vec::new();
    }

    let mut normalized = Vec::with_capacity(lines.len());
    let mut y_offset = settings.table_cell_padding_y;
    normalized.push(CellLine {
        y_offset,
        layout: lines[0].layout.clone(),
    });

    for pair in lines.windows(2) {
        y_offset += pair[1].y_offset - pair[0].y_offset;
        normalized.push(CellLine {
            y_offset,
            layout: pair[1].layout.clone(),
        });
    }

    normalized
}

fn cell_height(lines: &[CellLine], settings: PdfRenderSettings) -> f32 {
    lines
        .last()
        .map(|line| line.y_offset + line.layout.line_height + settings.table_cell_padding_y)
        .unwrap_or(settings.table_cell_padding_y * 2.0)
}

fn layout_paragraph_lines(
    styles: &Stylesheet,
    paragraph: &Paragraph,
    max_width: f32,
    settings: PdfRenderSettings,
    font_system: &mut PdfFontSystem,
) -> crate::Result<Vec<LineLayout>> {
    let alignment = styles
        .resolve_paragraph(paragraph)?
        .alignment
        .unwrap_or(ParagraphAlignment::Left);
    let mut lines = Vec::new();
    let mut current_spans = Vec::new();
    let mut current_width = 0.0;
    let mut current_line_height = settings.default_line_height;

    for run in paragraph.runs() {
        let run_properties = styles.resolve_run(paragraph, run)?;
        let style = style_from_run(&run_properties, settings, font_system.default_family());
        for token in tokenize(run.text()) {
            if token == "\n" {
                flush_line(
                    &mut lines,
                    &mut current_spans,
                    &mut current_width,
                    &mut current_line_height,
                    alignment.clone(),
                    settings.default_line_height,
                );
                continue;
            }

            if token.trim().is_empty() && current_spans.is_empty() {
                continue;
            }

            let shaped = font_system.shape_text(&style, &token)?;
            if shaped.spans.is_empty() {
                continue;
            }

            if current_width + shaped.width > max_width
                && !current_spans.is_empty()
                && !token.trim().is_empty()
            {
                flush_line(
                    &mut lines,
                    &mut current_spans,
                    &mut current_width,
                    &mut current_line_height,
                    alignment.clone(),
                    settings.default_line_height,
                );
            }

            if shaped.spans.iter().all(|span| span.text.trim().is_empty())
                && current_spans.is_empty()
            {
                continue;
            }

            current_line_height =
                current_line_height.max(settings.effective_line_height(style.size));
            current_width += shaped.width;
            current_spans.extend(shaped.spans);
        }
    }

    flush_line(
        &mut lines,
        &mut current_spans,
        &mut current_width,
        &mut current_line_height,
        alignment,
        settings.default_line_height,
    );

    Ok(lines)
}

fn flush_line(
    lines: &mut Vec<LineLayout>,
    current_spans: &mut Vec<TextSpan>,
    current_width: &mut f32,
    current_line_height: &mut f32,
    alignment: ParagraphAlignment,
    default_line_height: f32,
) {
    if current_spans.is_empty() {
        *current_width = 0.0;
        *current_line_height = default_line_height;
        return;
    }

    lines.push(LineLayout {
        spans: std::mem::take(current_spans),
        width: *current_width,
        line_height: *current_line_height,
        alignment,
    });
    *current_width = 0.0;
    *current_line_height = default_line_height;
}

fn tokenize(text: &str) -> Vec<String> {
    let mut tokens = Vec::new();
    let mut current = String::new();
    let mut current_is_space = None;

    for ch in text.chars() {
        if ch == '\n' {
            if !current.is_empty() {
                tokens.push(std::mem::take(&mut current));
            }
            current_is_space = None;
            tokens.push("\n".to_string());
            continue;
        }

        let normalized = if ch == '\t' { ' ' } else { ch };
        let is_space = normalized.is_whitespace();

        match current_is_space {
            Some(mode) if mode == is_space => current.push(normalized),
            Some(_) => {
                tokens.push(std::mem::take(&mut current));
                current.push(normalized);
                current_is_space = Some(is_space);
            }
            None => {
                current.push(normalized);
                current_is_space = Some(is_space);
            }
        }
    }

    if !current.is_empty() {
        tokens.push(current);
    }

    tokens
}

fn style_from_run(
    properties: &RunProperties,
    settings: PdfRenderSettings,
    default_family: &str,
) -> RequestedTextStyle {
    let font = match (properties.bold, properties.italic) {
        (true, true) => PdfFont::BoldOblique,
        (true, false) => PdfFont::Bold,
        (false, true) => PdfFont::Oblique,
        (false, false) => PdfFont::Regular,
    };

    RequestedTextStyle {
        font_request: PdfFontRequest::new(
            properties
                .font_family
                .as_deref()
                .map(str::trim)
                .filter(|family| !family.is_empty())
                .unwrap_or(default_family),
            font,
        ),
        size: properties
            .font_size
            .map(|value| f32::from(value) / 2.0)
            .unwrap_or(settings.default_text_size),
        color: properties
            .color
            .as_deref()
            .and_then(parse_hex_color)
            .unwrap_or(settings.default_text_color),
    }
}

fn twips_to_points(twips: u32) -> f32 {
    twips as f32 / 20.0
}

fn points_to_twips(points: f32) -> u32 {
    if !points.is_finite() || points <= 0.0 {
        0
    } else if points >= u32::MAX as f32 / 20.0 {
        u32::MAX
    } else {
        (points * 20.0).round() as u32
    }
}

fn points_to_u16(points: f32) -> u16 {
    if !points.is_finite() || points <= 0.0 {
        0
    } else if points >= f32::from(u16::MAX) {
        u16::MAX
    } else {
        points.round() as u16
    }
}

fn clamp_u32_to_u16(value: u32) -> u16 {
    if value > u32::from(u16::MAX) {
        u16::MAX
    } else {
        value as u16
    }
}

fn normalize_positive(value: f32, fallback: f32) -> f32 {
    if value.is_finite() && value > 0.0 {
        value
    } else {
        fallback
    }
}

fn normalize_non_negative(value: f32, fallback: f32) -> f32 {
    if value.is_finite() && value >= 0.0 {
        value
    } else {
        fallback
    }
}

fn parse_hex_color(value: &str) -> Option<Rgb> {
    let value = value.trim().trim_start_matches('#');
    if value.len() != 6 {
        return None;
    }

    let red = u8::from_str_radix(&value[0..2], 16).ok()?;
    let green = u8::from_str_radix(&value[2..4], 16).ok()?;
    let blue = u8::from_str_radix(&value[4..6], 16).ok()?;
    Some(Rgb(red, green, blue))
}

fn normalize_pdf_text(text: &str) -> String {
    text.chars()
        .filter_map(|ch| match ch {
            '\r' => None,
            '\u{00A0}' => Some(' '),
            ch if ch.is_control() && ch != '\n' && ch != '\t' => None,
            _ => Some(ch),
        })
        .collect()
}

fn scale_font_units(value: f32, units_per_em: f32) -> f32 {
    if units_per_em <= 0.0 {
        value
    } else {
        value * (1000.0 / units_per_em)
    }
}

fn supports_embedded_truetype(font_data: &[u8]) -> bool {
    matches!(font_data.get(0..4), Some(b"\0\x01\0\0") | Some(b"true"))
        && ttf_parser::fonts_in_collection(font_data).unwrap_or(1) == 1
}

fn system_font_db() -> &'static FontDatabase {
    static DB: OnceLock<FontDatabase> = OnceLock::new();
    DB.get_or_init(|| {
        let mut db = FontDatabase::new();
        db.load_system_fonts();
        db
    })
}

fn sanitize_pdf_name_component(value: &str) -> String {
    let mut sanitized = String::new();
    for ch in value.chars() {
        if ch.is_ascii_alphanumeric() || matches!(ch, '-' | '_' | '+') {
            sanitized.push(ch);
        } else if !sanitized.ends_with('-') {
            sanitized.push('-');
        }
    }
    let trimmed = sanitized.trim_matches('-');
    if trimmed.is_empty() {
        "RusdoxFont".to_string()
    } else {
        trimmed.to_string()
    }
}

fn split_font_family_list(value: &str) -> Vec<String> {
    value
        .split(',')
        .map(str::trim)
        .filter(|family| !family.is_empty())
        .map(|family| family.trim_matches('"').trim_matches('\'').to_string())
        .collect()
}

fn base_font_family_candidates(requested_family: &str) -> Vec<String> {
    let mut candidates = Vec::new();
    for family in split_font_family_list(requested_family) {
        push_font_family_candidate(&mut candidates, &family);
        match family.to_ascii_lowercase().as_str() {
            "sans-serif" | "sans serif" => {
                for fallback in PDF_SANS_FALLBACKS {
                    push_font_family_candidate(&mut candidates, fallback);
                }
            }
            "serif" => {
                for fallback in PDF_SERIF_FALLBACKS {
                    push_font_family_candidate(&mut candidates, fallback);
                }
            }
            "monospace" => {
                for fallback in PDF_MONO_FALLBACKS {
                    push_font_family_candidate(&mut candidates, fallback);
                }
            }
            _ => {}
        }
    }

    for fallback in PDF_SANS_FALLBACKS {
        push_font_family_candidate(&mut candidates, fallback);
    }

    candidates
}

fn push_font_family_candidate(candidates: &mut Vec<String>, family: &str) {
    let trimmed = family.trim();
    if trimmed.is_empty() {
        return;
    }
    if !candidates.iter().any(|candidate| candidate == trimmed) {
        candidates.push(trimmed.to_string());
    }
}

fn script_specific_fallback_families(ch: char) -> &'static [&'static str] {
    match ch as u32 {
        0x0590..=0x05FF => PDF_HEBREW_FALLBACKS,
        0x0600..=0x06FF | 0x0750..=0x077F | 0x08A0..=0x08FF => PDF_ARABIC_FALLBACKS,
        0x0900..=0x097F => PDF_DEVANAGARI_FALLBACKS,
        0x1100..=0x11FF
        | 0x2E80..=0xA4CF
        | 0xAC00..=0xD7AF
        | 0xF900..=0xFAFF
        | 0xFE30..=0xFE4F
        | 0xFF00..=0xFFEF
        | 0x20000..=0x2FA1F => PDF_CJK_FALLBACKS,
        _ => PDF_SANS_FALLBACKS,
    }
}

struct PdfContentWriter {
    buffer: Vec<u8>,
    current_fill: Option<Rgb>,
    current_stroke: Option<Rgb>,
    current_line_width: Option<f32>,
    page_height: f32,
    baseline_factor: f32,
}

impl PdfContentWriter {
    fn with_capacity(capacity: usize, settings: PdfRenderSettings) -> Self {
        Self {
            buffer: Vec::with_capacity(capacity),
            current_fill: None,
            current_stroke: None,
            current_line_width: None,
            page_height: settings.page_height,
            baseline_factor: settings.baseline_factor,
        }
    }

    fn finish(self) -> Vec<u8> {
        self.buffer
    }

    fn draw_rect(
        &mut self,
        x: f32,
        y_top: f32,
        width: f32,
        height: f32,
        fill: Option<Rgb>,
        stroke: Option<(Rgb, f32)>,
    ) {
        let y = self.page_height - y_top - height;

        if let Some(color) = fill {
            self.set_fill_color(color);
        }
        if let Some((color, line_width)) = stroke {
            self.set_stroke_color(color);
            self.set_line_width(line_width);
        }

        self.push_f32(x);
        self.buffer.push(b' ');
        self.push_f32(y);
        self.buffer.push(b' ');
        self.push_f32(width);
        self.buffer.push(b' ');
        self.push_f32(height);
        self.buffer.extend_from_slice(b" re ");

        match (fill.is_some(), stroke.is_some()) {
            (true, true) => self.buffer.extend_from_slice(b"B\n"),
            (true, false) => self.buffer.extend_from_slice(b"f\n"),
            (false, true) => self.buffer.extend_from_slice(b"S\n"),
            (false, false) => {}
        }
    }

    fn draw_line(
        &mut self,
        x: f32,
        y_top: f32,
        line: &LineLayout,
        max_width: f32,
        font_system: &PdfFontSystem,
    ) {
        if line.spans.is_empty() {
            return;
        }

        let start_x = match line.alignment {
            ParagraphAlignment::Center => x + ((max_width - line.width).max(0.0) / 2.0),
            ParagraphAlignment::Right => x + (max_width - line.width).max(0.0),
            _ => x,
        };
        let baseline_y = self.page_height - y_top - (line.line_height * self.baseline_factor);
        let mut cursor_x = start_x;
        let mut current_font = None;
        let mut current_color = None;

        self.buffer.extend_from_slice(b"BT\n");
        for span in &line.spans {
            let font_key = (span.style.font_id, span.style.size.to_bits());
            if current_font != Some(font_key) {
                self.buffer.push(b'/');
                self.buffer.extend_from_slice(
                    font_system
                        .face(span.style.font_id)
                        .resource_name
                        .as_bytes(),
                );
                self.buffer.push(b' ');
                self.push_f32(span.style.size);
                self.buffer.extend_from_slice(b" Tf\n");
                current_font = Some(font_key);
            }

            if current_color != Some(span.style.color) {
                self.set_fill_color(span.style.color);
                current_color = Some(span.style.color);
            }

            self.buffer.extend_from_slice(b"1 0 0 1 ");
            self.push_f32(cursor_x);
            self.buffer.push(b' ');
            self.push_f32(baseline_y);
            self.buffer.extend_from_slice(b" Tm\n");
            self.push_pdf_glyphs(&span.glyphs);
            self.buffer.extend_from_slice(b" Tj\n");

            cursor_x += span.width;
        }
        self.buffer.extend_from_slice(b"ET\n");
    }

    fn draw_image(&mut self, x: f32, y_top: f32, width: f32, height: f32, resource_name: &str) {
        let y = self.page_height - y_top - height;
        self.buffer.extend_from_slice(b"q\n");
        self.push_f32(width);
        self.buffer.extend_from_slice(b" 0 0 ");
        self.push_f32(height);
        self.buffer.push(b' ');
        self.push_f32(x);
        self.buffer.push(b' ');
        self.push_f32(y);
        self.buffer.extend_from_slice(b" cm\n/");
        self.buffer.extend_from_slice(resource_name.as_bytes());
        self.buffer.extend_from_slice(b" Do\nQ\n");
    }

    fn set_fill_color(&mut self, color: Rgb) {
        if self.current_fill == Some(color) {
            return;
        }
        self.push_rgb(color);
        self.buffer.extend_from_slice(b" rg\n");
        self.current_fill = Some(color);
    }

    fn set_stroke_color(&mut self, color: Rgb) {
        if self.current_stroke == Some(color) {
            return;
        }
        self.push_rgb(color);
        self.buffer.extend_from_slice(b" RG\n");
        self.current_stroke = Some(color);
    }

    fn set_line_width(&mut self, width: f32) {
        if self.current_line_width == Some(width) {
            return;
        }
        self.push_f32(width);
        self.buffer.extend_from_slice(b" w\n");
        self.current_line_width = Some(width);
    }

    fn push_rgb(&mut self, color: Rgb) {
        self.push_f32(f32::from(color.0) / 255.0);
        self.buffer.push(b' ');
        self.push_f32(f32::from(color.1) / 255.0);
        self.buffer.push(b' ');
        self.push_f32(f32::from(color.2) / 255.0);
    }

    fn push_f32(&mut self, value: f32) {
        let mut formatted = String::new();
        let _ = write!(&mut formatted, "{value:.2}");
        let trimmed = formatted.trim_end_matches('0').trim_end_matches('.');
        if trimmed.is_empty() {
            self.buffer.push(b'0');
        } else {
            self.buffer.extend_from_slice(trimmed.as_bytes());
        }
    }

    fn push_pdf_glyphs(&mut self, glyphs: &[u16]) {
        self.buffer.push(b'<');
        for &glyph in glyphs {
            self.push_hex_u16(glyph);
        }
        self.buffer.push(b'>');
    }

    fn push_hex_u16(&mut self, value: u16) {
        const HEX: &[u8; 16] = b"0123456789ABCDEF";
        self.buffer.push(HEX[((value >> 12) & 0x0F) as usize]);
        self.buffer.push(HEX[((value >> 8) & 0x0F) as usize]);
        self.buffer.push(HEX[((value >> 4) & 0x0F) as usize]);
        self.buffer.push(HEX[(value & 0x0F) as usize]);
    }
}

fn render_page_content(
    page: &Page,
    settings: PdfRenderSettings,
    font_system: &PdfFontSystem,
    images: &[PdfImageAsset],
) -> Vec<u8> {
    let mut writer = PdfContentWriter::with_capacity(page.ops.len() * 96, settings);

    for op in &page.ops {
        match op {
            DrawOp::TextLine {
                x,
                y_top,
                line,
                max_width,
            } => writer.draw_line(*x, *y_top, line, *max_width, font_system),
            DrawOp::Rect {
                x,
                y_top,
                width,
                height,
                fill,
                stroke,
            } => writer.draw_rect(*x, *y_top, *width, *height, *fill, *stroke),
            DrawOp::Image {
                x,
                y_top,
                width,
                height,
                image_id,
            } => writer.draw_image(
                *x,
                *y_top,
                *width,
                *height,
                &images[*image_id].resource_name,
            ),
        }
    }

    writer.finish()
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeSet;
    use std::fmt::Write as _;
    use std::{fs, path::PathBuf};

    use tempfile::tempdir;

    use super::{
        clamp_u32_to_u16, layout_document, layout_paragraph_lines, layout_row, normalize_pdf_text,
        parse_hex_color, points_to_u16, resolve_table_column_widths, style_from_run, tokenize,
        twips_to_points, DrawOp, Page, PdfFont, PdfFontRequest, PdfFontSystem, PdfRenderSettings,
        RequestedTextStyle, Rgb, Studio, DEFAULT_BASELINE_FACTOR, DEFAULT_LINE_HEIGHT,
        DEFAULT_LINE_HEIGHT_MULTIPLIER, DEFAULT_TABLE_AFTER_SPACING,
        DEFAULT_TABLE_GRID_STROKE_WIDTH, DEFAULT_TABLE_ROW_PADDING_X, DEFAULT_TABLE_ROW_PADDING_Y,
        DEFAULT_TEXT_SIZE, DEFAULT_TEXT_WIDTH_BIAS_BOLD, DEFAULT_TEXT_WIDTH_BIAS_REGULAR,
        MIN_CONTENT_WIDTH,
    };
    use crate::config::RusdoxConfig;
    use crate::spec::{
        DocumentSpec, ParagraphAlignmentSpec, ParagraphSpec, RunSpec, UnderlineStyleSpec,
        VerticalAlignSpec,
    };
    use crate::{
        Document, DocumentBlockRef, DocumentMetadata, HeaderFooter, PageNumberFormat,
        PageNumbering, PageSetup, Paragraph, ParagraphAlignment, ParagraphList, Run, Stylesheet,
        Table, TableCell, TableRow, VerticalAlign, Visual, VisualFormat, VisualKind,
    };

    fn default_pdf_settings() -> PdfRenderSettings {
        PdfRenderSettings::from_config(&RusdoxConfig::default())
    }

    fn pdf_font_system(config: &RusdoxConfig) -> PdfFontSystem {
        PdfFontSystem::new(config, PdfRenderSettings::from_config(config))
    }

    fn default_text_style(config: &RusdoxConfig, font: PdfFont, size: f32) -> RequestedTextStyle {
        RequestedTextStyle {
            font_request: PdfFontRequest::new(config.typography.font_family.clone(), font),
            size,
            color: Rgb(15, 23, 42),
        }
    }

    fn utf16_hex(text: &str) -> String {
        let mut encoded = String::new();
        for value in text.encode_utf16() {
            let _ = write!(&mut encoded, "{value:04X}");
        }
        encoded
    }

    fn assert_close(actual: f32, expected: f32) {
        assert!(
            (actual - expected).abs() < 0.01,
            "expected {expected:.2}, got {actual:.2}"
        );
    }

    fn page_line_texts(page: &Page) -> Vec<String> {
        page.ops
            .iter()
            .filter_map(|op| match op {
                DrawOp::TextLine { line, .. } => Some(
                    line.spans
                        .iter()
                        .map(|span| span.text.as_str())
                        .collect::<String>(),
                ),
                DrawOp::Rect { .. } => None,
                DrawOp::Image { .. } => None,
            })
            .collect()
    }

    fn find_family_triggering_multiscript_fallback() -> Option<String> {
        let mut tested_families = BTreeSet::new();
        let mut font_system = pdf_font_system(&RusdoxConfig::default());

        for face in super::system_font_db().faces() {
            for (family, _) in &face.families {
                if !tested_families.insert(family.clone()) {
                    continue;
                }

                let style = RequestedTextStyle {
                    font_request: PdfFontRequest::new(family.clone(), PdfFont::Regular),
                    size: DEFAULT_TEXT_SIZE,
                    color: Rgb(15, 23, 42),
                };
                let shaped = match font_system.shape_text(&style, "Latin 你好") {
                    Ok(shaped) => shaped,
                    Err(_) => continue,
                };
                let distinct_fonts = shaped
                    .spans
                    .iter()
                    .map(|span| span.style.font_id)
                    .collect::<BTreeSet<_>>();
                if distinct_fonts.len() >= 2 {
                    return Some(family.clone());
                }
            }
        }

        None
    }

    fn assert_multiscript_pdf_output_with_optional_fallback(fallback_family: Option<String>) {
        let temp = tempdir().expect("temp dir");
        let mut config = RusdoxConfig::default();
        config.output.docx_dir = temp.path().join("docx").to_string_lossy().to_string();
        config.output.pdf_dir = temp.path().join("pdf").to_string_lossy().to_string();
        config.output.emit_pdf_preview = true;
        if let Some(family) = fallback_family.clone() {
            config.typography.font_family = family;

            let mut font_system = pdf_font_system(&config);
            let shaped = font_system
                .shape_text(
                    &default_text_style(&config, PdfFont::Regular, config.typography.body_size_pt),
                    "Latin 你好",
                )
                .expect("shape text");
            let distinct_fonts = shaped
                .spans
                .iter()
                .map(|span| span.style.font_id)
                .collect::<BTreeSet<_>>();
            let used_families = distinct_fonts
                .iter()
                .map(|&font_id| font_system.face(font_id).family_name.clone())
                .collect::<Vec<_>>();
            assert!(
                distinct_fonts.len() >= 2,
                "expected multiple PDF faces for '{}', got {:?}",
                config.typography.font_family,
                used_families
            );
        }

        let studio = Studio::new(config);

        let mut document = Document::new();
        document.push_paragraph(Paragraph::new().add_run(Run::from_text("Latin 你好")));
        let docx_path = temp.path().join("docx/fallback.docx");

        studio
            .save_with_pdf_stats(&document, &docx_path)
            .expect("save should succeed");

        let pdf_path = temp.path().join("pdf/fallback.pdf");
        let pdf = fs::read(&pdf_path).expect("read pdf");
        let pdf_text = String::from_utf8_lossy(&pdf);

        let font_file_count = pdf_text.matches("/FontFile2").count();
        if fallback_family.is_some() {
            assert!(
                font_file_count >= 2,
                "expected multiple embedded fonts when fallback is available, got {font_file_count}"
            );
        } else {
            assert!(
                font_file_count >= 1,
                "expected at least one embedded font for mixed-script PDF output"
            );
        }
        assert!(pdf_text.contains(&format!("<{}>", utf16_hex("你"))));
        assert!(pdf_text.contains(&format!("<{}>", utf16_hex("好"))));
    }

    #[test]
    fn points_to_u16_clamps_and_rounds() {
        assert_eq!(points_to_u16(-1.0), 0);
        assert_eq!(points_to_u16(0.0), 0);
        assert_eq!(points_to_u16(12.49), 12);
        assert_eq!(points_to_u16(12.5), 13);
        assert_eq!(points_to_u16(f32::INFINITY), 0);
        assert_eq!(points_to_u16(1_000_000.0), u16::MAX);
    }

    #[test]
    fn clamp_u32_to_u16_respects_upper_bound() {
        assert_eq!(clamp_u32_to_u16(0), 0);
        assert_eq!(clamp_u32_to_u16(42), 42);
        assert_eq!(clamp_u32_to_u16(u32::from(u16::MAX)), u16::MAX);
        assert_eq!(clamp_u32_to_u16(u32::from(u16::MAX) + 1), u16::MAX);
    }

    #[test]
    fn parse_hex_color_accepts_and_rejects_expected_formats() {
        assert!(parse_hex_color("#A1b2C3") == Some(Rgb(161, 178, 195)));
        assert!(parse_hex_color("0F172A") == Some(Rgb(15, 23, 42)));
        assert!(parse_hex_color("").is_none());
        assert!(parse_hex_color("FFF").is_none());
        assert!(parse_hex_color("XYZ123").is_none());
    }

    #[test]
    fn normalize_pdf_text_preserves_unicode_and_normalizes_spaces() {
        assert_eq!(normalize_pdf_text("•\u{00A0}ok"), "• ok");
        assert_eq!(normalize_pdf_text("Привет 你好"), "Привет 你好");
        assert_eq!(normalize_pdf_text("line\rbreak"), "linebreak");
    }

    #[test]
    fn tokenize_splits_words_spaces_tabs_and_newlines() {
        let tokens = tokenize("a\tb  c\nd");
        assert_eq!(tokens, vec!["a", " ", "b", "  ", "c", "\n", "d"]);
    }

    #[test]
    fn pdf_font_system_shapes_unicode_without_ascii_degradation() {
        let config = RusdoxConfig::default();
        let style = default_text_style(&config, PdfFont::Regular, 12.0);
        let mut font_system = pdf_font_system(&config);

        let shaped = font_system
            .shape_text(&style, "Привет 你好")
            .expect("shape text");

        assert!(shaped.width > 0.0);
        assert_eq!(
            shaped
                .spans
                .iter()
                .map(|span| span.text.as_str())
                .collect::<String>(),
            "Привет 你好"
        );
        assert!(shaped.spans.iter().all(|span| !span.glyphs.is_empty()));
    }

    #[test]
    fn style_from_run_maps_font_size_and_color_and_weight() {
        let run = Run::from_text("x")
            .bold()
            .italic()
            .size_points(18)
            .color("AABBCC");
        let style = style_from_run(run.properties(), default_pdf_settings(), "Arial");
        assert!(style.font_request.font == PdfFont::BoldOblique);
        assert_eq!(style.font_request.family, "Arial");
        assert_eq!(style.size, 18.0);
        assert!(style.color == Rgb(170, 187, 204));
    }

    #[test]
    fn pdf_render_settings_honor_config_values() {
        let mut config = RusdoxConfig::default();
        config.pdf.page_width_pt = 500.0;
        config.pdf.page_height_pt = 700.0;
        config.pdf.margin_x_pt = 40.0;
        config.pdf.margin_top_pt = 30.0;
        config.pdf.margin_bottom_pt = 45.0;
        config.pdf.default_text_size_pt = 13.0;
        config.pdf.default_line_height_pt = 17.0;
        config.pdf.line_height_multiplier = 1.6;
        config.pdf.baseline_factor = 0.78;
        config.pdf.text_width_bias_regular = 0.92;
        config.pdf.text_width_bias_bold = 1.18;
        config.table.pdf_cell_padding_x_pt = 9.0;
        config.table.pdf_cell_padding_y_pt = 11.0;
        config.table.pdf_after_spacing_pt = 15.0;
        config.table.pdf_grid_stroke_width_pt = 1.25;
        config.colors.table_border = "112233".to_string();
        config.colors.ink = "445566".to_string();

        let settings = PdfRenderSettings::from_config(&config);

        assert_close(settings.page_width, 500.0);
        assert_close(settings.page_height, 700.0);
        assert_close(settings.margin_x, 40.0);
        assert_close(settings.margin_top, 30.0);
        assert_close(settings.margin_bottom, 45.0);
        assert_close(settings.content_width, 420.0);
        assert_close(settings.default_text_size, 13.0);
        assert_close(settings.default_line_height, 17.0);
        assert_close(settings.line_height_multiplier, 1.6);
        assert_close(settings.baseline_factor, 0.78);
        assert_close(settings.text_width_bias_regular, 0.92);
        assert_close(settings.text_width_bias_bold, 1.18);
        assert_close(settings.table_cell_padding_x, 9.0);
        assert_close(settings.table_cell_padding_y, 11.0);
        assert_close(settings.table_after_spacing, 15.0);
        assert_close(settings.table_grid_stroke_width, 1.25);
        assert_eq!(settings.table_grid_stroke_color, Rgb(17, 34, 51));
        assert_eq!(settings.default_text_color, Rgb(68, 85, 102));
    }

    #[test]
    fn pdf_font_system_uses_configured_bias_values() {
        let mut narrow_config = RusdoxConfig::default();
        narrow_config.pdf.text_width_bias_regular = 0.5;
        let mut wide_config = narrow_config.clone();
        wide_config.pdf.text_width_bias_regular = 1.5;

        let style_narrow = default_text_style(&narrow_config, PdfFont::Regular, 10.0);
        let style_wide = default_text_style(&wide_config, PdfFont::Regular, 10.0);
        let mut narrow_fonts = pdf_font_system(&narrow_config);
        let mut wide_fonts = pdf_font_system(&wide_config);

        let narrow = narrow_fonts
            .shape_text(&style_narrow, "Hello")
            .expect("shape narrow")
            .width;
        let wide = wide_fonts
            .shape_text(&style_wide, "Hello")
            .expect("shape wide")
            .width;

        assert!(narrow > 0.0);
        assert_close(wide / narrow, 3.0);
    }

    #[test]
    fn pdf_render_settings_normalize_invalid_and_extreme_values() {
        let mut config = RusdoxConfig::default();
        config.pdf.page_width_pt = 40.0;
        config.pdf.page_height_pt = 8.0;
        config.pdf.margin_x_pt = 30.0;
        config.pdf.margin_top_pt = 99.0;
        config.pdf.margin_bottom_pt = 99.0;
        config.pdf.default_text_size_pt = f32::NAN;
        config.pdf.default_line_height_pt = -1.0;
        config.pdf.line_height_multiplier = 0.0;
        config.pdf.baseline_factor = f32::INFINITY;
        config.pdf.text_width_bias_regular = -2.0;
        config.pdf.text_width_bias_bold = f32::NAN;
        config.table.pdf_cell_padding_x_pt = -4.0;
        config.table.pdf_cell_padding_y_pt = f32::NAN;
        config.table.pdf_after_spacing_pt = -2.0;
        config.table.pdf_grid_stroke_width_pt = f32::NAN;

        let settings = PdfRenderSettings::from_config(&config);

        assert_close(settings.page_width, 40.0);
        assert_close(settings.page_height, DEFAULT_LINE_HEIGHT);
        assert_close(settings.margin_x, 8.0);
        assert_close(settings.margin_top, 0.0);
        assert_close(settings.margin_bottom, 0.0);
        assert_close(settings.content_width, MIN_CONTENT_WIDTH);
        assert_close(settings.default_text_size, DEFAULT_TEXT_SIZE);
        assert_close(settings.default_line_height, DEFAULT_LINE_HEIGHT);
        assert_close(
            settings.line_height_multiplier,
            DEFAULT_LINE_HEIGHT_MULTIPLIER,
        );
        assert_close(settings.baseline_factor, DEFAULT_BASELINE_FACTOR);
        assert_close(
            settings.text_width_bias_regular,
            DEFAULT_TEXT_WIDTH_BIAS_REGULAR,
        );
        assert_close(settings.text_width_bias_bold, DEFAULT_TEXT_WIDTH_BIAS_BOLD);
        assert_close(settings.table_cell_padding_x, DEFAULT_TABLE_ROW_PADDING_X);
        assert_close(settings.table_cell_padding_y, DEFAULT_TABLE_ROW_PADDING_Y);
        assert_close(settings.table_after_spacing, DEFAULT_TABLE_AFTER_SPACING);
        assert_close(
            settings.table_grid_stroke_width,
            DEFAULT_TABLE_GRID_STROKE_WIDTH,
        );
    }

    #[test]
    fn text_run_uses_configured_font_family_and_body_size() {
        let mut config = RusdoxConfig::default();
        config.typography.font_family = "IBM Plex Sans".to_string();
        config.typography.body_size_pt = 12.5;
        let studio = Studio::new(config);

        let run = studio.text_run("hello");

        assert_eq!(run.text(), "hello");
        assert_eq!(
            run.properties().font_family.as_deref(),
            Some("IBM Plex Sans")
        );
        assert_eq!(run.properties().font_size, Some(26));
    }

    #[test]
    fn semantic_list_helpers_do_not_prefix_literal_markers() {
        let studio = Studio::new(RusdoxConfig::default());

        let bullet = studio.bullet("Alpha");
        let numbered = studio.numbered("Beta");

        assert_eq!(bullet.text(), "Alpha");
        assert_eq!(bullet.list(), Some(&ParagraphList::bullet()));
        assert_eq!(numbered.text(), "Beta");
        assert_eq!(numbered.list(), Some(&ParagraphList::numbered()));
    }

    #[test]
    fn high_level_spec_blocks_use_dedicated_config_fields() {
        let mut config = RusdoxConfig::default();
        config.typography.cover_title_size_pt = 32.0;
        config.typography.page_heading_size_pt = 19.5;
        config.typography.tagline_size_pt = 13.0;
        config.spacing.cover_title_before_twips = 900;
        config.spacing.cover_title_after_twips = 210;
        config.spacing.page_heading_after_twips = 160;
        config.spacing.tagline_after_twips = 95;
        let studio = Studio::new(config);

        let cover = studio.cover_title("Board Report");
        let page_heading = studio.page_heading("Financial Review");
        let tagline = studio.tagline("Confidential");

        assert_eq!(cover.spacing_before_value(), Some(900));
        assert_eq!(cover.spacing_after_value(), Some(210));
        assert_eq!(
            cover
                .runs()
                .next()
                .expect("cover run")
                .properties()
                .font_size,
            Some(points_to_u16(32.0).saturating_mul(2))
        );

        assert!(page_heading.has_page_break_before());
        assert_eq!(page_heading.spacing_after_value(), Some(160));
        assert_eq!(
            page_heading
                .runs()
                .next()
                .expect("page heading run")
                .properties()
                .font_size,
            Some(points_to_u16(19.5).saturating_mul(2))
        );

        assert_eq!(tagline.spacing_after_value(), Some(95));
        assert_eq!(
            tagline
                .runs()
                .next()
                .expect("tagline run")
                .properties()
                .font_size,
            Some(points_to_u16(13.0).saturating_mul(2))
        );
    }

    #[test]
    fn compose_applies_document_layout_controls() {
        let studio = Studio::new(RusdoxConfig::default());
        let page_setup = PageSetup::new(13_000, 17_000)
            .margins(900, 950, 1_000, 1_050)
            .header_footer_distances(500, 550)
            .gutter(120);
        let header = HeaderFooter::new("Board Report").with_alignment(ParagraphAlignment::Center);
        let footer =
            HeaderFooter::new("Page {page} of {pages}").with_alignment(ParagraphAlignment::Right);
        let numbering = PageNumbering::new(PageNumberFormat::UpperRoman).start_at(4);
        let mut spec = DocumentSpec::new();
        spec.output_name = Some("board-report".to_string());
        spec.metadata = DocumentMetadata::new()
            .title("Board Report")
            .author("Finance");
        spec.page_setup = Some(page_setup.clone());
        spec.header = Some(header.clone());
        spec.footer = Some(footer.clone());
        spec.page_numbering = Some(numbering.clone());
        spec.blocks = vec![crate::spec::title("Board Report")];

        let document = studio.compose(&spec);

        assert_eq!(document.page_setup(), &page_setup);
        assert_eq!(document.metadata().title.as_deref(), Some("Board Report"));
        assert_eq!(document.metadata().author.as_deref(), Some("Finance"));
        assert_eq!(document.header(), Some(&header));
        assert_eq!(document.footer(), Some(&footer));
        assert_eq!(document.page_numbering(), Some(&numbering));
    }

    #[test]
    fn compose_marks_visual_intro_paragraphs_keep_next() {
        let studio = Studio::new(RusdoxConfig::default());
        let mut spec = DocumentSpec::new();
        spec.blocks = vec![
            crate::spec::title("Visual Assets"),
            crate::spec::subtitle("Rendered by RusDox"),
            crate::spec::logo("mark.svg"),
            crate::spec::section("Benchmark"),
            crate::spec::body("Narrative before the chart."),
            crate::spec::chart("chart.svg"),
        ];

        let document = studio.compose(&spec);
        let blocks = document.blocks().collect::<Vec<_>>();

        let DocumentBlockRef::Paragraph(title) = blocks[0] else {
            panic!("expected title paragraph");
        };
        let DocumentBlockRef::Paragraph(subtitle) = blocks[1] else {
            panic!("expected subtitle paragraph");
        };
        let DocumentBlockRef::Paragraph(section) = blocks[3] else {
            panic!("expected section paragraph");
        };
        let DocumentBlockRef::Paragraph(body) = blocks[4] else {
            panic!("expected body paragraph");
        };

        assert!(title.has_keep_next());
        assert!(subtitle.has_keep_next());
        assert!(section.has_keep_next());
        assert!(body.has_keep_next());
    }

    #[test]
    fn paragraph_spec_uses_config_as_base_and_applies_overrides() {
        let mut config = RusdoxConfig::default();
        config.typography.font_family = "IBM Plex Sans".to_string();
        config.typography.body_size_pt = 12.0;
        let studio = Studio::new(config);

        let paragraph = studio.paragraph_from_spec(&ParagraphSpec {
            runs: vec![
                RunSpec {
                    text: "Hello".to_string(),
                    bold: true,
                    ..RunSpec::default()
                },
                RunSpec {
                    text: " world".to_string(),
                    italic: true,
                    underline: Some(UnderlineStyleSpec::Single),
                    color: Some("C1121F".to_string()),
                    size_pt: Some(14.0),
                    vertical_align: Some(VerticalAlignSpec::Superscript),
                    ..RunSpec::default()
                },
            ],
            alignment: Some(ParagraphAlignmentSpec::Center),
            spacing_after_twips: Some(140),
            ..ParagraphSpec::default()
        });

        let runs: Vec<_> = paragraph.runs().collect();
        assert_eq!(paragraph.alignment(), Some(&ParagraphAlignment::Center));
        assert_eq!(paragraph.spacing_after_value(), Some(140));
        assert_eq!(runs.len(), 2);
        assert_eq!(
            runs[0].properties().font_family.as_deref(),
            Some("IBM Plex Sans")
        );
        assert_eq!(runs[0].properties().font_size, Some(24));
        assert!(runs[0].properties().bold);
        assert!(runs[1].properties().italic);
        assert_eq!(runs[1].properties().color.as_deref(), Some("C1121F"));
        assert_eq!(runs[1].properties().font_size, Some(28));
        assert_eq!(
            runs[1].properties().vertical_align,
            Some(VerticalAlign::Superscript)
        );
    }

    #[test]
    fn layout_document_keeps_visual_intro_group_together() {
        let mut config = RusdoxConfig::default();
        config.pdf.page_height_pt = 110.0;
        config.pdf.margin_top_pt = 10.0;
        config.pdf.margin_bottom_pt = 10.0;
        config.pdf.margin_x_pt = 10.0;
        let settings = PdfRenderSettings::from_config(&config);
        let mut font_system = pdf_font_system(&config);

        let mut document = Document::new();
        document.push_paragraph(
            Paragraph::new().add_run(Run::from_text("Line 1\nLine 2\nLine 3\nLine 4")),
        );
        document.push_paragraph(
            Paragraph::new()
                .keep_next()
                .add_run(Run::from_text("Benchmark Chart")),
        );
        document.push_visual(
            Visual::from_bytes(
                br##"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 40">
  <rect width="120" height="40" fill="#E2E8F0"/>
</svg>"##
                    .to_vec(),
                VisualFormat::Svg,
            )
            .with_kind(VisualKind::Chart)
            .width_twips(1_200),
        );

        let layout = layout_document(&document, settings, &mut font_system).expect("layout");

        assert_eq!(layout.pages.len(), 2);
        assert!(!page_line_texts(&layout.pages[0])
            .join("\n")
            .contains("Benchmark Chart"));
        assert!(page_line_texts(&layout.pages[1])
            .join("\n")
            .contains("Benchmark Chart"));
        assert!(layout.pages[1]
            .ops
            .iter()
            .any(|op| matches!(op, DrawOp::Image { .. })));
    }

    #[test]
    fn layout_paragraph_lines_breaks_on_newline_and_wrap() {
        let paragraph = Paragraph::new().add_run(Run::from_text(
            "verylongword verylongword verylongword\nnext line",
        ));
        let config = RusdoxConfig::default();
        let mut font_system = pdf_font_system(&config);
        let lines = layout_paragraph_lines(
            &Stylesheet::default(),
            &paragraph,
            120.0,
            default_pdf_settings(),
            &mut font_system,
        )
        .expect("layout lines");
        assert!(lines.len() >= 2);
        assert!(lines.iter().all(|line| line.width > 0.0));
    }

    #[test]
    fn layout_paragraph_lines_use_configured_line_height_multiplier() {
        let mut config = RusdoxConfig::default();
        config.pdf.default_line_height_pt = 12.0;
        config.pdf.line_height_multiplier = 1.6;
        let settings = PdfRenderSettings::from_config(&config);
        let mut font_system = pdf_font_system(&config);

        let paragraph = Paragraph::new().add_run(Run::from_text("Scaled").size_points(20));
        let lines = layout_paragraph_lines(
            &Stylesheet::default(),
            &paragraph,
            400.0,
            settings,
            &mut font_system,
        )
        .expect("layout lines");

        assert_eq!(lines.len(), 1);
        assert_close(lines[0].line_height, 32.0);
    }

    #[test]
    fn layout_row_uses_cell_width_when_present_and_fallback_when_absent() {
        let row = TableRow::new()
            .add_cell(
                TableCell::new()
                    .width(2_000)
                    .add_paragraph(Paragraph::new().add_run(Run::from_text("left"))),
            )
            .add_cell(
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("right"))),
            );
        let config = RusdoxConfig::default();
        let mut font_system = pdf_font_system(&config);
        let layout = layout_row(
            &row,
            &Stylesheet::default(),
            &[100.0, 300.0],
            default_pdf_settings(),
            &mut font_system,
        )
        .expect("layout row");
        assert_eq!(layout.cells.len(), 2);
        assert!(layout.cells[0].width < layout.cells[1].width);
        assert!(layout.height >= 24.0);
    }

    #[test]
    fn layout_row_uses_configured_pdf_padding() {
        let mut config = RusdoxConfig::default();
        config.table.pdf_cell_padding_x_pt = 12.0;
        config.table.pdf_cell_padding_y_pt = 10.0;
        let settings = PdfRenderSettings::from_config(&config);
        let mut font_system = pdf_font_system(&config);

        let row = TableRow::new().add_cell(
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("left"))),
        );
        let layout = layout_row(
            &row,
            &Stylesheet::default(),
            &[240.0],
            settings,
            &mut font_system,
        )
        .expect("layout row");

        assert_eq!(layout.cells.len(), 1);
        assert_close(layout.cells[0].lines[0].y_offset, 10.0);
        assert_close(
            layout.height,
            10.0 + settings.effective_line_height(settings.default_text_size) + 10.0,
        );
    }

    #[test]
    fn layout_row_resolves_grid_spans_against_table_grid() {
        let config = RusdoxConfig::default();
        let settings = PdfRenderSettings::from_config(&config);
        let mut font_system = pdf_font_system(&config);
        let table = Table::new()
            .width(6_000)
            .add_row(
                TableRow::new()
                    .add_cell(
                        TableCell::new()
                            .width(2_000)
                            .add_paragraph(Paragraph::new().add_run(Run::from_text("L"))),
                    )
                    .add_cell(
                        TableCell::new()
                            .width(4_000)
                            .add_paragraph(Paragraph::new().add_run(Run::from_text("R"))),
                    ),
            )
            .add_row(
                TableRow::new().add_cell(
                    TableCell::new()
                        .grid_span(2)
                        .add_paragraph(Paragraph::new().add_run(Run::from_text("Wide"))),
                ),
            );

        let column_widths = resolve_table_column_widths(&table, twips_to_points(6_000));
        let layout = layout_row(
            table.rows().nth(1).expect("second row"),
            &Stylesheet::default(),
            &column_widths,
            settings,
            &mut font_system,
        )
        .expect("layout row");

        assert_eq!(column_widths.len(), 2);
        assert_close(column_widths[0], 100.0);
        assert_close(column_widths[1], 200.0);
        assert_eq!(layout.cells.len(), 1);
        assert_close(layout.cells[0].width, 300.0);
    }

    #[test]
    fn layout_document_repeats_table_headers_across_pages() {
        let mut config = RusdoxConfig::default();
        config.pdf.page_width_pt = 240.0;
        config.pdf.page_height_pt = 60.0;
        config.pdf.margin_x_pt = 12.0;
        config.pdf.margin_top_pt = 5.0;
        config.pdf.margin_bottom_pt = 5.0;
        config.pdf.default_text_size_pt = 10.0;
        config.pdf.default_line_height_pt = 12.0;
        config.table.pdf_cell_padding_x_pt = 4.0;
        config.table.pdf_cell_padding_y_pt = 4.0;
        let settings = PdfRenderSettings::from_config(&config);
        let mut font_system = pdf_font_system(&config);
        let mut document = Document::new();
        document.push_table(
            Table::new()
                .add_row(
                    TableRow::new().repeat_as_header().add_cell(
                        TableCell::new()
                            .add_paragraph(Paragraph::new().add_run(Run::from_text("Header"))),
                    ),
                )
                .add_row(
                    TableRow::new().add_cell(
                        TableCell::new()
                            .add_paragraph(Paragraph::new().add_run(Run::from_text("Alpha"))),
                    ),
                )
                .add_row(
                    TableRow::new().add_cell(
                        TableCell::new()
                            .add_paragraph(Paragraph::new().add_run(Run::from_text("Beta"))),
                    ),
                ),
        );

        let pages = layout_document(&document, settings, &mut font_system).expect("layout");
        let page_texts = pages.pages.iter().map(page_line_texts).collect::<Vec<_>>();

        assert_eq!(pages.pages.len(), 2);
        assert_eq!(
            page_texts[0],
            vec!["Header".to_string(), "Alpha".to_string()]
        );
        assert_eq!(
            page_texts[1],
            vec!["Header".to_string(), "Beta".to_string()]
        );
    }

    #[test]
    fn layout_document_splits_tall_table_rows_across_pages() {
        let mut config = RusdoxConfig::default();
        config.pdf.page_width_pt = 220.0;
        config.pdf.page_height_pt = 38.0;
        config.pdf.margin_x_pt = 12.0;
        config.pdf.margin_top_pt = 4.0;
        config.pdf.margin_bottom_pt = 4.0;
        config.pdf.default_text_size_pt = 10.0;
        config.pdf.default_line_height_pt = 12.0;
        config.table.pdf_cell_padding_x_pt = 4.0;
        config.table.pdf_cell_padding_y_pt = 4.0;
        let settings = PdfRenderSettings::from_config(&config);
        let mut font_system = pdf_font_system(&config);
        let mut document = Document::new();
        document.push_table(Table::new().add_row(TableRow::new().add_cell(
            TableCell::new().add_paragraph(
                Paragraph::new().add_run(Run::from_text("Alpha\nBeta\nGamma\nDelta")),
            ),
        )));

        let pages = layout_document(&document, settings, &mut font_system).expect("layout");
        let page_texts = pages.pages.iter().map(page_line_texts).collect::<Vec<_>>();

        assert!(pages.pages.len() >= 2);
        assert_eq!(page_texts[0], vec!["Alpha".to_string()]);
        assert_eq!(page_texts[1], vec!["Beta".to_string()]);
        assert!(
            page_texts.iter().flatten().any(|line| line == "Delta"),
            "expected the split row to continue rendering later lines"
        );
    }

    #[test]
    fn save_named_writes_docx_and_skips_pdf_when_disabled() {
        let temp = tempdir().expect("temp dir");
        let mut config = RusdoxConfig::default();
        config.output.docx_dir = temp.path().join("docx").to_string_lossy().to_string();
        config.output.pdf_dir = temp.path().join("pdf").to_string_lossy().to_string();
        config.output.emit_pdf_preview = false;
        let studio = Studio::new(config);

        let mut document = Document::new();
        document.push_paragraph(Paragraph::new().add_run(Run::from_text("hello")));

        let stats = studio
            .save_named(&document, "report")
            .expect("save should succeed");

        let docx_path = PathBuf::from(studio.config().output.docx_dir.clone()).join("report.docx");
        let pdf_path = PathBuf::from(studio.config().output.pdf_dir.clone()).join("report.pdf");
        assert!(docx_path.exists());
        assert!(!pdf_path.exists());
        assert!(stats.docx_bytes > 0);
        assert_eq!(stats.pdf_bytes, 0);
    }

    #[test]
    fn save_with_pdf_stats_writes_both_artifacts_when_enabled() {
        let temp = tempdir().expect("temp dir");
        let mut config = RusdoxConfig::default();
        config.output.docx_dir = temp.path().join("docx").to_string_lossy().to_string();
        config.output.pdf_dir = temp.path().join("pdf").to_string_lossy().to_string();
        config.output.emit_pdf_preview = true;
        let studio = Studio::new(config);

        let mut document = Document::new();
        document.push_paragraph(Paragraph::new().add_run(Run::from_text("hello pdf")));
        let docx_path = temp.path().join("docx/manual.docx");

        let stats = studio
            .save_with_pdf_stats(&document, &docx_path)
            .expect("save should succeed");
        let pdf_path = temp.path().join("pdf/manual.pdf");

        assert!(docx_path.exists());
        assert!(pdf_path.exists());
        assert!(stats.docx_bytes > 0);
        assert!(stats.pdf_bytes > 0);
    }

    #[test]
    fn save_with_pdf_stats_uses_configured_page_size() {
        let temp = tempdir().expect("temp dir");
        let mut config = RusdoxConfig::default();
        config.output.docx_dir = temp.path().join("docx").to_string_lossy().to_string();
        config.output.pdf_dir = temp.path().join("pdf").to_string_lossy().to_string();
        config.output.emit_pdf_preview = true;
        config.pdf.page_width_pt = 300.0;
        config.pdf.page_height_pt = 500.0;
        let studio = Studio::new(config);

        let mut document = Document::new();
        document.push_paragraph(Paragraph::new().add_run(Run::from_text("custom pdf")));
        let docx_path = temp.path().join("docx/custom.docx");

        studio
            .save_with_pdf_stats(&document, &docx_path)
            .expect("save should succeed");

        let pdf_path = temp.path().join("pdf/custom.pdf");
        let pdf = fs::read(&pdf_path).expect("read pdf");
        let pdf_text = String::from_utf8_lossy(&pdf);

        assert!(pdf_text.contains("/MediaBox [0 0 300 500]"));
    }

    #[test]
    fn save_with_pdf_stats_uses_configured_text_origin_and_baseline() {
        let temp = tempdir().expect("temp dir");
        let mut config = RusdoxConfig::default();
        config.output.docx_dir = temp.path().join("docx").to_string_lossy().to_string();
        config.output.pdf_dir = temp.path().join("pdf").to_string_lossy().to_string();
        config.output.emit_pdf_preview = true;
        config.pdf.page_width_pt = 300.0;
        config.pdf.page_height_pt = 500.0;
        config.pdf.margin_x_pt = 33.0;
        config.pdf.margin_top_pt = 17.0;
        config.pdf.margin_bottom_pt = 12.0;
        config.pdf.default_text_size_pt = 10.0;
        config.pdf.default_line_height_pt = 20.0;
        config.pdf.baseline_factor = 0.5;
        let studio = Studio::new(config);

        let mut document = Document::new();
        document.push_paragraph(Paragraph::new().add_run(Run::from_text("origin")));
        let docx_path = temp.path().join("docx/origin.docx");

        studio
            .save_with_pdf_stats(&document, &docx_path)
            .expect("save should succeed");

        let pdf_path = temp.path().join("pdf/origin.pdf");
        let pdf = fs::read(&pdf_path).expect("read pdf");
        let pdf_text = String::from_utf8_lossy(&pdf);

        assert!(pdf_text.contains("1 0 0 1 33 473 Tm"));
    }

    #[test]
    fn save_with_pdf_stats_uses_configured_table_stroke_and_padding() {
        let temp = tempdir().expect("temp dir");
        let mut config = RusdoxConfig::default();
        config.output.docx_dir = temp.path().join("docx").to_string_lossy().to_string();
        config.output.pdf_dir = temp.path().join("pdf").to_string_lossy().to_string();
        config.output.emit_pdf_preview = true;
        config.pdf.page_width_pt = 240.0;
        config.pdf.page_height_pt = 240.0;
        config.pdf.margin_x_pt = 21.0;
        config.pdf.margin_top_pt = 11.0;
        config.pdf.default_text_size_pt = 10.0;
        config.pdf.default_line_height_pt = 20.0;
        config.pdf.baseline_factor = 0.5;
        config.table.pdf_cell_padding_x_pt = 13.0;
        config.table.pdf_cell_padding_y_pt = 4.0;
        config.table.pdf_grid_stroke_width_pt = 2.5;
        let studio = Studio::new(config);

        let mut document = Document::new();
        document.push_table(crate::Table::new().add_row(crate::TableRow::new().add_cell(
            crate::TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("cell"))),
        )));
        let docx_path = temp.path().join("docx/table.docx");

        studio
            .save_with_pdf_stats(&document, &docx_path)
            .expect("save should succeed");

        let pdf_path = temp.path().join("pdf/table.pdf");
        let pdf = fs::read(&pdf_path).expect("read pdf");
        let pdf_text = String::from_utf8_lossy(&pdf);

        assert!(pdf_text.contains("2.5 w"));
        assert!(pdf_text.contains("1 0 0 1 34 215 Tm"));
    }

    #[test]
    fn save_with_pdf_stats_embeds_truetype_fonts_and_unicode_maps() {
        let temp = tempdir().expect("temp dir");
        let mut config = RusdoxConfig::default();
        config.output.docx_dir = temp.path().join("docx").to_string_lossy().to_string();
        config.output.pdf_dir = temp.path().join("pdf").to_string_lossy().to_string();
        config.output.emit_pdf_preview = true;
        let studio = Studio::new(config);

        let mut document = Document::new();
        document.push_paragraph(Paragraph::new().add_run(Run::from_text("Привет مرحبا 你好")));
        let docx_path = temp.path().join("docx/unicode.docx");

        studio
            .save_with_pdf_stats(&document, &docx_path)
            .expect("save should succeed");

        let pdf_path = temp.path().join("pdf/unicode.pdf");
        let pdf = fs::read(&pdf_path).expect("read pdf");
        let pdf_text = String::from_utf8_lossy(&pdf);

        assert!(pdf_text.contains("/Subtype /Type0"));
        assert!(pdf_text.contains("/FontFile2"));
        assert!(pdf_text.contains("/ToUnicode"));
        assert!(pdf_text.contains("Identity-H"));
        assert!(pdf_text.contains(&format!("<{}>", utf16_hex("П"))));
        assert!(pdf_text.contains(&format!("<{}>", utf16_hex("م"))));
        assert!(pdf_text.contains(&format!("<{}>", utf16_hex("你"))));
    }

    #[test]
    fn save_with_pdf_stats_uses_font_fallback_for_multiscript_text() {
        assert_multiscript_pdf_output_with_optional_fallback(
            find_family_triggering_multiscript_fallback(),
        );
    }

    #[test]
    fn save_with_pdf_stats_handles_mixed_script_text_without_needing_fallback() {
        assert_multiscript_pdf_output_with_optional_fallback(None);
    }
}
