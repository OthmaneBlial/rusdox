//! Configurable document composition and PDF preview helpers.
//!
//! This module is intended to be the core orchestration layer for projects
//! that build rich `.docx` files programmatically and optionally emit PDF
//! previews.

use std::fmt::Write as _;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::OnceLock;
use std::time::{Duration, Instant};

use pdf_writer::{Name, Pdf, Rect, Ref};

use crate::{
    config::RusdoxConfig,
    spec::{
        BlockSpec, CellSpec, DocumentSpec, ParagraphAlignmentSpec,
        ParagraphSpec as BlockParagraphSpec, RunSpec as BlockRunSpec, TableSpec as BlockTableSpec,
        Tone, UnderlineStyleSpec, VerticalAlignSpec,
    },
    Border, BorderStyle, Document, DocumentBlockRef, Paragraph, ParagraphAlignment, Run, Table,
    TableBorders, TableCell, TableRow, UnderlineStyle, VerticalAlign,
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

    /// Renders a high-level document specification into a [`Document`].
    pub fn compose(&self, spec: &DocumentSpec) -> Document {
        let mut document = Document::new();
        self.append_spec(&mut document, spec);
        document
    }

    /// Appends a high-level document specification to an existing [`Document`].
    pub fn append_spec(&self, document: &mut Document, spec: &DocumentSpec) {
        for block in &spec.blocks {
            self.push_spec_block(document, block);
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

    /// Writes DOCX and optional PDF output and returns timing stats.
    pub fn save_with_pdf_stats(
        &self,
        document: &Document,
        docx_path: impl AsRef<Path>,
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
            println!("{}", docx_path.display());
            println!("{}", pdf_path.display());
        } else {
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

    /// Builds a bullet paragraph.
    pub fn bullet(&self, text: &str) -> Paragraph {
        Paragraph::new()
            .spacing_after(self.config.spacing.bullet_after_twips)
            .add_run(
                Run::from_text("• ")
                    .font(self.config.typography.font_family.clone())
                    .size_points(points_to_u16(self.config.typography.body_size_pt))
                    .bold()
                    .color(&self.config.colors.accent),
            )
            .add_run(
                Run::from_text(text)
                    .font(self.config.typography.font_family.clone())
                    .size_points(points_to_u16(self.config.typography.body_size_pt))
                    .color(&self.config.colors.slate),
            )
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

    fn push_spec_block(&self, document: &mut Document, block: &BlockSpec) {
        match block {
            BlockSpec::CoverTitle { text } => {
                document.push_paragraph(self.cover_title(text));
            }
            BlockSpec::Title { text } => {
                document.push_paragraph(self.title(text));
            }
            BlockSpec::Subtitle { text } => {
                document.push_paragraph(self.subtitle(text));
            }
            BlockSpec::Hero { text } => {
                document.push_paragraph(self.hero(text));
            }
            BlockSpec::CenteredNote { text } => {
                document.push_paragraph(self.centered_note(text));
            }
            BlockSpec::PageHeading { text } => {
                document.push_paragraph(self.page_heading(text));
            }
            BlockSpec::Section { text } => {
                document.push_paragraph(self.section(text));
            }
            BlockSpec::Body { text } => {
                document.push_paragraph(self.body(text));
            }
            BlockSpec::Tagline { text } => {
                document.push_paragraph(self.tagline(text));
            }
            BlockSpec::Paragraph { spec } => {
                document.push_paragraph(self.paragraph_from_spec(spec));
            }
            BlockSpec::Bullets { items } => {
                for item in items {
                    document.push_paragraph(self.bullet(item));
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
            BlockSpec::Spacer => {
                document.push_paragraph(self.spacer());
            }
        }
    }

    fn paragraph_from_spec(&self, spec: &BlockParagraphSpec) -> Paragraph {
        let mut paragraph = Paragraph::new();

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
        let mut table = Table::new()
            .width(if total_width == 0 {
                self.config.table.default_width_twips
            } else {
                total_width
            })
            .borders(self.grid_borders())
            .add_row(spec.columns.iter().fold(TableRow::new(), |row, column| {
                row.add_cell(self.header_cell(&column.label, column.width))
            }));

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

    fn tone_foreground(&self, tone: Tone) -> &str {
        match tone {
            Tone::Positive => &self.config.colors.green,
            Tone::Risk => &self.config.colors.red,
            Tone::Neutral | Tone::Warning => &self.config.colors.ink,
        }
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

#[derive(Clone, Copy, PartialEq, Eq)]
enum PdfFont {
    Regular,
    Bold,
    Oblique,
    BoldOblique,
}

impl PdfFont {
    fn resource_name(self) -> Name<'static> {
        match self {
            Self::Regular => Name(b"F1"),
            Self::Bold => Name(b"F2"),
            Self::Oblique => Name(b"F3"),
            Self::BoldOblique => Name(b"F4"),
        }
    }

    fn base_font_name(self) -> Name<'static> {
        match self {
            Self::Regular => Name(b"Helvetica"),
            Self::Bold => Name(b"Helvetica-Bold"),
            Self::Oblique => Name(b"Helvetica-Oblique"),
            Self::BoldOblique => Name(b"Helvetica-BoldOblique"),
        }
    }
}

#[derive(Clone, Copy)]
struct TextStyle {
    font: PdfFont,
    size: f32,
    color: Rgb,
}

#[derive(Clone)]
struct TextSpan {
    text: String,
    style: TextStyle,
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
}

struct PdfLayout {
    pages: Vec<Page>,
    cursor_y: f32,
    settings: PdfRenderSettings,
}

impl PdfLayout {
    fn new(settings: PdfRenderSettings) -> Self {
        Self {
            pages: vec![Page::default()],
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

fn render_pdf(document: &Document, pdf_path: &Path, config: &RusdoxConfig) -> crate::Result<()> {
    let settings = PdfRenderSettings::from_config(config);
    let pages = layout_document(document, settings);
    let catalog_id = Ref::new(1);
    let page_tree_id = Ref::new(2);
    let mut next_id = 3;
    let mut page_ids = Vec::with_capacity(pages.len());
    let mut content_ids = Vec::with_capacity(pages.len());

    for _ in &pages {
        page_ids.push(Ref::new(next_id));
        next_id += 1;
        content_ids.push(Ref::new(next_id));
        next_id += 1;
    }

    let estimated_capacity = pages
        .iter()
        .map(|page| page.ops.len() * 96)
        .sum::<usize>()
        .max(64 * 1024);
    let mut pdf = Pdf::with_capacity(estimated_capacity);
    pdf.catalog(catalog_id).pages(page_tree_id);
    pdf.pages(page_tree_id)
        .kids(page_ids.iter().copied())
        .count(i32::try_from(page_ids.len()).unwrap_or(i32::MAX));

    for ((page, page_id), content_id) in pages.iter().zip(&page_ids).zip(&content_ids) {
        let content = render_page_content(page, settings);
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
        let mut fonts = resources.fonts();
        for font in [
            PdfFont::Regular,
            PdfFont::Bold,
            PdfFont::Oblique,
            PdfFont::BoldOblique,
        ] {
            fonts
                .insert(font.resource_name())
                .start::<pdf_writer::writers::Type1Font>()
                .base_font(font.base_font_name());
        }
    }

    fs::write(pdf_path, pdf.finish())?;
    Ok(())
}

fn layout_document(document: &Document, settings: PdfRenderSettings) -> Vec<Page> {
    let mut layout = PdfLayout::new(settings);

    for block in document.blocks() {
        match block {
            DocumentBlockRef::Paragraph(paragraph) => {
                layout_paragraph_block(&mut layout, paragraph)
            }
            DocumentBlockRef::Table(table) => layout_table_block(&mut layout, table),
        }
    }

    layout.pages
}

fn layout_paragraph_block(layout: &mut PdfLayout, paragraph: &Paragraph) {
    if paragraph.has_page_break_before() && layout.cursor_y > layout.settings.margin_top {
        layout.push_page();
    }

    layout.cursor_y += twips_to_points(paragraph.spacing_before_value().unwrap_or(0));
    let lines = layout_paragraph_lines(paragraph, layout.settings.content_width, layout.settings);

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

    layout.cursor_y += twips_to_points(paragraph.spacing_after_value().unwrap_or(0));
}

fn layout_table_block(layout: &mut PdfLayout, table: &Table) {
    let total_width = table
        .properties()
        .width
        .map(twips_to_points)
        .unwrap_or(layout.settings.content_width)
        .min(layout.settings.content_width);

    for row in table.rows() {
        let row_layout = layout_row(row, total_width, layout.settings);
        layout.ensure_space(row_layout.height);
        let y_top = layout.cursor_y;

        for cell in row_layout.cells {
            layout.push_op(DrawOp::Rect {
                x: layout.settings.margin_x + cell.x_offset,
                y_top,
                width: cell.width,
                height: row_layout.height,
                fill: cell.background,
                stroke: Some((
                    layout.settings.table_grid_stroke_color,
                    layout.settings.table_grid_stroke_width,
                )),
            });

            for line in cell.lines {
                layout.push_op(DrawOp::TextLine {
                    x: layout.settings.margin_x
                        + cell.x_offset
                        + layout.settings.table_cell_padding_x,
                    y_top: y_top + line.y_offset,
                    line: line.layout,
                    max_width: cell.width - (layout.settings.table_cell_padding_x * 2.0),
                });
            }
        }

        layout.cursor_y += row_layout.height;
    }

    layout.cursor_y += layout.settings.table_after_spacing;
}

struct RowLayout {
    cells: Vec<CellLayout>,
    height: f32,
}

struct CellLayout {
    x_offset: f32,
    width: f32,
    background: Option<Rgb>,
    lines: Vec<CellLine>,
}

struct CellLine {
    y_offset: f32,
    layout: LineLayout,
}

fn layout_row(row: &TableRow, total_width: f32, settings: PdfRenderSettings) -> RowLayout {
    let cell_count = row.cells().len().max(1) as f32;
    let fallback_width = total_width / cell_count;
    let mut x_offset = 0.0;
    let mut cells = Vec::new();
    let mut row_height: f32 = 0.0;

    for cell in row.cells() {
        let width = cell
            .properties()
            .width
            .map(twips_to_points)
            .unwrap_or(fallback_width);
        let content_width = (width - (settings.table_cell_padding_x * 2.0)).max(MIN_CONTENT_WIDTH);
        let mut lines = Vec::new();
        let mut y_offset = settings.table_cell_padding_y;

        for paragraph in cell.paragraphs() {
            y_offset += twips_to_points(paragraph.spacing_before_value().unwrap_or(0));
            let paragraph_lines = layout_paragraph_lines(paragraph, content_width, settings);
            if paragraph_lines.is_empty() {
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
            y_offset += twips_to_points(paragraph.spacing_after_value().unwrap_or(0));
        }

        let height = y_offset + settings.table_cell_padding_y;
        row_height = row_height.max(height);
        cells.push(CellLayout {
            x_offset,
            width,
            background: cell
                .properties()
                .background_color
                .as_deref()
                .and_then(parse_hex_color),
            lines,
        });
        x_offset += width;
    }

    RowLayout {
        cells,
        height: row_height.max(MIN_CONTENT_WIDTH),
    }
}

fn layout_paragraph_lines(
    paragraph: &Paragraph,
    max_width: f32,
    settings: PdfRenderSettings,
) -> Vec<LineLayout> {
    let alignment = paragraph
        .alignment()
        .cloned()
        .unwrap_or(ParagraphAlignment::Left);
    let mut lines = Vec::new();
    let mut current_spans = Vec::new();
    let mut current_width = 0.0;
    let mut current_line_height = settings.default_line_height;

    for run in paragraph.runs() {
        let style = style_from_run(run, settings);
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

            let sanitized = sanitize_text(&token);
            if sanitized.is_empty() {
                continue;
            }
            let width = estimate_text_width(style.font, style.size, &sanitized, settings);

            if current_width + width > max_width
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

            if sanitized.trim().is_empty() && current_spans.is_empty() {
                continue;
            }

            current_line_height =
                current_line_height.max(settings.effective_line_height(style.size));
            current_width += width;
            current_spans.push(TextSpan {
                text: sanitized,
                style,
                width,
            });
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

    lines
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

fn style_from_run(run: &Run, settings: PdfRenderSettings) -> TextStyle {
    let properties = run.properties();
    let font = match (properties.bold, properties.italic) {
        (true, true) => PdfFont::BoldOblique,
        (true, false) => PdfFont::Bold,
        (false, true) => PdfFont::Oblique,
        (false, false) => PdfFont::Regular,
    };

    TextStyle {
        font,
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

fn sanitize_text(text: &str) -> String {
    text.chars()
        .map(|ch| match ch {
            '•' => '-',
            '\u{00A0}' => ' ',
            ch if (ch as u32) <= 0x00FF => ch,
            _ => '?',
        })
        .collect()
}

fn estimate_text_width(font: PdfFont, size: f32, text: &str, settings: PdfRenderSettings) -> f32 {
    let font_bias = match font {
        PdfFont::Regular | PdfFont::Oblique => settings.text_width_bias_regular,
        PdfFont::Bold | PdfFont::BoldOblique => settings.text_width_bias_bold,
    };

    let units = text
        .chars()
        .map(|ch| match ch {
            'i' | 'j' | 'l' | '!' | '\'' | ',' | '.' | ':' | ';' | '|' => 0.28,
            'f' | 'r' | 't' | '(' | ')' | '[' | ']' => 0.34,
            ' ' => 0.28,
            'm' | 'w' | 'M' | 'W' | '@' | '#' | '%' => 0.88,
            'A'..='Z' => 0.66,
            '0'..='9' => 0.56,
            _ => 0.54,
        })
        .sum::<f32>();

    units * size * font_bias
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

    fn draw_line(&mut self, x: f32, y_top: f32, line: &LineLayout, max_width: f32) {
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
            let font_key = (span.style.font, span.style.size.to_bits());
            if current_font != Some(font_key) {
                self.buffer.push(b'/');
                self.buffer
                    .extend_from_slice(span.style.font.resource_name().0);
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
            self.buffer.extend_from_slice(b" Tm\n(");
            self.push_pdf_text(&span.text);
            self.buffer.extend_from_slice(b") Tj\n");

            cursor_x += span.width;
        }
        self.buffer.extend_from_slice(b"ET\n");
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

    fn push_pdf_text(&mut self, text: &str) {
        for byte in text.bytes() {
            match byte {
                b'(' | b')' | b'\\' => {
                    self.buffer.push(b'\\');
                    self.buffer.push(byte);
                }
                b'\r' => {}
                b'\n' => self.buffer.extend_from_slice(br"\n"),
                _ => self.buffer.push(byte),
            }
        }
    }
}

fn render_page_content(page: &Page, settings: PdfRenderSettings) -> Vec<u8> {
    let mut writer = PdfContentWriter::with_capacity(page.ops.len() * 96, settings);

    for op in &page.ops {
        match op {
            DrawOp::TextLine {
                x,
                y_top,
                line,
                max_width,
            } => writer.draw_line(*x, *y_top, line, *max_width),
            DrawOp::Rect {
                x,
                y_top,
                width,
                height,
                fill,
                stroke,
            } => writer.draw_rect(*x, *y_top, *width, *height, *fill, *stroke),
        }
    }

    writer.finish()
}

#[cfg(test)]
mod tests {
    use std::{fs, path::PathBuf};

    use tempfile::tempdir;

    use super::{
        clamp_u32_to_u16, estimate_text_width, layout_paragraph_lines, layout_row, parse_hex_color,
        points_to_u16, sanitize_text, style_from_run, tokenize, PdfFont, PdfRenderSettings, Rgb,
        Studio, DEFAULT_BASELINE_FACTOR, DEFAULT_LINE_HEIGHT, DEFAULT_LINE_HEIGHT_MULTIPLIER,
        DEFAULT_TABLE_AFTER_SPACING, DEFAULT_TABLE_GRID_STROKE_WIDTH, DEFAULT_TABLE_ROW_PADDING_X,
        DEFAULT_TABLE_ROW_PADDING_Y, DEFAULT_TEXT_SIZE, DEFAULT_TEXT_WIDTH_BIAS_BOLD,
        DEFAULT_TEXT_WIDTH_BIAS_REGULAR, MIN_CONTENT_WIDTH,
    };
    use crate::config::RusdoxConfig;
    use crate::spec::{
        ParagraphAlignmentSpec, ParagraphSpec, RunSpec, UnderlineStyleSpec, VerticalAlignSpec,
    };
    use crate::{Document, Paragraph, ParagraphAlignment, Run, TableCell, TableRow, VerticalAlign};

    fn default_pdf_settings() -> PdfRenderSettings {
        PdfRenderSettings::from_config(&RusdoxConfig::default())
    }

    fn assert_close(actual: f32, expected: f32) {
        assert!(
            (actual - expected).abs() < 0.01,
            "expected {expected:.2}, got {actual:.2}"
        );
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
    fn sanitize_text_normalizes_special_chars() {
        assert_eq!(sanitize_text("•\u{00A0}ok"), "- ok");
        assert_eq!(sanitize_text("é"), "é");
        assert_eq!(sanitize_text("🤖"), "?");
    }

    #[test]
    fn tokenize_splits_words_spaces_tabs_and_newlines() {
        let tokens = tokenize("a\tb  c\nd");
        assert_eq!(tokens, vec!["a", " ", "b", "  ", "c", "\n", "d"]);
    }

    #[test]
    fn estimate_text_width_reflects_font_bias() {
        let regular = estimate_text_width(PdfFont::Regular, 12.0, "Hello", default_pdf_settings());
        let bold = estimate_text_width(PdfFont::Bold, 12.0, "Hello", default_pdf_settings());
        assert!(regular > 0.0);
        assert!(bold > regular);
    }

    #[test]
    fn style_from_run_maps_font_size_and_color_and_weight() {
        let run = Run::from_text("x")
            .bold()
            .italic()
            .size_points(18)
            .color("AABBCC");
        let style = style_from_run(&run, default_pdf_settings());
        assert!(style.font == PdfFont::BoldOblique);
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
    fn estimate_text_width_uses_configured_bias_values() {
        let mut config = RusdoxConfig::default();
        config.pdf.text_width_bias_regular = 0.5;
        config.pdf.text_width_bias_bold = 1.5;
        let settings = PdfRenderSettings::from_config(&config);

        let regular = estimate_text_width(PdfFont::Regular, 10.0, "Hello", settings);
        let bold = estimate_text_width(PdfFont::Bold, 10.0, "Hello", settings);

        assert!(regular > 0.0);
        assert_close(bold / regular, 3.0);
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
    fn layout_paragraph_lines_breaks_on_newline_and_wrap() {
        let paragraph = Paragraph::new().add_run(Run::from_text(
            "verylongword verylongword verylongword\nnext line",
        ));
        let lines = layout_paragraph_lines(&paragraph, 120.0, default_pdf_settings());
        assert!(lines.len() >= 2);
        assert!(lines.iter().all(|line| line.width > 0.0));
    }

    #[test]
    fn layout_paragraph_lines_use_configured_line_height_multiplier() {
        let mut config = RusdoxConfig::default();
        config.pdf.default_line_height_pt = 12.0;
        config.pdf.line_height_multiplier = 1.6;
        let settings = PdfRenderSettings::from_config(&config);

        let paragraph = Paragraph::new().add_run(Run::from_text("Scaled").size_points(20));
        let lines = layout_paragraph_lines(&paragraph, 400.0, settings);

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
        let layout = layout_row(&row, 400.0, default_pdf_settings());
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

        let row = TableRow::new().add_cell(
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("left"))),
        );
        let layout = layout_row(&row, 240.0, settings);

        assert_eq!(layout.cells.len(), 1);
        assert_close(layout.cells[0].lines[0].y_offset, 10.0);
        assert_close(
            layout.height,
            10.0 + settings.effective_line_height(settings.default_text_size) + 10.0,
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
}
