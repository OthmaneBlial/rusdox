use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};

use serde::Serialize;

use crate::config::RusdoxConfig;
use crate::spec::{
    BlockSpec, CellSpec, DocumentSpec, ParagraphSpec, RunSpec, TableSpec, VisualSpec,
};
use crate::{DocumentMetadata, ParagraphList, ParagraphListKind, Stylesheet, TableBorders};

const BUILTIN_PARAGRAPH_STYLE_ID: &str = "Normal";
const BUILTIN_RUN_STYLE_ID: &str = "DefaultParagraphFont";
const BUILTIN_TABLE_STYLE_ID: &str = "TableNormal";

/// Severity level for a validation issue.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum ValidationSeverity {
    Error,
    Warning,
}

/// A semantic validation issue discovered in a spec or config.
#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct ValidationIssue {
    pub severity: ValidationSeverity,
    pub path: String,
    pub message: String,
}

/// A collection of semantic validation issues.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize)]
pub struct ValidationReport {
    pub issues: Vec<ValidationIssue>,
}

impl ValidationReport {
    /// Returns whether the report contains at least one error.
    pub fn has_errors(&self) -> bool {
        self.issues
            .iter()
            .any(|issue| issue.severity == ValidationSeverity::Error)
    }

    /// Returns whether the report contains at least one warning.
    pub fn has_warnings(&self) -> bool {
        self.issues
            .iter()
            .any(|issue| issue.severity == ValidationSeverity::Warning)
    }

    /// Returns the number of validation errors.
    pub fn error_count(&self) -> usize {
        self.issues
            .iter()
            .filter(|issue| issue.severity == ValidationSeverity::Error)
            .count()
    }

    /// Returns the number of validation warnings.
    pub fn warning_count(&self) -> usize {
        self.issues
            .iter()
            .filter(|issue| issue.severity == ValidationSeverity::Warning)
            .count()
    }

    pub(crate) fn push_error(&mut self, path: impl Into<String>, message: impl Into<String>) {
        self.issues.push(ValidationIssue {
            severity: ValidationSeverity::Error,
            path: path.into(),
            message: message.into(),
        });
    }

    pub(crate) fn push_warning(&mut self, path: impl Into<String>, message: impl Into<String>) {
        self.issues.push(ValidationIssue {
            severity: ValidationSeverity::Warning,
            path: path.into(),
            message: message.into(),
        });
    }

    pub(crate) fn extend(&mut self, other: Self) {
        self.issues.extend(other.issues);
    }
}

/// Validates a loaded RusDox configuration for semantic problems.
pub fn validate_config(config: &RusdoxConfig) -> ValidationReport {
    let mut report = ValidationReport::default();

    validate_positive_f32(
        &mut report,
        "typography.cover_title_size_pt",
        config.typography.cover_title_size_pt,
    );
    validate_positive_f32(
        &mut report,
        "typography.title_size_pt",
        config.typography.title_size_pt,
    );
    validate_positive_f32(
        &mut report,
        "typography.subtitle_size_pt",
        config.typography.subtitle_size_pt,
    );
    validate_positive_f32(
        &mut report,
        "typography.hero_size_pt",
        config.typography.hero_size_pt,
    );
    validate_positive_f32(
        &mut report,
        "typography.page_heading_size_pt",
        config.typography.page_heading_size_pt,
    );
    validate_positive_f32(
        &mut report,
        "typography.section_size_pt",
        config.typography.section_size_pt,
    );
    validate_positive_f32(
        &mut report,
        "typography.body_size_pt",
        config.typography.body_size_pt,
    );
    validate_positive_f32(
        &mut report,
        "typography.tagline_size_pt",
        config.typography.tagline_size_pt,
    );
    validate_positive_f32(
        &mut report,
        "typography.note_size_pt",
        config.typography.note_size_pt,
    );
    validate_positive_f32(
        &mut report,
        "typography.table_size_pt",
        config.typography.table_size_pt,
    );
    validate_positive_f32(
        &mut report,
        "typography.metric_label_size_pt",
        config.typography.metric_label_size_pt,
    );
    validate_positive_f32(
        &mut report,
        "typography.metric_value_size_pt",
        config.typography.metric_value_size_pt,
    );

    validate_hex_color(&mut report, "colors.ink", &config.colors.ink);
    validate_hex_color(&mut report, "colors.slate", &config.colors.slate);
    validate_hex_color(&mut report, "colors.muted", &config.colors.muted);
    validate_hex_color(&mut report, "colors.accent", &config.colors.accent);
    validate_hex_color(&mut report, "colors.gold", &config.colors.gold);
    validate_hex_color(&mut report, "colors.red", &config.colors.red);
    validate_hex_color(&mut report, "colors.green", &config.colors.green);
    validate_hex_color(&mut report, "colors.soft", &config.colors.soft);
    validate_hex_color(&mut report, "colors.pale", &config.colors.pale);
    validate_hex_color(&mut report, "colors.mint", &config.colors.mint);
    validate_hex_color(&mut report, "colors.amber", &config.colors.amber);
    validate_hex_color(&mut report, "colors.rose", &config.colors.rose);
    validate_hex_color(
        &mut report,
        "colors.table_border",
        &config.colors.table_border,
    );

    validate_positive_u32(
        &mut report,
        "table.default_width_twips",
        config.table.default_width_twips,
    );
    validate_positive_u32(
        &mut report,
        "table.metric_cell_width_twips",
        config.table.metric_cell_width_twips,
    );
    validate_positive_u32(
        &mut report,
        "table.grid_border_size_eighth_pt",
        config.table.grid_border_size_eighth_pt,
    );
    validate_positive_u32(
        &mut report,
        "table.card_border_size_eighth_pt",
        config.table.card_border_size_eighth_pt,
    );
    validate_non_negative_f32(
        &mut report,
        "table.pdf_cell_padding_x_pt",
        config.table.pdf_cell_padding_x_pt,
    );
    validate_non_negative_f32(
        &mut report,
        "table.pdf_cell_padding_y_pt",
        config.table.pdf_cell_padding_y_pt,
    );
    validate_non_negative_f32(
        &mut report,
        "table.pdf_after_spacing_pt",
        config.table.pdf_after_spacing_pt,
    );
    validate_positive_f32(
        &mut report,
        "table.pdf_grid_stroke_width_pt",
        config.table.pdf_grid_stroke_width_pt,
    );

    validate_positive_f32(&mut report, "pdf.page_width_pt", config.pdf.page_width_pt);
    validate_positive_f32(&mut report, "pdf.page_height_pt", config.pdf.page_height_pt);
    validate_non_negative_f32(&mut report, "pdf.margin_x_pt", config.pdf.margin_x_pt);
    validate_non_negative_f32(&mut report, "pdf.margin_top_pt", config.pdf.margin_top_pt);
    validate_non_negative_f32(
        &mut report,
        "pdf.margin_bottom_pt",
        config.pdf.margin_bottom_pt,
    );
    validate_positive_f32(
        &mut report,
        "pdf.default_text_size_pt",
        config.pdf.default_text_size_pt,
    );
    validate_positive_f32(
        &mut report,
        "pdf.default_line_height_pt",
        config.pdf.default_line_height_pt,
    );
    validate_positive_f32(
        &mut report,
        "pdf.line_height_multiplier",
        config.pdf.line_height_multiplier,
    );
    validate_positive_f32(
        &mut report,
        "pdf.baseline_factor",
        config.pdf.baseline_factor,
    );
    validate_positive_f32(
        &mut report,
        "pdf.text_width_bias_regular",
        config.pdf.text_width_bias_regular,
    );
    validate_positive_f32(
        &mut report,
        "pdf.text_width_bias_bold",
        config.pdf.text_width_bias_bold,
    );

    if config.pdf.page_width_pt.is_finite()
        && config.pdf.margin_x_pt.is_finite()
        && (config.pdf.page_width_pt - (config.pdf.margin_x_pt * 2.0)) <= 0.0
    {
        report.push_error(
            "pdf.margin_x_pt",
            "horizontal page margins leave no remaining content width",
        );
    }
    if config.pdf.page_height_pt.is_finite()
        && config.pdf.margin_top_pt.is_finite()
        && config.pdf.margin_bottom_pt.is_finite()
        && (config.pdf.page_height_pt - config.pdf.margin_top_pt - config.pdf.margin_bottom_pt)
            <= 0.0
    {
        report.push_error(
            "pdf.margin_top_pt",
            "vertical page margins leave no remaining content height",
        );
    }

    report
}

/// Validates a loaded document spec for semantic problems.
pub fn validate_spec(spec: &DocumentSpec) -> ValidationReport {
    let mut report = ValidationReport::default();
    let mut list_registry = BTreeMap::new();

    if let Some(output_name) = spec.output_name.as_ref() {
        if output_name.trim().is_empty() {
            report.push_error("output_name", "output name cannot be blank");
        }
    }
    validate_metadata(&mut report, &spec.metadata);

    validate_stylesheet(&mut report, &mut list_registry, &spec.styles);

    for (index, block) in spec.blocks.iter().enumerate() {
        let block_path = format!("blocks[{index}]");
        validate_block(
            &mut report,
            &block_path,
            block,
            spec,
            &spec.styles,
            &mut list_registry,
        );
    }

    report
}

/// Validates a document spec and the runtime configuration it will use.
pub fn validate_spec_with_config(spec: &DocumentSpec, config: &RusdoxConfig) -> ValidationReport {
    let mut report = validate_config(config);
    report.extend(validate_spec(spec));
    report
}

fn validate_stylesheet(
    report: &mut ValidationReport,
    list_registry: &mut BTreeMap<u32, (ParagraphListKind, String)>,
    stylesheet: &Stylesheet,
) {
    let paragraph_ids = collect_style_ids(
        report,
        "styles.paragraph",
        stylesheet.paragraph.iter().map(|style| style.id.as_str()),
    );
    let run_ids = collect_style_ids(
        report,
        "styles.run",
        stylesheet.run.iter().map(|style| style.id.as_str()),
    );
    let table_ids = collect_style_ids(
        report,
        "styles.table",
        stylesheet.table.iter().map(|style| style.id.as_str()),
    );

    for (index, style) in stylesheet.paragraph.iter().enumerate() {
        let path = format!("styles.paragraph[{index}]");
        if let Some(parent) = style.based_on.as_deref() {
            if !is_known_paragraph_style(parent, &paragraph_ids) {
                report.push_error(
                    format!("{path}.based_on"),
                    format!("unknown paragraph style '{parent}'"),
                );
            }
        }
        if let Some(next) = style.next.as_deref() {
            if !is_known_paragraph_style(next, &paragraph_ids) {
                report.push_error(
                    format!("{path}.next"),
                    format!("unknown paragraph style '{next}'"),
                );
            }
        }
        if let Some(list) = style.paragraph.list {
            validate_list(
                report,
                list_registry,
                format!("{path}.paragraph.list"),
                list,
            );
        }
        validate_run_style_properties(report, &format!("{path}.run"), &style.run);
    }

    for (index, style) in stylesheet.run.iter().enumerate() {
        let path = format!("styles.run[{index}]");
        if let Some(parent) = style.based_on.as_deref() {
            if !is_known_run_style(parent, &run_ids) {
                report.push_error(
                    format!("{path}.based_on"),
                    format!("unknown run style '{parent}'"),
                );
            }
        }
        validate_run_style_properties(report, &format!("{path}.properties"), &style.properties);
    }

    for (index, style) in stylesheet.table.iter().enumerate() {
        let path = format!("styles.table[{index}]");
        if let Some(parent) = style.based_on.as_deref() {
            if !is_known_table_style(parent, &table_ids) {
                report.push_error(
                    format!("{path}.based_on"),
                    format!("unknown table style '{parent}'"),
                );
            }
        }
        validate_table_borders(
            report,
            &format!("{path}.properties.borders"),
            style.properties.borders.as_ref(),
        );
        if let Some(width) = style.properties.width {
            validate_positive_u32(report, format!("{path}.properties.width"), width);
        }
    }

    for style in stylesheet.paragraph_styles() {
        if let Err(error) = stylesheet.resolve_paragraph_style(Some(style.id.as_str())) {
            report.push_error(format!("styles.paragraph.{}", style.id), error.to_string());
        }
    }
    for style in stylesheet.run_styles() {
        if let Err(error) = stylesheet.resolve_run_style(Some(style.id.as_str())) {
            report.push_error(format!("styles.run.{}", style.id), error.to_string());
        }
    }
    for style in stylesheet.table_styles() {
        if let Err(error) = stylesheet.resolve_table_style(Some(style.id.as_str())) {
            report.push_error(format!("styles.table.{}", style.id), error.to_string());
        }
    }
}

fn validate_metadata(report: &mut ValidationReport, metadata: &DocumentMetadata) {
    if metadata
        .title
        .as_deref()
        .is_some_and(|value| value.trim().is_empty())
    {
        report.push_error("metadata.title", "metadata title cannot be blank");
    }
    if metadata
        .author
        .as_deref()
        .is_some_and(|value| value.trim().is_empty())
    {
        report.push_error("metadata.author", "metadata author cannot be blank");
    }
    if metadata
        .subject
        .as_deref()
        .is_some_and(|value| value.trim().is_empty())
    {
        report.push_error("metadata.subject", "metadata subject cannot be blank");
    }

    for (index, keyword) in metadata.keywords.iter().enumerate() {
        if keyword.trim().is_empty() {
            report.push_error(
                format!("metadata.keywords[{index}]"),
                "metadata keyword cannot be blank",
            );
        }
    }

    for name in metadata.custom_properties.keys() {
        if name.trim().is_empty() {
            report.push_error(
                "metadata.custom_properties",
                "custom property names cannot be blank",
            );
        }
    }
}

fn collect_style_ids<'a>(
    report: &mut ValidationReport,
    base_path: &str,
    ids: impl Iterator<Item = &'a str>,
) -> BTreeMap<String, usize> {
    let mut seen = BTreeMap::new();
    for (index, id) in ids.enumerate() {
        let path = format!("{base_path}[{index}].id");
        if id.trim().is_empty() {
            report.push_error(path, "style id cannot be blank");
            continue;
        }
        if let Some(previous) = seen.insert(id.to_string(), index) {
            report.push_error(
                path,
                format!("duplicate style id '{id}' also defined at index {previous}"),
            );
        }
    }
    seen
}

fn validate_block(
    report: &mut ValidationReport,
    path: &str,
    block: &BlockSpec,
    spec: &DocumentSpec,
    stylesheet: &Stylesheet,
    list_registry: &mut BTreeMap<u32, (ParagraphListKind, String)>,
) {
    match block {
        BlockSpec::Paragraph { spec: paragraph } => {
            validate_paragraph_block(report, path, paragraph, stylesheet, list_registry);
        }
        BlockSpec::Table { spec: table } => validate_table_block(report, path, table, stylesheet),
        BlockSpec::Image { spec: visual }
        | BlockSpec::Logo { spec: visual }
        | BlockSpec::Signature { spec: visual }
        | BlockSpec::Chart { spec: visual } => validate_visual_block(report, path, visual, spec),
        BlockSpec::CoverTitle { text }
        | BlockSpec::Title { text }
        | BlockSpec::Subtitle { text }
        | BlockSpec::Hero { text }
        | BlockSpec::CenteredNote { text }
        | BlockSpec::PageHeading { text }
        | BlockSpec::Section { text }
        | BlockSpec::Body { text }
        | BlockSpec::Tagline { text } => {
            if text.trim().is_empty() {
                report.push_warning(format!("{path}.text"), "text is blank");
            }
        }
        BlockSpec::Bullets { items } | BlockSpec::Numbered { items } => {
            if items.is_empty() {
                report.push_warning(format!("{path}.items"), "list has no items");
            }
        }
        BlockSpec::LabelValues { items } => {
            if items.is_empty() {
                report.push_warning(format!("{path}.items"), "label/value block has no items");
            }
        }
        BlockSpec::Metrics { items } => {
            if items.is_empty() {
                report.push_warning(format!("{path}.items"), "metrics block has no items");
            }
        }
        BlockSpec::Spacer => {}
    }
}

fn validate_paragraph_block(
    report: &mut ValidationReport,
    path: &str,
    paragraph: &ParagraphSpec,
    stylesheet: &Stylesheet,
    list_registry: &mut BTreeMap<u32, (ParagraphListKind, String)>,
) {
    if let Some(style_id) = paragraph.style_id.as_deref() {
        validate_style_id_reference(
            report,
            format!("{path}.spec.style_id"),
            style_id,
            is_known_paragraph_style(
                style_id,
                &collect_known_ids(stylesheet.paragraph_styles().map(|style| style.id.as_str())),
            ),
            "paragraph",
        );
        if let Some(style) = stylesheet.paragraph_style(style_id) {
            if let Some(list) = style.paragraph.list {
                validate_list(report, list_registry, format!("{path}.spec.style_id"), list);
            }
        }
    }

    if paragraph.runs.is_empty() {
        report.push_warning(format!("{path}.spec.runs"), "paragraph has no runs");
    }

    for (index, run) in paragraph.runs.iter().enumerate() {
        validate_run_spec(
            report,
            &format!("{path}.spec.runs[{index}]"),
            run,
            stylesheet,
        );
    }
}

fn validate_run_spec(
    report: &mut ValidationReport,
    path: &str,
    run: &RunSpec,
    stylesheet: &Stylesheet,
) {
    if let Some(style_id) = run.style_id.as_deref() {
        validate_style_id_reference(
            report,
            format!("{path}.style_id"),
            style_id,
            is_known_run_style(
                style_id,
                &collect_known_ids(stylesheet.run_styles().map(|style| style.id.as_str())),
            ),
            "run",
        );
    }
    if let Some(color) = run.color.as_deref() {
        validate_hex_color(report, format!("{path}.color"), color);
    }
    if let Some(font_family) = run.font_family.as_deref() {
        if font_family.trim().is_empty() {
            report.push_error(format!("{path}.font_family"), "font family cannot be blank");
        }
    }
    if let Some(size_pt) = run.size_pt {
        validate_positive_f32(report, format!("{path}.size_pt"), size_pt);
    }
}

fn validate_table_block(
    report: &mut ValidationReport,
    path: &str,
    table: &TableSpec,
    stylesheet: &Stylesheet,
) {
    if let Some(style_id) = table.style_id.as_deref() {
        validate_style_id_reference(
            report,
            format!("{path}.spec.style_id"),
            style_id,
            is_known_table_style(
                style_id,
                &collect_known_ids(stylesheet.table_styles().map(|style| style.id.as_str())),
            ),
            "table",
        );
    }

    if table.columns.is_empty() {
        report.push_error(
            format!("{path}.spec.columns"),
            "table must define at least one column",
        );
    }

    for (index, column) in table.columns.iter().enumerate() {
        if column.label.trim().is_empty() {
            report.push_warning(
                format!("{path}.spec.columns[{index}].label"),
                "column label is blank",
            );
        }
        validate_positive_u32(
            report,
            format!("{path}.spec.columns[{index}].width"),
            column.width,
        );
    }

    for (row_index, row) in table.rows.iter().enumerate() {
        let cell_count = row.cells.len();
        let column_count = table.columns.len();
        if cell_count > column_count {
            report.push_error(
                format!("{path}.spec.rows[{row_index}].cells"),
                format!(
                    "row has {cell_count} cells but the table only defines {column_count} columns"
                ),
            );
        } else if cell_count < column_count {
            report.push_warning(
                format!("{path}.spec.rows[{row_index}].cells"),
                format!(
                    "row has {cell_count} cells but the table defines {column_count} columns; blank cells will be inserted"
                ),
            );
        }

        for (cell_index, cell) in row.cells.iter().enumerate() {
            match cell {
                CellSpec::Text { text } => {
                    if text.trim().is_empty() {
                        report.push_warning(
                            format!("{path}.spec.rows[{row_index}].cells[{cell_index}].text"),
                            "cell text is blank",
                        );
                    }
                }
                CellSpec::Status(status) => {
                    if status.text.trim().is_empty() {
                        report.push_warning(
                            format!("{path}.spec.rows[{row_index}].cells[{cell_index}].text"),
                            "status cell text is blank",
                        );
                    }
                }
            }
        }
    }
}

fn validate_visual_block(
    report: &mut ValidationReport,
    path: &str,
    visual: &VisualSpec,
    spec: &DocumentSpec,
) {
    if visual.path.trim().is_empty() {
        report.push_error(format!("{path}.path"), "visual path cannot be blank");
        return;
    }

    let resolved = resolve_visual_path(spec, &visual.path);
    match supported_visual_extension(&resolved) {
        Some(_) => {}
        None => report.push_error(
            format!("{path}.path"),
            format!(
                "unsupported visual format '{}'; expected png, jpg, jpeg, or svg",
                resolved
                    .extension()
                    .and_then(|ext| ext.to_str())
                    .unwrap_or_default()
            ),
        ),
    }

    match fs::metadata(&resolved) {
        Ok(metadata) => {
            if !metadata.is_file() {
                report.push_error(
                    format!("{path}.path"),
                    format!("visual asset is not a file: {}", resolved.display()),
                );
            }
        }
        Err(_) => report.push_error(
            format!("{path}.path"),
            format!("visual asset does not exist: {}", resolved.display()),
        ),
    }

    if let Some(width) = visual.width_twips {
        validate_positive_u32(report, format!("{path}.width_twips"), width);
    }
    if let Some(height) = visual.height_twips {
        validate_positive_u32(report, format!("{path}.height_twips"), height);
    }
    if let Some(width) = visual.max_width_twips {
        validate_positive_u32(report, format!("{path}.max_width_twips"), width);
    }
    if let Some(height) = visual.max_height_twips {
        validate_positive_u32(report, format!("{path}.max_height_twips"), height);
    }
}

fn resolve_visual_path(spec: &DocumentSpec, raw_path: &str) -> PathBuf {
    let mut path = PathBuf::from(raw_path);
    if path.is_relative() {
        if let Some(base_dir) = spec.asset_base_dir() {
            path = base_dir.join(path);
        }
    }
    path
}

fn supported_visual_extension(path: &Path) -> Option<&'static str> {
    match path
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase()
        .as_str()
    {
        "png" => Some("png"),
        "jpg" | "jpeg" => Some("jpeg"),
        "svg" => Some("svg"),
        _ => None,
    }
}

fn validate_run_style_properties(
    report: &mut ValidationReport,
    path: &str,
    properties: &crate::RunStyleProperties,
) {
    if let Some(color) = properties.color.as_deref() {
        validate_hex_color(report, format!("{path}.color"), color);
    }
    if let Some(font_family) = properties.font_family.as_deref() {
        if font_family.trim().is_empty() {
            report.push_error(format!("{path}.font_family"), "font family cannot be blank");
        }
    }
    if let Some(font_size) = properties.font_size {
        if font_size == 0 {
            report.push_error(
                format!("{path}.font_size"),
                "font size must be greater than zero",
            );
        }
    }
}

fn validate_table_borders(
    report: &mut ValidationReport,
    path: &str,
    borders: Option<&TableBorders>,
) {
    let Some(borders) = borders else {
        return;
    };

    validate_border_color(report, format!("{path}.top.color"), borders.top.as_ref());
    validate_border_color(
        report,
        format!("{path}.bottom.color"),
        borders.bottom.as_ref(),
    );
    validate_border_color(report, format!("{path}.left.color"), borders.left.as_ref());
    validate_border_color(
        report,
        format!("{path}.right.color"),
        borders.right.as_ref(),
    );
    validate_border_color(
        report,
        format!("{path}.inside_horizontal.color"),
        borders.inside_horizontal.as_ref(),
    );
    validate_border_color(
        report,
        format!("{path}.inside_vertical.color"),
        borders.inside_vertical.as_ref(),
    );
}

fn validate_border_color(
    report: &mut ValidationReport,
    path: impl Into<String>,
    border: Option<&crate::Border>,
) {
    if let Some(color) = border.and_then(|border| border.color.as_deref()) {
        validate_hex_color(report, path, color);
    }
}

fn validate_list(
    report: &mut ValidationReport,
    registry: &mut BTreeMap<u32, (ParagraphListKind, String)>,
    path: impl Into<String>,
    list: ParagraphList,
) {
    let path = path.into();
    if list.id() == 0 {
        report.push_error(format!("{path}.id"), "list id must be greater than zero");
    }
    if list.level() > 8 {
        report.push_error(
            format!("{path}.level"),
            "list level must be between 0 and 8",
        );
    }
    if let Some((existing_kind, existing_path)) = registry.get(&list.id()) {
        if *existing_kind != list.kind() {
            report.push_error(
                path,
                format!(
                    "list id {} is reused with conflicting kinds; first seen at {existing_path}",
                    list.id()
                ),
            );
        }
    } else {
        registry.insert(list.id(), (list.kind(), path));
    }
}

fn validate_style_id_reference(
    report: &mut ValidationReport,
    path: impl Into<String>,
    style_id: &str,
    known: bool,
    style_kind: &str,
) {
    let path = path.into();
    if style_id.trim().is_empty() {
        report.push_error(path, format!("{style_kind} style id cannot be blank"));
    } else if !known {
        report.push_error(path, format!("unknown {style_kind} style '{style_id}'"));
    }
}

fn collect_known_ids<'a>(ids: impl Iterator<Item = &'a str>) -> BTreeMap<String, usize> {
    ids.enumerate()
        .map(|(index, id)| (id.to_string(), index))
        .collect()
}

fn is_known_paragraph_style(style_id: &str, known_ids: &BTreeMap<String, usize>) -> bool {
    style_id == BUILTIN_PARAGRAPH_STYLE_ID || known_ids.contains_key(style_id)
}

fn is_known_run_style(style_id: &str, known_ids: &BTreeMap<String, usize>) -> bool {
    style_id == BUILTIN_RUN_STYLE_ID || known_ids.contains_key(style_id)
}

fn is_known_table_style(style_id: &str, known_ids: &BTreeMap<String, usize>) -> bool {
    style_id == BUILTIN_TABLE_STYLE_ID || known_ids.contains_key(style_id)
}

fn validate_positive_u32(report: &mut ValidationReport, path: impl Into<String>, value: u32) {
    if value == 0 {
        report.push_error(path, "value must be greater than zero");
    }
}

fn validate_positive_f32(report: &mut ValidationReport, path: impl Into<String>, value: f32) {
    if !value.is_finite() || value <= 0.0 {
        report.push_error(path, "value must be finite and greater than zero");
    }
}

fn validate_non_negative_f32(report: &mut ValidationReport, path: impl Into<String>, value: f32) {
    if !value.is_finite() || value < 0.0 {
        report.push_error(path, "value must be finite and zero or greater");
    }
}

fn validate_hex_color(report: &mut ValidationReport, path: impl Into<String>, value: &str) {
    let normalized = value.trim();
    if normalized.len() != 6 || !normalized.chars().all(|ch| ch.is_ascii_hexdigit()) {
        report.push_error(
            path,
            format!("invalid color '{value}', expected six hex digits without '#'"),
        );
    }
}

#[cfg(test)]
mod tests {
    use std::fs;

    use tempfile::tempdir;

    use super::{validate_config, validate_spec, ValidationSeverity};
    use crate::config::RusdoxConfig;
    use crate::spec::{
        BlockSpec, CellSpec, ColumnSpec, DocumentSpec, ParagraphSpec, RowSpec, RunSpec, TableSpec,
        VisualSpec,
    };
    use crate::{
        ParagraphAlignment, ParagraphList, ParagraphStyle, ParagraphStyleProperties, RunStyle,
        RunStyleProperties, Stylesheet, TableStyle, TableStyleProperties,
    };

    #[test]
    fn validate_spec_accepts_well_formed_named_style_document() {
        let mut spec = DocumentSpec::new();
        spec.output_name = Some("ok".to_string());
        spec.styles = Stylesheet::new()
            .add_paragraph_style(
                ParagraphStyle::new("lead")
                    .based_on("Normal")
                    .paragraph(
                        ParagraphStyleProperties::new()
                            .alignment(ParagraphAlignment::Center)
                            .spacing_after(120),
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
                    .properties(TableStyleProperties::new().width(9_360)),
            );
        spec.blocks = vec![
            BlockSpec::Paragraph {
                spec: ParagraphSpec {
                    style_id: Some("lead".to_string()),
                    runs: vec![RunSpec {
                        text: "Hello".to_string(),
                        style_id: Some("accent".to_string()),
                        ..RunSpec::default()
                    }],
                    ..ParagraphSpec::default()
                },
            },
            BlockSpec::Table {
                spec: TableSpec {
                    style_id: Some("grid".to_string()),
                    columns: vec![ColumnSpec {
                        label: "Metric".to_string(),
                        width: 4_000,
                    }],
                    rows: vec![RowSpec {
                        cells: vec![CellSpec::Text {
                            text: "ARR".to_string(),
                        }],
                    }],
                },
            },
        ];

        let report = validate_spec(&spec);
        assert!(!report.has_errors(), "{report:?}");
        assert!(!report.has_warnings(), "{report:?}");
    }

    #[test]
    fn validate_spec_reports_style_table_and_visual_errors() {
        let temp = tempdir().expect("temp dir");
        let mut spec = DocumentSpec::new().with_asset_base_dir(temp.path());
        spec.metadata.title = Some(" ".to_string());
        spec.metadata.keywords = vec!["".to_string()];
        spec.styles = Stylesheet::new()
            .add_paragraph_style(ParagraphStyle::new("loop").based_on("loop").paragraph(
                ParagraphStyleProperties {
                    list: Some(ParagraphList::numbered_with_id(3)),
                    ..ParagraphStyleProperties::default()
                },
            ))
            .add_run_style(
                RunStyle::new("accent")
                    .based_on("missing")
                    .properties(RunStyleProperties::new().color("#AA5500")),
            );
        spec.blocks = vec![
            BlockSpec::Paragraph {
                spec: ParagraphSpec {
                    style_id: Some("missing".to_string()),
                    runs: vec![RunSpec {
                        text: "Styled".to_string(),
                        style_id: Some("accent".to_string()),
                        color: Some("XYZ123".to_string()),
                        ..RunSpec::default()
                    }],
                    ..ParagraphSpec::default()
                },
            },
            BlockSpec::Table {
                spec: TableSpec {
                    style_id: None,
                    columns: vec![ColumnSpec {
                        label: "Only".to_string(),
                        width: 0,
                    }],
                    rows: vec![RowSpec {
                        cells: vec![
                            CellSpec::Text {
                                text: "A".to_string(),
                            },
                            CellSpec::Text {
                                text: "B".to_string(),
                            },
                        ],
                    }],
                },
            },
            BlockSpec::Image {
                spec: VisualSpec {
                    path: "missing.gif".to_string(),
                    max_width_twips: Some(0),
                    ..VisualSpec::default()
                },
            },
        ];

        let report = validate_spec(&spec);
        assert!(report.has_errors());
        assert!(report.issues.iter().any(|issue| {
            issue.path == "styles.paragraph.loop" && issue.severity == ValidationSeverity::Error
        }));
        assert!(report
            .issues
            .iter()
            .any(|issue| issue.path == "metadata.title"));
        assert!(report
            .issues
            .iter()
            .any(|issue| issue.path == "metadata.keywords[0]"));
        assert!(report
            .issues
            .iter()
            .any(|issue| issue.message.contains("unknown paragraph style 'missing'")));
        assert!(report
            .issues
            .iter()
            .any(|issue| issue.message.contains("unsupported visual format")));
        assert!(report.issues.iter().any(|issue| issue
            .message
            .contains("row has 2 cells but the table only defines 1 columns")));
        assert!(report
            .issues
            .iter()
            .any(|issue| issue.message.contains("invalid color 'XYZ123'")
                || issue.message.contains("invalid color '#AA5500'")));
    }

    #[test]
    fn validate_config_reports_invalid_colors_and_pdf_geometry() {
        let mut config = RusdoxConfig::default();
        config.colors.ink = "#0F172A".to_string();
        config.typography.body_size_pt = 0.0;
        config.pdf.page_width_pt = 100.0;
        config.pdf.margin_x_pt = 60.0;

        let report = validate_config(&config);
        assert!(report.has_errors());
        assert!(report
            .issues
            .iter()
            .any(|issue| issue.path == "colors.ink" && issue.message.contains("invalid color")));
        assert!(report
            .issues
            .iter()
            .any(|issue| issue.path == "typography.body_size_pt"));
        assert!(report
            .issues
            .iter()
            .any(|issue| issue.message.contains("no remaining content width")));
    }

    #[test]
    fn validate_visual_paths_relative_to_spec_base_dir() {
        let temp = tempdir().expect("temp dir");
        let asset_path = temp.path().join("logo.svg");
        fs::write(&asset_path, "<svg xmlns=\"http://www.w3.org/2000/svg\"/>").expect("write asset");

        let mut spec = DocumentSpec::new().with_asset_base_dir(temp.path());
        spec.blocks = vec![BlockSpec::Logo {
            spec: VisualSpec::new("logo.svg"),
        }];

        let report = validate_spec(&spec);
        assert!(!report.has_errors(), "{report:?}");
    }
}
