//! Machine-readable and human-readable evidence for DOCX/PDF parity checks.
//!
//! The parity contract compares the typed source document with the DOCX that
//! RusDox writes and with the semantic projection consumed by the native PDF
//! renderer. It deliberately reports semantic and layout evidence separately:
//! a green semantic check does not pretend to prove viewer-perfect pixels.

use std::fs;
use std::path::{Path, PathBuf};

use image::GenericImageView;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

use crate::{
    Document, DocumentBlockRef, DocumentMetadata, HeaderFooter, PageNumbering, PageSetup, Result,
};

/// Current version of the parity report contract.
pub const PARITY_REPORT_VERSION: &str = "1";

/// A normalized, serializable view of document semantics.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct DocumentProjection {
    /// Normalized text.
    pub normalized_text: String,
    /// Ordered block order.
    pub block_order: Vec<String>,
    /// Ordered headings.
    pub headings: Vec<String>,
    /// Ordered tables.
    pub tables: Vec<Vec<Vec<String>>>,
    /// Ordered images.
    pub images: Vec<ParityImage>,
    /// Ordered page breaks.
    pub page_breaks: Vec<usize>,
    /// Ordered section breaks.
    pub section_breaks: Vec<usize>,
    /// Ordered hyperlinks.
    pub hyperlinks: Vec<ParityHyperlink>,
    /// Ordered bookmarks.
    pub bookmarks: Vec<String>,
    /// Ordered fields.
    pub fields: Vec<String>,
    /// Ordered footnotes.
    pub footnotes: Vec<String>,
    /// Ordered table layout.
    pub table_layout: Vec<ParityTableLayout>,
    /// Metadata.
    pub metadata: DocumentMetadata,
    /// Page setup.
    pub page_setup: PageSetup,
    /// Optional header.
    pub header: Option<HeaderFooter>,
    /// Optional footer.
    pub footer: Option<HeaderFooter>,
    /// Optional page numbering.
    pub page_numbering: Option<PageNumbering>,
}

impl DocumentProjection {
    /// Builds the semantic projection used by the parity contract.
    pub fn from_document(document: &Document) -> Self {
        let mut text = Vec::new();
        let mut block_order = Vec::new();
        let mut headings = Vec::new();
        let mut tables = Vec::new();
        let mut images = Vec::new();
        let mut page_breaks = Vec::new();
        let mut section_breaks = Vec::new();
        let mut hyperlinks = Vec::new();
        let mut bookmarks = Vec::new();
        let mut fields = Vec::new();
        let mut footnotes = Vec::new();
        let mut table_layout = Vec::new();

        for (index, block) in document.blocks().enumerate() {
            match block {
                DocumentBlockRef::Paragraph(paragraph) => {
                    block_order.push("paragraph".to_string());
                    let normalized = normalize_text(&paragraph.text());
                    if !normalized.is_empty() {
                        text.push(normalized.clone());
                    }
                    if paragraph
                        .style_id()
                        .is_some_and(is_heading_style_identifier)
                        && !normalized.is_empty()
                    {
                        headings.push(normalized);
                    }
                    if paragraph.has_page_break_before() {
                        page_breaks.push(index);
                    }
                    if paragraph.has_section_break_before() {
                        section_breaks.push(index);
                    }
                    collect_run_semantics(
                        paragraph,
                        &mut hyperlinks,
                        &mut bookmarks,
                        &mut fields,
                        &mut footnotes,
                    );
                }
                DocumentBlockRef::Table(table) => {
                    block_order.push("table".to_string());
                    let rows = table
                        .rows()
                        .map(|row| {
                            row.cells()
                                .map(|cell| {
                                    let normalized = normalize_text(&cell.text());
                                    if !normalized.is_empty() {
                                        text.push(normalized.clone());
                                    }
                                    normalized
                                })
                                .collect::<Vec<_>>()
                        })
                        .collect::<Vec<_>>();
                    tables.push(rows);
                    collect_table_semantics(
                        table,
                        &mut hyperlinks,
                        &mut bookmarks,
                        &mut fields,
                        &mut footnotes,
                        &mut table_layout,
                    );
                }
                DocumentBlockRef::Visual(visual) => {
                    block_order.push("visual".to_string());
                    images.push(ParityImage {
                        kind: format!("{:?}", visual.kind()).to_ascii_lowercase(),
                        alt_text: visual.alt_text().map(normalize_text),
                    });
                }
            }
        }

        Self {
            normalized_text: text.join("\n"),
            block_order,
            headings,
            tables,
            images,
            page_breaks,
            section_breaks,
            hyperlinks,
            bookmarks,
            fields,
            footnotes,
            table_layout,
            metadata: normalized_metadata(document.metadata()),
            page_setup: document.page_setup().clone(),
            header: document.header().cloned(),
            footer: document.footer().cloned(),
            page_numbering: document.page_numbering().cloned(),
        }
    }
}

fn collect_run_semantics(
    paragraph: &crate::Paragraph,
    hyperlinks: &mut Vec<ParityHyperlink>,
    bookmarks: &mut Vec<String>,
    fields: &mut Vec<String>,
    footnotes: &mut Vec<String>,
) {
    for run in paragraph.runs() {
        if let Some(target) = run.hyperlink_target() {
            hyperlinks.push(ParityHyperlink {
                text: normalize_text(run.text()),
                target: target.to_string(),
            });
        }
        if let Some(bookmark) = run.bookmark_name() {
            bookmarks.push(bookmark.to_string());
        }
        if let Some(field) = run.field_kind() {
            fields.push(format!("{field:?}").to_ascii_lowercase());
        }
        if let Some(footnote) = run.footnote_text() {
            footnotes.push(normalize_text(footnote));
        }
    }
}

fn collect_table_semantics(
    table: &crate::Table,
    hyperlinks: &mut Vec<ParityHyperlink>,
    bookmarks: &mut Vec<String>,
    fields: &mut Vec<String>,
    footnotes: &mut Vec<String>,
    table_layout: &mut Vec<ParityTableLayout>,
) {
    let mut rows = Vec::new();
    for row in table.rows() {
        let mut cells = Vec::new();
        for cell in row.cells() {
            for paragraph in cell.paragraphs() {
                collect_run_semantics(paragraph, hyperlinks, bookmarks, fields, footnotes);
            }
            for nested in cell.nested_tables() {
                collect_table_semantics(
                    nested,
                    hyperlinks,
                    bookmarks,
                    fields,
                    footnotes,
                    table_layout,
                );
            }
            cells.push(ParityTableCellLayout {
                grid_span: cell.properties().grid_span.unwrap_or(1),
                paragraph_count: cell
                    .paragraphs()
                    .filter(|paragraph| !normalize_text(&paragraph.text()).is_empty())
                    .count(),
                nested_table_count: cell.nested_tables().count(),
            });
        }
        rows.push(ParityTableRowLayout {
            repeat_as_header: row.properties().repeat_as_header,
            allow_split_across_pages: row.properties().allow_split_across_pages,
            cells,
        });
    }
    table_layout.push(ParityTableLayout { rows });
}

/// Hyperlink semantics independent of container formatting.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ParityHyperlink {
    /// Text.
    pub text: String,
    /// Target.
    pub target: String,
}

/// Row and cell layout controls compared after DOCX round-trips.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ParityTableLayout {
    /// Ordered rows.
    pub rows: Vec<ParityTableRowLayout>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
/// Describes parity table row layout.
pub struct ParityTableRowLayout {
    /// Whether the row repeats as a header on later pages.
    pub repeat_as_header: bool,
    /// Whether the row may split across pages.
    pub allow_split_across_pages: bool,
    /// Ordered cells.
    pub cells: Vec<ParityTableCellLayout>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
/// Describes parity table cell layout.
pub struct ParityTableCellLayout {
    /// Grid span.
    pub grid_span: u32,
    /// Number of paragraph values.
    pub paragraph_count: usize,
    /// Number of nested table values.
    pub nested_table_count: usize,
}

/// Image semantics that can be compared across output paths.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ParityImage {
    /// Kind.
    pub kind: String,
    /// Optional alt text.
    pub alt_text: Option<String>,
}

/// Evidence returned by the native PDF renderer.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct PdfRenderEvidence {
    /// Projection.
    pub projection: DocumentProjection,
    /// Number of page values.
    pub page_count: usize,
    /// Draw operations.
    pub draw_operations: usize,
    /// Text lines.
    pub text_lines: usize,
    /// Image operations.
    pub image_operations: usize,
}

/// Result state for one parity assertion.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum CheckStatus {
    /// Selects the passed form.
    Passed,
    /// Selects the failed form.
    Failed,
    /// Selects the skipped form.
    Skipped,
}

/// One independently named assertion in a parity report.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ParityCheck {
    /// Identifier.
    pub id: String,
    /// Label.
    pub label: String,
    /// Status.
    pub status: CheckStatus,
    /// Detail.
    pub detail: String,
}

/// Hash and size evidence for one generated artifact.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ArtifactEvidence {
    /// Path.
    pub path: String,
    /// Bytes.
    pub bytes: u64,
    /// SHA-256.
    pub sha256: String,
}

impl ArtifactEvidence {
    /// Reads an artifact and captures its byte size and SHA-256 digest.
    pub fn from_path(path: &Path) -> Result<Self> {
        let bytes = fs::read(path)?;
        let digest = Sha256::digest(&bytes);
        Ok(Self {
            path: path.display().to_string(),
            bytes: bytes.len() as u64,
            sha256: format!("{digest:x}"),
        })
    }
}

/// Per-page result from an optional deterministic visual-layout comparison.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct VisualPageDiff {
    /// Page.
    pub page: usize,
    /// Current.
    pub current: String,
    /// Optional baseline.
    pub baseline: Option<String>,
    /// Optional different pixel ratio.
    pub different_pixel_ratio: Option<f64>,
    /// Whether this page stayed within the configured threshold.
    pub passed: bool,
    /// Detail.
    pub detail: String,
}

/// Summary of optional visual-layout comparisons.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct VisualDiffSummary {
    /// Whether rendered-page comparison was enabled.
    pub enabled: bool,
    /// Threshold.
    pub threshold: f64,
    /// Whether every compared page stayed within the threshold.
    pub passed: bool,
    /// Ordered pages.
    pub pages: Vec<VisualPageDiff>,
}

impl VisualDiffSummary {
    /// Skipped.
    pub fn skipped(threshold: f64) -> Self {
        Self {
            enabled: false,
            threshold,
            passed: true,
            pages: Vec::new(),
        }
    }
}

/// Complete JSON/HTML parity report.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ParityReport {
    /// Report version.
    pub report_version: String,
    /// Generator version.
    pub generator_version: String,
    /// Source.
    pub source: String,
    /// Whether every required parity check passed.
    pub passed: bool,
    /// Ordered checks.
    pub checks: Vec<ParityCheck>,
    /// Expected.
    pub expected: DocumentProjection,
    /// DOCX.
    pub docx: DocumentProjection,
    /// PDF.
    pub pdf: PdfRenderEvidence,
    /// Visual diff.
    pub visual_diff: VisualDiffSummary,
    /// Ordered artifacts.
    pub artifacts: Vec<ArtifactEvidence>,
}

impl ParityReport {
    /// Compares the source model, reopened DOCX, and native PDF projection.
    #[allow(clippy::too_many_arguments)]
    pub fn compare(
        source: impl Into<String>,
        expected: DocumentProjection,
        docx: DocumentProjection,
        pdf: PdfRenderEvidence,
        visual_diff: VisualDiffSummary,
        artifacts: Vec<ArtifactEvidence>,
        docx_package_valid: bool,
        pdf_file_valid: bool,
    ) -> Self {
        let mut checks = vec![
            equality_check(
                "normalized_text",
                "Normalized text",
                &expected.normalized_text,
                &docx.normalized_text,
                &pdf.projection.normalized_text,
            ),
            equality_check(
                "block_order",
                "Block order",
                &expected.block_order,
                &docx.block_order,
                &pdf.projection.block_order,
            ),
            equality_check(
                "headings",
                "Heading sequence",
                &expected.headings,
                &docx.headings,
                &pdf.projection.headings,
            ),
            equality_check(
                "table_content",
                "Table rows and cells",
                &expected.tables,
                &docx.tables,
                &pdf.projection.tables,
            ),
            equality_check(
                "images",
                "Image count and alt text",
                &expected.images,
                &docx.images,
                &pdf.projection.images,
            ),
            equality_check(
                "page_breaks",
                "Explicit page breaks",
                &expected.page_breaks,
                &docx.page_breaks,
                &pdf.projection.page_breaks,
            ),
            equality_check(
                "section_breaks",
                "Explicit section breaks",
                &expected.section_breaks,
                &docx.section_breaks,
                &pdf.projection.section_breaks,
            ),
            equality_check(
                "hyperlinks",
                "Hyperlink text and targets",
                &expected.hyperlinks,
                &docx.hyperlinks,
                &pdf.projection.hyperlinks,
            ),
            equality_check(
                "bookmarks",
                "Bookmark anchors",
                &expected.bookmarks,
                &docx.bookmarks,
                &pdf.projection.bookmarks,
            ),
            equality_check(
                "fields",
                "Dynamic fields",
                &expected.fields,
                &docx.fields,
                &pdf.projection.fields,
            ),
            equality_check(
                "footnotes",
                "Footnote text",
                &expected.footnotes,
                &docx.footnotes,
                &pdf.projection.footnotes,
            ),
            equality_check(
                "table_layout",
                "Table row and rich-cell controls",
                &expected.table_layout,
                &docx.table_layout,
                &pdf.projection.table_layout,
            ),
            equality_check(
                "metadata",
                "Document metadata",
                &expected.metadata,
                &docx.metadata,
                &pdf.projection.metadata,
            ),
            equality_check(
                "page_setup",
                "Page setup",
                &expected.page_setup,
                &docx.page_setup,
                &pdf.projection.page_setup,
            ),
            equality_check(
                "headers_footers",
                "Headers and footers",
                &(expected.header.clone(), expected.footer.clone()),
                &(docx.header.clone(), docx.footer.clone()),
                &(pdf.projection.header.clone(), pdf.projection.footer.clone()),
            ),
            equality_check(
                "page_numbering",
                "Page-number fields",
                &expected.page_numbering,
                &docx.page_numbering,
                &pdf.projection.page_numbering,
            ),
            boolean_check(
                "docx_package",
                "DOCX package integrity",
                docx_package_valid,
                "required OOXML parts and relationships are present",
            ),
            boolean_check(
                "pdf_file",
                "PDF file integrity",
                pdf_file_valid,
                "PDF header, trailer, pages, and renderer evidence are present",
            ),
        ];

        checks.push(if visual_diff.enabled {
            boolean_check(
                "visual_diff",
                "Rendered-page visual threshold",
                visual_diff.passed,
                &format!(
                    "{} page(s), allowed different-pixel ratio {:.4}",
                    visual_diff.pages.len(),
                    visual_diff.threshold
                ),
            )
        } else {
            ParityCheck {
                id: "visual_diff".to_string(),
                label: "Rendered-page visual threshold".to_string(),
                status: CheckStatus::Skipped,
                detail: "no baseline supplied; deterministic page snapshots were still generated"
                    .to_string(),
            }
        });

        let passed = checks
            .iter()
            .all(|check| check.status != CheckStatus::Failed);
        Self {
            report_version: PARITY_REPORT_VERSION.to_string(),
            generator_version: env!("CARGO_PKG_VERSION").to_string(),
            source: source.into(),
            passed,
            checks,
            expected,
            docx,
            pdf,
            visual_diff,
            artifacts,
        }
    }

    /// Serializes this report as stable, pretty JSON.
    pub fn to_json_pretty(&self) -> Result<String> {
        serde_json::to_string_pretty(self).map_err(|error| {
            crate::DocxError::Parse(format!("failed to serialize parity report: {error}"))
        })
    }

    /// Renders a standalone, no-JavaScript HTML parity report.
    pub fn to_html(&self, canonical_url: &str) -> String {
        let status = if self.passed { "PASS" } else { "FAIL" };
        let status_class = if self.passed { "pass" } else { "fail" };
        let playground_id = Path::new(&self.source)
            .file_stem()
            .and_then(|stem| stem.to_str())
            .unwrap_or(&self.source)
            .replace('_', "-");
        let playground_url = format!(
            "https://othmaneblial.github.io/rusdox/playground/?example={}",
            escape_html(&playground_id)
        );
        let checks = self
            .checks
            .iter()
            .map(|check| {
                let (class, label) = match check.status {
                    CheckStatus::Passed => ("pass", "PASS"),
                    CheckStatus::Failed => ("fail", "FAIL"),
                    CheckStatus::Skipped => ("skip", "SKIP"),
                };
                format!(
                    "<tr><td><code>{}</code></td><td>{}</td><td><span class=\"badge {}\">{}</span></td><td>{}</td></tr>",
                    escape_html(&check.id),
                    escape_html(&check.label),
                    class,
                    label,
                    escape_html(&check.detail),
                )
            })
            .collect::<String>();
        let artifacts = self
            .artifacts
            .iter()
            .map(|artifact| {
                format!(
                    "<li><strong>{}</strong><span>{} bytes</span><code>{}</code></li>",
                    escape_html(&artifact.path),
                    artifact.bytes,
                    escape_html(&artifact.sha256),
                )
            })
            .collect::<String>();
        let pages = self
            .visual_diff
            .pages
            .iter()
            .map(|page| {
                let current_path = Path::new(&page.current);
                let current_href = current_path
                    .parent()
                    .and_then(Path::file_name)
                    .and_then(|parent| {
                        current_path.file_name().map(|file| {
                            format!(
                                "{}/{}",
                                parent.to_string_lossy(),
                                file.to_string_lossy()
                            )
                        })
                    })
                    .unwrap_or_default();
                format!(
                    "<li><figure><img src=\"{}\" alt=\"Deterministic layout snapshot for page {}\" loading=\"lazy\" /><figcaption>Page {} — {}{}</figcaption></figure></li>",
                    escape_html(&current_href),
                    page.page,
                    page.page,
                    escape_html(&page.detail),
                    page.different_pixel_ratio
                        .map(|ratio| format!(" ({ratio:.6} different pixels)"))
                        .unwrap_or_default(),
                )
            })
            .collect::<String>();

        format!(
            r#"<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{source} parity report | RusDox</title>
  <meta name="description" content="RusDox semantic and visual parity evidence for {source}." />
  <link rel="canonical" href="{canonical}" />
  <style>
    :root {{ color-scheme: light; --ink:#17211b; --paper:#f6f1e7; --line:#d7cbb9; --green:#17633a; --red:#9d2f2f; --amber:#8a5a13; }}
    * {{ box-sizing:border-box }} body {{ margin:0; font:15px/1.55 ui-sans-serif,system-ui,sans-serif; color:var(--ink); background:var(--paper) }}
    main {{ width:min(1100px,calc(100% - 32px)); margin:40px auto 80px }} header {{ display:flex; gap:24px; justify-content:space-between; align-items:end; border-bottom:2px solid var(--ink); padding-bottom:20px }}
    h1 {{ margin:.2rem 0; font:700 clamp(2rem,6vw,4.5rem)/.95 Georgia,serif }} h2 {{ margin-top:2.5rem }} p {{ max-width:75ch }}
    .eyebrow {{ margin:0; font:700 .75rem/1.2 ui-monospace,monospace; text-transform:uppercase; letter-spacing:.14em }}
    .hero-status {{ font:800 1rem/1 ui-monospace,monospace; padding:.7rem 1rem; border:2px solid currentColor }} .hero-status.pass,.badge.pass {{ color:var(--green) }} .hero-status.fail,.badge.fail {{ color:var(--red) }} .badge.skip {{ color:var(--amber) }}
    .summary {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(170px,1fr)); gap:12px; margin:24px 0 }} .metric {{ background:#fff9; border:1px solid var(--line); padding:16px }} .metric strong {{ display:block; font:700 1.6rem/1.2 Georgia,serif }}
    .table-wrap {{ overflow:auto; border:1px solid var(--line); background:#fff9 }} table {{ border-collapse:collapse; width:100%; min-width:760px }} th,td {{ padding:12px; text-align:left; vertical-align:top; border-bottom:1px solid var(--line) }} th {{ font-size:.72rem; text-transform:uppercase; letter-spacing:.08em }}
    .badge {{ font:800 .7rem/1 ui-monospace,monospace }} .artifacts {{ list-style:none; padding:0; display:grid; gap:8px }} .artifacts li {{ display:grid; gap:4px; padding:12px; background:#fff9; border:1px solid var(--line) }} code {{ overflow-wrap:anywhere }} figure {{ margin:1rem 0 }} figure img {{ display:block; width:min(100%,612px); height:auto; border:1px solid var(--line); background:white }} figcaption {{ margin-top:.5rem }}
    .playground-link {{ display:inline-block; margin-top:8px; padding:.65rem .9rem; color:white; background:#b85c30; text-decoration:none; font-weight:700 }} .playground-link:hover,.playground-link:focus-visible {{ background:#813b20 }}
    footer {{ margin-top:48px; border-top:1px solid var(--line); padding-top:16px }} @media(max-width:620px) {{ header {{ align-items:start; flex-direction:column }} }}
  </style>
</head>
<body><main>
  <header><div><p class="eyebrow">RusDox parity contract v{report_version}</p><h1>{source}</h1><p>Semantic comparison of the typed source, reopened DOCX package, and the projection consumed by the native PDF renderer.</p><a class="playground-link" href="{playground_url}">Open this example</a></div><strong class="hero-status {status_class}">{status}</strong></header>
  <section class="summary" aria-label="Report summary"><div class="metric"><span>Checks</span><strong>{check_count}</strong></div><div class="metric"><span>PDF pages</span><strong>{page_count}</strong></div><div class="metric"><span>Blocks</span><strong>{block_count}</strong></div><div class="metric"><span>Images</span><strong>{image_count}</strong></div></section>
  <h2>Contract checks</h2><div class="table-wrap"><table><thead><tr><th>ID</th><th>Assertion</th><th>Status</th><th>Evidence</th></tr></thead><tbody>{checks}</tbody></table></div>
  <h2>Generated artifacts</h2><ul class="artifacts">{artifacts}</ul>
  <h2>Visual layout comparison</h2><p>Page snapshots are deterministic geometry rasters generated from the same layout operations as the PDF. They detect movement, wrapping, page-count, table, image, and spacing regressions; they are not screenshots from Microsoft Word or a PDF viewer.</p><ul>{pages}</ul>
  <footer><p>Generated by RusDox {generator_version}. Exit code 0 means all enabled checks passed; exit code 2 means parity failed; exit code 1 means the command could not complete.</p></footer>
</main></body></html>"#,
            source = escape_html(&self.source),
            canonical = escape_html(canonical_url),
            report_version = escape_html(&self.report_version),
            playground_url = playground_url,
            status_class = status_class,
            status = status,
            check_count = self.checks.len(),
            page_count = self.pdf.page_count,
            block_count = self.expected.block_order.len(),
            image_count = self.expected.images.len(),
            checks = checks,
            artifacts = artifacts,
            pages = pages,
            generator_version = escape_html(&self.generator_version),
        )
    }
}

/// Compares current deterministic page PNGs with a baseline directory.
pub fn compare_visual_pages(
    current_dir: &Path,
    baseline_dir: Option<&Path>,
    threshold: f64,
) -> Result<VisualDiffSummary> {
    let threshold = threshold.clamp(0.0, 1.0);
    let current_pages = png_files(current_dir)?;
    let Some(baseline_dir) = baseline_dir else {
        return Ok(VisualDiffSummary {
            enabled: false,
            threshold,
            passed: true,
            pages: current_pages
                .iter()
                .enumerate()
                .map(|(index, current)| VisualPageDiff {
                    page: index + 1,
                    current: current.display().to_string(),
                    baseline: None,
                    different_pixel_ratio: None,
                    passed: true,
                    detail: "snapshot generated; no visual baseline supplied".to_string(),
                })
                .collect(),
        });
    };

    let mut pages = Vec::with_capacity(current_pages.len());
    for (index, current) in current_pages.iter().enumerate() {
        let file_name = current.file_name().unwrap_or_default();
        let baseline = baseline_dir.join(file_name);
        if !baseline.exists() {
            pages.push(VisualPageDiff {
                page: index + 1,
                current: current.display().to_string(),
                baseline: Some(baseline.display().to_string()),
                different_pixel_ratio: None,
                passed: false,
                detail: "baseline page is missing".to_string(),
            });
            continue;
        }

        let current_image = image::open(current)?;
        let baseline_image = image::open(&baseline)?;
        let (width, height) = current_image.dimensions();
        let same_dimensions = baseline_image.dimensions() == (width, height);
        let ratio = if same_dimensions {
            different_pixel_ratio(&current_image.to_rgba8(), &baseline_image.to_rgba8())
        } else {
            1.0
        };
        let passed = ratio <= threshold;
        pages.push(VisualPageDiff {
            page: index + 1,
            current: current.display().to_string(),
            baseline: Some(baseline.display().to_string()),
            different_pixel_ratio: Some(ratio),
            passed,
            detail: if same_dimensions {
                if passed {
                    "within threshold".to_string()
                } else {
                    "different-pixel ratio exceeds threshold".to_string()
                }
            } else {
                "page dimensions differ".to_string()
            },
        });
    }

    let baseline_pages = png_files(baseline_dir)?;
    if baseline_pages.len() > current_pages.len() {
        for (index, baseline) in baseline_pages.iter().enumerate().skip(current_pages.len()) {
            pages.push(VisualPageDiff {
                page: index + 1,
                current: current_dir
                    .join(baseline.file_name().unwrap_or_default())
                    .display()
                    .to_string(),
                baseline: Some(baseline.display().to_string()),
                different_pixel_ratio: None,
                passed: false,
                detail: "current render is missing a baseline page".to_string(),
            });
        }
    }

    Ok(VisualDiffSummary {
        enabled: true,
        threshold,
        passed: pages.iter().all(|page| page.passed),
        pages,
    })
}

fn equality_check<T>(id: &str, label: &str, expected: &T, docx: &T, pdf: &T) -> ParityCheck
where
    T: PartialEq + std::fmt::Debug,
{
    let docx_matches = expected == docx;
    let pdf_matches = expected == pdf;
    let passed = docx_matches && pdf_matches;
    ParityCheck {
        id: id.to_string(),
        label: label.to_string(),
        status: if passed {
            CheckStatus::Passed
        } else {
            CheckStatus::Failed
        },
        detail: if passed {
            "source = DOCX = PDF projection".to_string()
        } else {
            format!(
                "DOCX match: {docx_matches}; PDF match: {pdf_matches}; expected={expected:?}; docx={docx:?}; pdf={pdf:?}"
            )
        },
    }
}

fn boolean_check(id: &str, label: &str, passed: bool, detail: &str) -> ParityCheck {
    ParityCheck {
        id: id.to_string(),
        label: label.to_string(),
        status: if passed {
            CheckStatus::Passed
        } else {
            CheckStatus::Failed
        },
        detail: detail.to_string(),
    }
}

fn normalize_text(value: &str) -> String {
    value.split_whitespace().collect::<Vec<_>>().join(" ")
}

fn normalized_metadata(metadata: &DocumentMetadata) -> DocumentMetadata {
    let mut normalized = metadata.clone();
    if normalized
        .title
        .as_deref()
        .is_none_or(|value| value.trim().is_empty())
    {
        normalized.title = Some(crate::metadata::DEFAULT_TITLE.to_string());
    }
    if normalized
        .author
        .as_deref()
        .is_none_or(|value| value.trim().is_empty())
    {
        normalized.author = Some(crate::metadata::DEFAULT_AUTHOR.to_string());
    }
    normalized
}

fn is_heading_style_identifier(style_id: &str) -> bool {
    let style = style_id.to_ascii_lowercase();
    style.contains("title") || style.contains("heading") || style.contains("section")
}

fn png_files(directory: &Path) -> Result<Vec<PathBuf>> {
    if !directory.exists() {
        return Ok(Vec::new());
    }
    let mut files = fs::read_dir(directory)?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .filter(|path| {
            path.is_file()
                && path
                    .extension()
                    .and_then(|extension| extension.to_str())
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("png"))
        })
        .collect::<Vec<_>>();
    files.sort();
    Ok(files)
}

fn different_pixel_ratio(current: &image::RgbaImage, baseline: &image::RgbaImage) -> f64 {
    let different = current
        .pixels()
        .zip(baseline.pixels())
        .filter(|(current, baseline)| current.0 != baseline.0)
        .count();
    different as f64 / current.pixels().len().max(1) as f64
}

fn escape_html(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&#39;")
}

impl From<image::ImageError> for crate::DocxError {
    fn from(error: image::ImageError) -> Self {
        Self::Parse(format!("image processing error: {error}"))
    }
}

#[cfg(test)]
mod tests {
    use super::{
        CheckStatus, DocumentProjection, ParityReport, PdfRenderEvidence, VisualDiffSummary,
    };
    use crate::{Document, Paragraph, Run};

    #[test]
    fn projection_normalizes_text_and_tracks_breaks() {
        let mut document = Document::new();
        document.push_paragraph(Paragraph::new().add_run(Run::from_text("Hello   world")));
        document.push_paragraph(
            Paragraph::new()
                .page_break_before()
                .add_run(Run::from_text("Page two")),
        );
        let projection = DocumentProjection::from_document(&document);
        assert_eq!(projection.normalized_text, "Hello world\nPage two");
        assert_eq!(projection.page_breaks, vec![1]);
    }

    #[test]
    fn identical_projections_pass_the_semantic_contract() {
        let document =
            Document::new().add_paragraph(Paragraph::new().add_run(Run::from_text("Verified")));
        let projection = DocumentProjection::from_document(&document);
        let report = ParityReport::compare(
            "fixture.yaml",
            projection.clone(),
            projection.clone(),
            PdfRenderEvidence {
                projection,
                page_count: 1,
                draw_operations: 1,
                text_lines: 1,
                image_operations: 0,
            },
            VisualDiffSummary::skipped(0.0),
            Vec::new(),
            true,
            true,
        );
        assert!(report.passed);
        assert!(report
            .checks
            .iter()
            .all(|check| check.status != CheckStatus::Failed));
        assert!(report
            .to_html("https://example.test/report.html")
            .contains("PASS"));
    }
}
