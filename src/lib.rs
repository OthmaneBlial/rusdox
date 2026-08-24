//! `RusDox` is a focused Rust library for building, reading, and saving
//! Microsoft Word `.docx` documents.
//!
//! The crate models the document body with strongly typed paragraphs, runs,
//! and tables while using fast ZIP and XML primitives internally.
//!
//! # Example
//!
//! ```rust
//! use rusdox::{Document, Paragraph, Run};
//!
//! let mut document = Document::new();
//! document.push_paragraph(
//!     Paragraph::new()
//!         .add_run(Run::from_text("Hello ").bold())
//!         .add_run(Run::from_text("RusDox").italic()),
//! );
//!
//! assert_eq!(document.paragraphs().count(), 1);
//! ```

pub mod config;
mod document;
mod error;
mod io_utils;
mod layout;
mod limits;
mod metadata;
mod package_validate;
mod paragraph;
pub mod parity;
mod run;
pub mod schema;
mod source;
pub mod spec;
mod spec_expand;
pub mod studio;
mod style;
mod table;
mod template;
pub mod validate;
mod visual;
mod xml_utils;

pub use document::{Document, DocumentBlockRef, DocumentMode};
pub use error::{DocxError, Result};
pub use io_utils::atomic_write as atomic_write_file;
pub use layout::{HeaderFooter, PageNumberFormat, PageNumbering, PageOrientation, PageSetup};
pub use limits::InputLimits;
pub use metadata::DocumentMetadata;
pub use package_validate::{
    validate_docx_package, validate_docx_package_with_limits, validate_docx_reader_with_limits,
    PackageValidationReport,
};
pub use paragraph::{Paragraph, ParagraphAlignment, ParagraphList, ParagraphListKind};
pub use run::{Run, RunField, RunProperties, UnderlineStyle, VerticalAlign};
pub use schema::{document_spec_schema, document_spec_schema_pretty, DOCUMENT_SPEC_SCHEMA_ID};
pub use source::{attach_source_spans, attach_source_spans_from_str};
pub use style::{
    ParagraphStyle, ParagraphStyleProperties, RunStyle, RunStyleProperties, Stylesheet, TableStyle,
    TableStyleProperties,
};
pub use table::{
    Border, BorderStyle, Table, TableBorders, TableCell, TableCellProperties, TableProperties,
    TableRow, TableRowProperties,
};
pub use template::{
    DocxTemplate, TemplateDiagnostic, TemplateDiagnosticSeverity, TemplateInspection,
    TemplatePlaceholder, TemplateRenderReport, TEMPLATE_SYNTAX_VERSION,
};
pub use validate::{
    validate_config, validate_spec, validate_spec_with_config, SourceSpan, ValidationIssue,
    ValidationReport, ValidationSeverity,
};
pub use visual::{Visual, VisualFormat, VisualKind, VisualSizing, VisualSource};
