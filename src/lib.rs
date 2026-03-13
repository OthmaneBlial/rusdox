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
mod layout;
mod paragraph;
mod run;
pub mod spec;
pub mod studio;
mod style;
mod table;
mod visual;
mod xml_utils;

pub use document::{Document, DocumentBlockRef, DocumentMode};
pub use error::{DocxError, Result};
pub use layout::{HeaderFooter, PageNumberFormat, PageNumbering, PageSetup};
pub use paragraph::{Paragraph, ParagraphAlignment, ParagraphList, ParagraphListKind};
pub use run::{Run, RunProperties, UnderlineStyle, VerticalAlign};
pub use style::{
    ParagraphStyle, ParagraphStyleProperties, RunStyle, RunStyleProperties, Stylesheet, TableStyle,
    TableStyleProperties,
};
pub use table::{
    Border, BorderStyle, Table, TableBorders, TableCell, TableCellProperties, TableProperties,
    TableRow, TableRowProperties,
};
pub use visual::{Visual, VisualFormat, VisualKind, VisualSizing, VisualSource};
