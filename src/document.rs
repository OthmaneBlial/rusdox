use std::collections::BTreeMap;
use std::fs::File;
use std::io::{BufReader, BufWriter, Read, Seek, Write};
use std::path::{Path, PathBuf};

use tempfile::NamedTempFile;
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use crate::error::{DocxError, Result};
use crate::paragraph::Paragraph;
use crate::table::Table;
use crate::xml_utils::{
    parse_document_xml, parse_numbering_xml, write_document_xml, write_numbering_xml,
    SectionProperties,
};

const CONTENT_TYPES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"#;

const PACKAGE_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"#;

const DOCUMENT_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdRusDoxNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>
"#;

const APP_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>RusDox</Application>
</Properties>
"#;

const CORE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>RusDox Document</dc:title>
  <dc:creator>RusDox</dc:creator>
  <cp:lastModifiedBy>RusDox</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-03-10T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-03-10T00:00:00Z</dcterms:modified>
</cp:coreProperties>
"#;

const STYLES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>
"#;

/// Controls how a document instance is intended to be used.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentMode {
    /// A new document created from scratch.
    New,
    /// A document opened for read-only inspection.
    ReadOnly,
    /// A document opened for reading and writing.
    ReadWrite,
}

/// An ordered view of top-level document content.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentBlockRef<'a> {
    /// A paragraph block.
    Paragraph(&'a Paragraph),
    /// A table block.
    Table(&'a Table),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum BodyBlock {
    Paragraph(Paragraph),
    Table(Box<Table>),
}

impl BodyBlock {
    fn text(&self) -> String {
        match self {
            Self::Paragraph(paragraph) => paragraph.text(),
            Self::Table(table) => table.text(),
        }
    }
}

/// A `.docx` document containing paragraphs and tables.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Document {
    mode: DocumentMode,
    body: Vec<BodyBlock>,
    package_parts: BTreeMap<String, Vec<u8>>,
    source_path: Option<PathBuf>,
    section_properties: SectionProperties,
}

impl Default for Document {
    fn default() -> Self {
        Self::new()
    }
}

impl Document {
    /// Creates a new empty document with a minimal OOXML package template.
    ///
    /// ```rust
    /// use rusdox::Document;
    ///
    /// let document = Document::new();
    /// assert_eq!(document.paragraphs().count(), 0);
    /// ```
    pub fn new() -> Self {
        Self {
            mode: DocumentMode::New,
            body: Vec::new(),
            package_parts: default_package_parts(),
            source_path: None,
            section_properties: SectionProperties::default(),
        }
    }

    /// Opens a document for reading and writing.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_mode(path, DocumentMode::ReadWrite)
    }

    /// Opens a document in read-only mode.
    pub fn open_read_only(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_mode(path, DocumentMode::ReadOnly)
    }

    /// Opens a document with an explicit mode.
    pub fn open_with_mode(path: impl AsRef<Path>, mode: DocumentMode) -> Result<Self> {
        let path = path.as_ref();
        let file = File::open(path)?;
        let mut document = Self::open_from_reader(BufReader::new(file), mode)?;
        document.source_path = Some(path.to_path_buf());
        Ok(document)
    }

    /// Opens a document from any reader implementing `Read + Seek`.
    pub fn open_from_reader<R>(reader: R, mode: DocumentMode) -> Result<Self>
    where
        R: Read + Seek,
    {
        let mut archive = ZipArchive::new(reader)?;
        let mut package_parts = BTreeMap::new();
        let mut document_xml = None;
        let mut numbering = None;

        for index in 0..archive.len() {
            let mut entry = archive.by_index(index)?;
            let name = entry.name().to_string();
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes)?;
            if name == "word/document.xml" {
                document_xml = Some(bytes);
            } else if name == "word/numbering.xml" {
                numbering = Some(parse_numbering_xml(&bytes)?);
                package_parts.insert(name, bytes);
            } else {
                package_parts.insert(name, bytes);
            }
        }

        let document_xml = document_xml
            .ok_or_else(|| DocxError::parse("missing OOXML part: word/document.xml"))?;
        let parsed = parse_document_xml(&document_xml, numbering.as_ref())?;

        Ok(Self {
            mode: match mode {
                DocumentMode::New => DocumentMode::ReadWrite,
                other => other,
            },
            body: parsed.body,
            package_parts,
            source_path: None,
            section_properties: parsed.section_properties,
        })
    }

    /// Returns the current document mode.
    pub fn mode(&self) -> DocumentMode {
        self.mode
    }

    /// Returns the original source path, if the document was opened from disk.
    pub fn source_path(&self) -> Option<&Path> {
        self.source_path.as_deref()
    }

    /// Adds a paragraph to the end of the document body.
    pub fn push_paragraph(&mut self, paragraph: Paragraph) -> &mut Self {
        self.body.push(BodyBlock::Paragraph(paragraph));
        self
    }

    /// Adds a paragraph in a builder-style fashion.
    pub fn add_paragraph(mut self, paragraph: Paragraph) -> Self {
        self.body.push(BodyBlock::Paragraph(paragraph));
        self
    }

    /// Adds a table to the end of the document body.
    pub fn push_table(&mut self, table: Table) -> &mut Self {
        self.body.push(BodyBlock::Table(Box::new(table)));
        self
    }

    /// Adds a table in a builder-style fashion.
    pub fn add_table(mut self, table: Table) -> Self {
        self.body.push(BodyBlock::Table(Box::new(table)));
        self
    }

    /// Returns immutable access to top-level paragraphs.
    pub fn paragraphs(&self) -> impl Iterator<Item = &Paragraph> {
        self.body.iter().filter_map(|block| match block {
            BodyBlock::Paragraph(paragraph) => Some(paragraph),
            BodyBlock::Table(_) => None,
        })
    }

    /// Returns mutable access to top-level paragraphs.
    pub fn paragraphs_mut(&mut self) -> impl Iterator<Item = &mut Paragraph> {
        self.body.iter_mut().filter_map(|block| match block {
            BodyBlock::Paragraph(paragraph) => Some(paragraph),
            BodyBlock::Table(_) => None,
        })
    }

    /// Returns immutable access to top-level tables.
    pub fn tables(&self) -> impl Iterator<Item = &Table> {
        self.body.iter().filter_map(|block| match block {
            BodyBlock::Paragraph(_) => None,
            BodyBlock::Table(table) => Some(table.as_ref()),
        })
    }

    /// Returns mutable access to top-level tables.
    pub fn tables_mut(&mut self) -> impl Iterator<Item = &mut Table> {
        self.body.iter_mut().filter_map(|block| match block {
            BodyBlock::Paragraph(_) => None,
            BodyBlock::Table(table) => Some(table.as_mut()),
        })
    }

    /// Returns top-level document blocks in their original order.
    pub fn blocks(&self) -> impl Iterator<Item = DocumentBlockRef<'_>> {
        self.body.iter().map(|block| match block {
            BodyBlock::Paragraph(paragraph) => DocumentBlockRef::Paragraph(paragraph),
            BodyBlock::Table(table) => DocumentBlockRef::Table(table.as_ref()),
        })
    }

    /// Extracts the document plain text.
    pub fn text(&self) -> String {
        self.body
            .iter()
            .map(BodyBlock::text)
            .filter(|value| !value.is_empty())
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Saves the document to a path using a temporary file and rename step.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        if self.mode == DocumentMode::ReadOnly {
            return Err(DocxError::parse("cannot save a read-only document"));
        }

        let destination = path.as_ref();
        let parent = destination.parent().unwrap_or_else(|| Path::new("."));
        let mut temporary = NamedTempFile::new_in(parent)?;

        {
            let mut writer = BufWriter::new(temporary.as_file_mut());
            self.save_to_writer(&mut writer)?;
            writer.flush()?;
        }

        #[cfg(windows)]
        if destination.exists() {
            std::fs::remove_file(destination)?;
        }

        temporary
            .persist(destination)
            .map_err(|error| DocxError::Io(error.error))?;
        Ok(())
    }

    /// Writes the document archive to any writer implementing `Write + Seek`.
    pub fn save_to_writer<W>(&self, writer: W) -> Result<()>
    where
        W: Write + Seek,
    {
        if self.mode == DocumentMode::ReadOnly {
            return Err(DocxError::parse("cannot save a read-only document"));
        }

        let mut archive = ZipWriter::new(writer);
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        let package_parts = self.render_package_parts()?;

        archive.start_file("word/document.xml", options)?;
        let document_xml = write_document_xml(&self.body, &self.section_properties)?;
        archive.write_all(&document_xml)?;

        for (name, contents) in &package_parts {
            archive.start_file(name, options)?;
            archive.write_all(contents)?;
        }

        archive.finish()?;
        Ok(())
    }

    fn render_package_parts(&self) -> Result<BTreeMap<String, Vec<u8>>> {
        let mut package_parts = self.package_parts.clone();
        package_parts.insert(
            "word/numbering.xml".to_string(),
            write_numbering_xml(&self.body)?,
        );
        ensure_numbering_content_type(&mut package_parts)?;
        ensure_numbering_relationship(&mut package_parts)?;
        Ok(package_parts)
    }
}

fn ensure_numbering_content_type(parts: &mut BTreeMap<String, Vec<u8>>) -> Result<()> {
    let xml = parts
        .entry("[Content_Types].xml".to_string())
        .or_insert_with(|| CONTENT_TYPES_XML.as_bytes().to_vec());
    ensure_xml_fragment(
        xml,
        "</Types>",
        r#"PartName="/word/numbering.xml""#,
        r#"  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
"#,
    )
}

fn ensure_numbering_relationship(parts: &mut BTreeMap<String, Vec<u8>>) -> Result<()> {
    let xml = parts
        .entry("word/_rels/document.xml.rels".to_string())
        .or_insert_with(|| DOCUMENT_RELS_XML.as_bytes().to_vec());
    ensure_xml_fragment(
        xml,
        "</Relationships>",
        r#"Target="numbering.xml""#,
        r#"  <Relationship Id="rIdRusDoxNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
"#,
    )
}

fn ensure_xml_fragment(
    xml_bytes: &mut Vec<u8>,
    closing_tag: &str,
    needle: &str,
    fragment: &str,
) -> Result<()> {
    let mut xml = String::from_utf8(xml_bytes.clone())
        .map_err(|_| DocxError::parse("OOXML package part was not valid UTF-8"))?;

    if xml.contains(needle) {
        return Ok(());
    }

    let Some(index) = xml.rfind(closing_tag) else {
        return Err(DocxError::parse(format!(
            "malformed OOXML package part: missing closing tag {closing_tag}"
        )));
    };

    xml.insert_str(index, fragment);
    *xml_bytes = xml.into_bytes();
    Ok(())
}

fn default_package_parts() -> BTreeMap<String, Vec<u8>> {
    let mut parts = BTreeMap::new();
    parts.insert(
        "[Content_Types].xml".to_string(),
        CONTENT_TYPES_XML.as_bytes().to_vec(),
    );
    parts.insert(
        "_rels/.rels".to_string(),
        PACKAGE_RELS_XML.as_bytes().to_vec(),
    );
    parts.insert("docProps/app.xml".to_string(), APP_XML.as_bytes().to_vec());
    parts.insert(
        "docProps/core.xml".to_string(),
        CORE_XML.as_bytes().to_vec(),
    );
    parts.insert(
        "word/_rels/document.xml.rels".to_string(),
        DOCUMENT_RELS_XML.as_bytes().to_vec(),
    );
    parts.insert(
        "word/styles.xml".to_string(),
        STYLES_XML.as_bytes().to_vec(),
    );
    parts.insert("word/numbering.xml".to_string(), Vec::new());
    parts
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};

    use tempfile::tempdir;
    use zip::write::SimpleFileOptions;
    use zip::{CompressionMethod, ZipWriter};

    use super::{BodyBlock, Document, DocumentBlockRef, DocumentMode};
    use crate::{Paragraph, Run, Table, TableCell, TableRow};

    fn sample_document() -> Document {
        let mut document = Document::new();
        document.push_paragraph(Paragraph::new().add_run(Run::from_text("P1")));
        document.push_table(Table::new().add_row(TableRow::new().add_cell(
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("T1"))),
        )));
        document
    }

    #[test]
    fn new_document_has_expected_defaults() {
        let document = Document::new();
        assert_eq!(document.mode(), DocumentMode::New);
        assert!(document.source_path().is_none());
        assert_eq!(document.paragraphs().count(), 0);
        assert_eq!(document.tables().count(), 0);
    }

    #[test]
    fn open_from_reader_converts_new_mode_to_read_write() {
        let document = sample_document();
        let mut buffer = Cursor::new(Vec::new());
        document
            .save_to_writer(&mut buffer)
            .expect("save test document");
        buffer.set_position(0);

        let reopened =
            Document::open_from_reader(&mut buffer, DocumentMode::New).expect("open from reader");
        assert_eq!(reopened.mode(), DocumentMode::ReadWrite);
    }

    #[test]
    fn add_builder_methods_append_blocks() {
        let document = Document::new()
            .add_paragraph(Paragraph::new().add_run(Run::from_text("A")))
            .add_table(Table::new().add_row(TableRow::new().add_cell(
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("B"))),
            )));

        assert_eq!(document.paragraphs().count(), 1);
        assert_eq!(document.tables().count(), 1);
    }

    #[test]
    fn blocks_preserve_original_top_level_order() {
        let document = sample_document()
            .add_paragraph(Paragraph::new().add_run(Run::from_text("P2")))
            .add_table(Table::new().add_row(TableRow::new().add_cell(
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("T2"))),
            )));

        let kinds: Vec<_> = document
            .blocks()
            .map(|block| match block {
                DocumentBlockRef::Paragraph(_) => "p",
                DocumentBlockRef::Table(_) => "t",
            })
            .collect();
        assert_eq!(kinds, vec!["p", "t", "p", "t"]);
    }

    #[test]
    fn text_joins_paragraph_and_table_content() {
        let mut document = Document::new();
        document.push_paragraph(Paragraph::new().add_run(Run::from_text("Alpha")));
        document.push_paragraph(Paragraph::new().add_run(Run::from_text("Beta")));
        document.push_table(
            Table::new().add_row(
                TableRow::new()
                    .add_cell(
                        TableCell::new()
                            .add_paragraph(Paragraph::new().add_run(Run::from_text("C1"))),
                    )
                    .add_cell(
                        TableCell::new()
                            .add_paragraph(Paragraph::new().add_run(Run::from_text("C2"))),
                    ),
            ),
        );

        assert_eq!(document.text(), "Alpha\nBeta\nC1\tC2");
    }

    #[test]
    fn paragraphs_mut_and_tables_mut_enable_updates() {
        let mut document = sample_document();

        for paragraph in document.paragraphs_mut() {
            paragraph.push_run(Run::from_text("!"));
        }
        for table in document.tables_mut() {
            table.properties_mut().width = Some(9000);
        }

        assert_eq!(
            document.paragraphs().next().expect("paragraph").text(),
            "P1!"
        );
        assert_eq!(
            document.tables().next().expect("table").properties().width,
            Some(9000)
        );
    }

    #[test]
    fn save_to_writer_rejects_read_only_mode() {
        let document = sample_document();
        let mut buffer = Cursor::new(Vec::new());
        document
            .save_to_writer(&mut buffer)
            .expect("save for reopen");
        buffer.set_position(0);

        let reopened = Document::open_from_reader(&mut buffer, DocumentMode::ReadOnly)
            .expect("open read-only");
        let error = reopened
            .save_to_writer(Cursor::new(Vec::new()))
            .expect_err("must fail");
        assert!(matches!(error, crate::DocxError::Parse(message) if message.contains("read-only")));
    }

    #[test]
    fn open_from_reader_fails_when_main_document_part_missing() {
        let mut archive_buffer = Cursor::new(Vec::new());
        {
            let mut zip = ZipWriter::new(&mut archive_buffer);
            let options =
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
            zip.start_file("word/styles.xml", options)
                .expect("start styles");
            zip.write_all(b"<w:styles/>").expect("write styles");
            zip.finish().expect("finish zip");
        }

        archive_buffer.set_position(0);
        let error = Document::open_from_reader(&mut archive_buffer, DocumentMode::ReadWrite)
            .expect_err("missing word/document.xml must fail");
        assert!(
            matches!(error, crate::DocxError::Parse(message) if message.contains("word/document.xml"))
        );
    }

    #[test]
    fn save_and_open_round_trip_from_disk_sets_source_path() {
        let temp = tempdir().expect("temp dir");
        let path = temp.path().join("doc.docx");
        let mut document = sample_document();
        document.push_paragraph(Paragraph::new().add_run(Run::from_text("disk")));
        document.save(&path).expect("save to disk");

        let reopened = Document::open(&path).expect("open from disk");
        assert_eq!(reopened.mode(), DocumentMode::ReadWrite);
        assert_eq!(reopened.source_path(), Some(path.as_path()));
        assert!(reopened.text().contains("disk"));
    }

    #[test]
    fn save_overwrites_existing_file() {
        let temp = tempdir().expect("temp dir");
        let path = temp.path().join("replace.docx");

        let mut first = Document::new();
        first.push_paragraph(Paragraph::new().add_run(Run::from_text("first")));
        first.save(&path).expect("save first");

        let mut second = Document::new();
        second.push_paragraph(Paragraph::new().add_run(Run::from_text("second")));
        second.save(&path).expect("save second");

        let reopened = Document::open(&path).expect("reopen");
        assert_eq!(reopened.text(), "second");
    }

    #[test]
    fn body_block_text_behaves_for_paragraph_and_table() {
        let paragraph = BodyBlock::Paragraph(Paragraph::new().add_run(Run::from_text("p")));
        let table = BodyBlock::Table(Box::new(Table::new().add_row(TableRow::new().add_cell(
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("t"))),
        ))));

        assert_eq!(paragraph.text(), "p");
        assert_eq!(table.text(), "t");
    }

    #[test]
    fn render_package_parts_preserves_existing_relationships_and_adds_numbering() {
        let mut document = Document::new();
        document.push_paragraph(Paragraph::new().add_run(Run::from_text("Item")));
        document.package_parts.insert(
            "[Content_Types].xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"#
            .to_vec(),
        );
        document.package_parts.insert(
            "word/_rels/document.xml.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"#
            .to_vec(),
        );

        let parts = document
            .render_package_parts()
            .expect("render package parts");
        let content_types =
            String::from_utf8(parts["[Content_Types].xml"].clone()).expect("utf8 content types");
        let rels = String::from_utf8(parts["word/_rels/document.xml.rels"].clone())
            .expect("utf8 relationships");

        assert!(content_types.contains(r#"PartName="/word/numbering.xml""#));
        assert!(rels.contains(r#"Target="styles.xml""#));
        assert!(rels.contains(r#"Target="numbering.xml""#));
    }
}
