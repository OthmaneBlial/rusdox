use std::collections::BTreeMap;
use std::fs::File;
use std::io::{BufReader, BufWriter, Read, Seek, Write};
use std::path::{Path, PathBuf};

use tempfile::NamedTempFile;
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use crate::error::{DocxError, Result};
use crate::layout::{HeaderFooter, PageNumbering, PageSetup};
use crate::metadata::{
    parse_core_properties_xml, parse_custom_properties_xml, render_core_properties_xml,
    render_custom_properties_xml, DocumentMetadata,
};
use crate::paragraph::Paragraph;
use crate::style::Stylesheet;
use crate::table::Table;
use crate::visual::Visual;
use crate::xml_utils::{
    hydrate_section_from_package_parts, parse_document_xml, parse_numbering_xml, parse_styles_xml,
    render_header_footer_xml, write_document_xml, write_numbering_xml, write_styles_xml,
    DocxVisual, SectionProperties, DEFAULT_FOOTER_PART, DEFAULT_FOOTER_REL_ID, DEFAULT_HEADER_PART,
    DEFAULT_HEADER_REL_ID, FOOTER_REL_TYPE, HEADER_REL_TYPE, IMAGE_REL_TYPE,
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
  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
</Types>
"#;

const PACKAGE_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>
</Relationships>
"#;

const DOCUMENT_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdRusDoxStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
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

const CUSTOM_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
</Properties>
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
    /// A visual/image block.
    Visual(&'a Visual),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum BodyBlock {
    Paragraph(Paragraph),
    Table(Box<Table>),
    Visual(Visual),
}

impl BodyBlock {
    fn text(&self) -> String {
        match self {
            Self::Paragraph(paragraph) => paragraph.text(),
            Self::Table(table) => table.text(),
            Self::Visual(visual) => visual.alt_text().unwrap_or_default().to_string(),
        }
    }
}

struct RenderedPackage {
    parts: BTreeMap<String, Vec<u8>>,
    visuals: Vec<DocxVisual>,
}

struct RenderedVisualPart {
    xml: DocxVisual,
    part_name: String,
    target: String,
    content_type: &'static str,
    bytes: Vec<u8>,
}

/// A `.docx` document containing paragraphs and tables.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Document {
    mode: DocumentMode,
    body: Vec<BodyBlock>,
    metadata: DocumentMetadata,
    styles: Stylesheet,
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
            metadata: DocumentMetadata::default(),
            styles: Stylesheet::default(),
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
        let mut styles_xml = None;
        let mut core_xml = None;
        let mut custom_xml = None;

        for index in 0..archive.len() {
            let mut entry = archive.by_index(index)?;
            let name = entry.name().to_string();
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes)?;
            if name == "word/document.xml" {
                document_xml = Some(bytes);
            } else if name == "docProps/core.xml" {
                core_xml = Some(bytes.clone());
                package_parts.insert(name, bytes);
            } else if name == "docProps/custom.xml" {
                custom_xml = Some(bytes.clone());
                package_parts.insert(name, bytes);
            } else if name == "word/numbering.xml" {
                numbering = Some(parse_numbering_xml(&bytes)?);
                package_parts.insert(name, bytes);
            } else if name == "word/styles.xml" {
                styles_xml = Some(bytes.clone());
                package_parts.insert(name, bytes);
            } else {
                package_parts.insert(name, bytes);
            }
        }

        let document_xml = document_xml
            .ok_or_else(|| DocxError::parse("missing OOXML part: word/document.xml"))?;
        let styles = match styles_xml.as_deref() {
            Some(bytes) => parse_styles_xml(bytes, numbering.as_ref())?,
            None => Stylesheet::default(),
        };
        let mut parsed = parse_document_xml(&document_xml, numbering.as_ref(), &package_parts)?;
        hydrate_section_from_package_parts(&mut parsed.section_properties, &package_parts)?;
        let mut metadata = core_xml
            .as_deref()
            .map(parse_core_properties_xml)
            .transpose()?
            .unwrap_or_default();
        if let Some(bytes) = custom_xml.as_deref() {
            parse_custom_properties_xml(bytes, &mut metadata)?;
        }

        Ok(Self {
            mode: match mode {
                DocumentMode::New => DocumentMode::ReadWrite,
                other => other,
            },
            body: parsed.body,
            metadata,
            styles,
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

    /// Returns the active page setup.
    pub fn page_setup(&self) -> &PageSetup {
        self.section_properties.page_setup()
    }

    /// Replaces the active page setup.
    pub fn set_page_setup(&mut self, page_setup: PageSetup) -> &mut Self {
        self.section_properties.set_page_setup(page_setup);
        self
    }

    /// Replaces the active page setup in builder style.
    pub fn with_page_setup(mut self, page_setup: PageSetup) -> Self {
        self.section_properties.set_page_setup(page_setup);
        self
    }

    /// Returns the default header template when present.
    pub fn header(&self) -> Option<&HeaderFooter> {
        self.section_properties.header()
    }

    /// Sets or clears the default header template.
    pub fn set_header(&mut self, header: Option<HeaderFooter>) -> &mut Self {
        self.section_properties.set_header(header);
        self
    }

    /// Sets the default header template in builder style.
    pub fn with_header(mut self, header: HeaderFooter) -> Self {
        self.section_properties.set_header(Some(header));
        self
    }

    /// Returns the default footer template when present.
    pub fn footer(&self) -> Option<&HeaderFooter> {
        self.section_properties.footer()
    }

    /// Sets or clears the default footer template.
    pub fn set_footer(&mut self, footer: Option<HeaderFooter>) -> &mut Self {
        self.section_properties.set_footer(footer);
        self
    }

    /// Sets the default footer template in builder style.
    pub fn with_footer(mut self, footer: HeaderFooter) -> Self {
        self.section_properties.set_footer(Some(footer));
        self
    }

    /// Returns page numbering settings when present.
    pub fn page_numbering(&self) -> Option<&PageNumbering> {
        self.section_properties.page_numbering()
    }

    /// Returns the current document metadata.
    pub fn metadata(&self) -> &DocumentMetadata {
        &self.metadata
    }

    /// Returns mutable access to the document metadata.
    pub fn metadata_mut(&mut self) -> &mut DocumentMetadata {
        &mut self.metadata
    }

    /// Replaces the document metadata.
    pub fn set_metadata(&mut self, metadata: DocumentMetadata) -> &mut Self {
        self.metadata = metadata;
        self
    }

    /// Replaces the document metadata in builder style.
    pub fn with_metadata(mut self, metadata: DocumentMetadata) -> Self {
        self.metadata = metadata;
        self
    }

    /// Returns the document stylesheet.
    pub fn styles(&self) -> &Stylesheet {
        &self.styles
    }

    /// Returns mutable access to the document stylesheet.
    pub fn styles_mut(&mut self) -> &mut Stylesheet {
        &mut self.styles
    }

    /// Replaces the document stylesheet.
    pub fn set_styles(&mut self, styles: Stylesheet) -> &mut Self {
        self.styles = styles;
        self
    }

    /// Replaces the document stylesheet in builder style.
    pub fn with_styles(mut self, styles: Stylesheet) -> Self {
        self.styles = styles;
        self
    }

    /// Sets or clears page numbering settings.
    pub fn set_page_numbering(&mut self, page_numbering: Option<PageNumbering>) -> &mut Self {
        self.section_properties.set_page_numbering(page_numbering);
        self
    }

    /// Sets page numbering settings in builder style.
    pub fn with_page_numbering(mut self, page_numbering: PageNumbering) -> Self {
        self.section_properties
            .set_page_numbering(Some(page_numbering));
        self
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

    /// Adds a visual block to the end of the document body.
    pub fn push_visual(&mut self, visual: Visual) -> &mut Self {
        self.body.push(BodyBlock::Visual(visual));
        self
    }

    /// Adds a visual block in a builder-style fashion.
    pub fn add_visual(mut self, visual: Visual) -> Self {
        self.body.push(BodyBlock::Visual(visual));
        self
    }

    /// Returns immutable access to top-level paragraphs.
    pub fn paragraphs(&self) -> impl Iterator<Item = &Paragraph> {
        self.body.iter().filter_map(|block| match block {
            BodyBlock::Paragraph(paragraph) => Some(paragraph),
            BodyBlock::Table(_) | BodyBlock::Visual(_) => None,
        })
    }

    /// Returns mutable access to top-level paragraphs.
    pub fn paragraphs_mut(&mut self) -> impl Iterator<Item = &mut Paragraph> {
        self.body.iter_mut().filter_map(|block| match block {
            BodyBlock::Paragraph(paragraph) => Some(paragraph),
            BodyBlock::Table(_) | BodyBlock::Visual(_) => None,
        })
    }

    /// Returns immutable access to top-level tables.
    pub fn tables(&self) -> impl Iterator<Item = &Table> {
        self.body.iter().filter_map(|block| match block {
            BodyBlock::Paragraph(_) | BodyBlock::Visual(_) => None,
            BodyBlock::Table(table) => Some(table.as_ref()),
        })
    }

    /// Returns mutable access to top-level tables.
    pub fn tables_mut(&mut self) -> impl Iterator<Item = &mut Table> {
        self.body.iter_mut().filter_map(|block| match block {
            BodyBlock::Paragraph(_) | BodyBlock::Visual(_) => None,
            BodyBlock::Table(table) => Some(table.as_mut()),
        })
    }

    /// Returns immutable access to top-level visual blocks.
    pub fn visuals(&self) -> impl Iterator<Item = &Visual> {
        self.body.iter().filter_map(|block| match block {
            BodyBlock::Visual(visual) => Some(visual),
            BodyBlock::Paragraph(_) | BodyBlock::Table(_) => None,
        })
    }

    /// Returns mutable access to top-level visual blocks.
    pub fn visuals_mut(&mut self) -> impl Iterator<Item = &mut Visual> {
        self.body.iter_mut().filter_map(|block| match block {
            BodyBlock::Visual(visual) => Some(visual),
            BodyBlock::Paragraph(_) | BodyBlock::Table(_) => None,
        })
    }

    /// Returns top-level document blocks in their original order.
    pub fn blocks(&self) -> impl Iterator<Item = DocumentBlockRef<'_>> {
        self.body.iter().map(|block| match block {
            BodyBlock::Paragraph(paragraph) => DocumentBlockRef::Paragraph(paragraph),
            BodyBlock::Table(table) => DocumentBlockRef::Table(table.as_ref()),
            BodyBlock::Visual(visual) => DocumentBlockRef::Visual(visual),
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
        let rendered_package = self.render_package_parts()?;

        archive.start_file("word/document.xml", options)?;
        let document_xml = write_document_xml(
            &self.body,
            &self.section_properties,
            &rendered_package.visuals,
        )?;
        archive.write_all(&document_xml)?;

        for (name, contents) in &rendered_package.parts {
            archive.start_file(name, options)?;
            archive.write_all(contents)?;
        }

        archive.finish()?;
        Ok(())
    }

    fn render_package_parts(&self) -> Result<RenderedPackage> {
        let mut package_parts = self.package_parts.clone();
        let visuals = self.render_visual_parts()?;
        package_parts.insert(
            "docProps/core.xml".to_string(),
            render_core_properties_xml(&self.metadata)?,
        );
        package_parts.insert(
            "docProps/custom.xml".to_string(),
            render_custom_properties_xml(&self.metadata)?,
        );
        package_parts.insert(
            "word/styles.xml".to_string(),
            write_styles_xml(&self.styles)?,
        );
        ensure_custom_content_type(&mut package_parts)?;
        ensure_package_relationship(
            &mut package_parts,
            "rId4",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
            "docProps/custom.xml",
        )?;
        ensure_styles_relationship(&mut package_parts)?;
        package_parts.insert(
            "word/numbering.xml".to_string(),
            write_numbering_xml(&self.body, &self.styles)?,
        );
        ensure_numbering_content_type(&mut package_parts)?;
        ensure_numbering_relationship(&mut package_parts)?;
        for visual in &visuals {
            package_parts.insert(visual.part_name.clone(), visual.bytes.clone());
            ensure_default_content_type(
                &mut package_parts,
                visual.part_name.rsplit('.').next().unwrap_or_default(),
                visual.content_type,
            )?;
            ensure_header_footer_relationship(
                &mut package_parts,
                &visual.xml.relation_id,
                IMAGE_REL_TYPE,
                &visual.target,
            )?;
        }
        if let Some(header) = self.header() {
            package_parts.insert(
                DEFAULT_HEADER_PART.to_string(),
                render_header_footer_xml(header, "w:hdr")?,
            );
            ensure_header_footer_content_type(
                &mut package_parts,
                DEFAULT_HEADER_PART,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
            )?;
            ensure_header_footer_relationship(
                &mut package_parts,
                DEFAULT_HEADER_REL_ID,
                HEADER_REL_TYPE,
                "header1.xml",
            )?;
        }
        if let Some(footer) = self.footer() {
            package_parts.insert(
                DEFAULT_FOOTER_PART.to_string(),
                render_header_footer_xml(footer, "w:ftr")?,
            );
            ensure_header_footer_content_type(
                &mut package_parts,
                DEFAULT_FOOTER_PART,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
            )?;
            ensure_header_footer_relationship(
                &mut package_parts,
                DEFAULT_FOOTER_REL_ID,
                FOOTER_REL_TYPE,
                "footer1.xml",
            )?;
        }
        Ok(RenderedPackage {
            parts: package_parts,
            visuals: visuals.into_iter().map(|visual| visual.xml).collect(),
        })
    }

    fn render_visual_parts(&self) -> Result<Vec<RenderedVisualPart>> {
        let page_setup = self.page_setup();
        let content_width_twips = page_setup
            .width_twips
            .saturating_sub(page_setup.margin_left_twips)
            .saturating_sub(page_setup.margin_right_twips)
            .saturating_sub(page_setup.gutter_twips)
            .max(1);
        let content_height_twips = page_setup
            .height_twips
            .saturating_sub(page_setup.margin_top_twips)
            .saturating_sub(page_setup.margin_bottom_twips)
            .max(1);

        let mut rendered = Vec::new();
        for (index, visual) in self.visuals().enumerate() {
            let sequence = index + 1;
            let (width_twips, height_twips) =
                visual.resolved_dimensions_twips(content_width_twips, content_height_twips)?;
            let (format, bytes) = visual.docx_media(width_twips, height_twips)?;
            let extension = format.extension();
            rendered.push(RenderedVisualPart {
                xml: DocxVisual {
                    relation_id: format!("rIdRusDoxImage{sequence}"),
                    width_emu: width_twips.saturating_mul(635),
                    height_emu: height_twips.saturating_mul(635),
                    doc_pr_id: sequence as u32,
                    name: format!("{} {sequence}", visual.docx_name()),
                    alt_text: visual.alt_text().map(str::to_owned),
                },
                part_name: format!("word/media/rusdox-image-{sequence}.{extension}"),
                target: format!("media/rusdox-image-{sequence}.{extension}"),
                content_type: format.content_type(),
                bytes,
            });
        }

        Ok(rendered)
    }
}

fn ensure_styles_relationship(parts: &mut BTreeMap<String, Vec<u8>>) -> Result<()> {
    let xml = parts
        .entry("word/_rels/document.xml.rels".to_string())
        .or_insert_with(|| DOCUMENT_RELS_XML.as_bytes().to_vec());
    ensure_xml_fragment(
        xml,
        "</Relationships>",
        r#"Target="styles.xml""#,
        r#"  <Relationship Id="rIdRusDoxStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
"#,
    )
}

fn ensure_custom_content_type(parts: &mut BTreeMap<String, Vec<u8>>) -> Result<()> {
    let xml = parts
        .entry("[Content_Types].xml".to_string())
        .or_insert_with(|| CONTENT_TYPES_XML.as_bytes().to_vec());
    ensure_xml_fragment(
        xml,
        "</Types>",
        r#"PartName="/docProps/custom.xml""#,
        r#"  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
"#,
    )
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

fn ensure_package_relationship(
    parts: &mut BTreeMap<String, Vec<u8>>,
    relation_id: &str,
    relation_type: &str,
    target: &str,
) -> Result<()> {
    let xml = parts
        .entry("_rels/.rels".to_string())
        .or_insert_with(|| PACKAGE_RELS_XML.as_bytes().to_vec());
    ensure_xml_fragment(
        xml,
        "</Relationships>",
        relation_id,
        &format!(
            r#"  <Relationship Id="{relation_id}" Type="{relation_type}" Target="{target}"/>
"#
        ),
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

fn ensure_default_content_type(
    parts: &mut BTreeMap<String, Vec<u8>>,
    extension: &str,
    content_type: &str,
) -> Result<()> {
    let xml = parts
        .entry("[Content_Types].xml".to_string())
        .or_insert_with(|| CONTENT_TYPES_XML.as_bytes().to_vec());
    ensure_xml_fragment(
        xml,
        "</Types>",
        &format!(r#"Extension="{extension}""#),
        &format!(
            r#"  <Default Extension="{extension}" ContentType="{content_type}"/>
"#
        ),
    )
}

fn ensure_header_footer_content_type(
    parts: &mut BTreeMap<String, Vec<u8>>,
    part_name: &str,
    content_type: &str,
) -> Result<()> {
    let xml = parts
        .entry("[Content_Types].xml".to_string())
        .or_insert_with(|| CONTENT_TYPES_XML.as_bytes().to_vec());
    ensure_xml_fragment(
        xml,
        "</Types>",
        part_name,
        &format!(
            r#"  <Override PartName="/{part_name}" ContentType="{content_type}"/>
"#
        ),
    )
}

fn ensure_header_footer_relationship(
    parts: &mut BTreeMap<String, Vec<u8>>,
    relation_id: &str,
    relation_type: &str,
    target: &str,
) -> Result<()> {
    let xml = parts
        .entry("word/_rels/document.xml.rels".to_string())
        .or_insert_with(|| DOCUMENT_RELS_XML.as_bytes().to_vec());
    ensure_xml_fragment(
        xml,
        "</Relationships>",
        relation_id,
        &format!(
            r#"  <Relationship Id="{relation_id}" Type="{relation_type}" Target="{target}"/>
"#
        ),
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
        "docProps/custom.xml".to_string(),
        CUSTOM_XML.as_bytes().to_vec(),
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
    use crate::{DocumentMetadata, Paragraph, Run, Table, TableCell, TableRow};

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
                DocumentBlockRef::Visual(_) => "v",
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
        let content_types = String::from_utf8(parts.parts["[Content_Types].xml"].clone())
            .expect("utf8 content types");
        let rels = String::from_utf8(parts.parts["word/_rels/document.xml.rels"].clone())
            .expect("utf8 relationships");

        assert!(content_types.contains(r#"PartName="/word/numbering.xml""#));
        assert!(rels.contains(r#"Target="styles.xml""#));
        assert!(rels.contains(r#"Target="numbering.xml""#));
    }

    #[test]
    fn metadata_round_trips_through_saved_package() {
        let mut document = sample_document();
        document.set_metadata(
            DocumentMetadata::new()
                .title("Board Report")
                .author("Finance")
                .subject("Q2 review")
                .keyword("board")
                .custom_property("Client", "Acme"),
        );

        let mut buffer = Cursor::new(Vec::new());
        document
            .save_to_writer(&mut buffer)
            .expect("save document with metadata");
        buffer.set_position(0);

        let reopened =
            Document::open_from_reader(&mut buffer, DocumentMode::ReadWrite).expect("reopen");
        assert_eq!(reopened.metadata().title.as_deref(), Some("Board Report"));
        assert_eq!(reopened.metadata().author.as_deref(), Some("Finance"));
        assert_eq!(reopened.metadata().subject.as_deref(), Some("Q2 review"));
        assert_eq!(reopened.metadata().keywords, vec!["board"]);
        assert_eq!(
            reopened
                .metadata()
                .custom_properties
                .get("Client")
                .map(String::as_str),
            Some("Acme")
        );
    }
}
