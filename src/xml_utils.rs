use std::collections::BTreeMap;
use std::io::{BufRead, Cursor, Write};
use std::path::Path;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use crate::document::BodyBlock;
use crate::error::{DocxError, Result};
use crate::layout::{HeaderFooter, PageNumberFormat, PageNumbering, PageSetup};
use crate::paragraph::{Paragraph, ParagraphAlignment, ParagraphList, ParagraphListKind};
use crate::run::{Run, RunProperties, UnderlineStyle, VerticalAlign};
use crate::table::{
    Border, BorderStyle, Table, TableBorders, TableCell, TableCellProperties, TableProperties,
    TableRow, TableRowProperties,
};
use crate::visual::{Visual, VisualFormat, VisualKind};

const WORD_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WP_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
pub(crate) const HEADER_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
pub(crate) const FOOTER_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
pub(crate) const IMAGE_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
pub(crate) const DEFAULT_HEADER_REL_ID: &str = "rIdRusDoxHeaderDefault";
pub(crate) const DEFAULT_FOOTER_REL_ID: &str = "rIdRusDoxFooterDefault";
pub(crate) const DEFAULT_HEADER_PART: &str = "word/header1.xml";
pub(crate) const DEFAULT_FOOTER_PART: &str = "word/footer1.xml";
const EMUS_PER_TWIP: u32 = 635;

/// Internal section settings preserved for the generated document body.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct SectionProperties {
    section_type: String,
    page_setup: PageSetup,
    header: Option<HeaderFooter>,
    footer: Option<HeaderFooter>,
    page_numbering: Option<PageNumbering>,
    header_reference_id: Option<String>,
    footer_reference_id: Option<String>,
}

impl Default for SectionProperties {
    fn default() -> Self {
        Self {
            section_type: "nextPage".to_string(),
            page_setup: PageSetup::default(),
            header: None,
            footer: None,
            page_numbering: None,
            header_reference_id: None,
            footer_reference_id: None,
        }
    }
}

impl SectionProperties {
    pub(crate) fn page_setup(&self) -> &PageSetup {
        &self.page_setup
    }

    pub(crate) fn set_page_setup(&mut self, page_setup: PageSetup) {
        self.page_setup = page_setup;
    }

    pub(crate) fn header(&self) -> Option<&HeaderFooter> {
        self.header.as_ref()
    }

    pub(crate) fn set_header(&mut self, header: Option<HeaderFooter>) {
        self.header = header;
        self.header_reference_id = None;
    }

    pub(crate) fn footer(&self) -> Option<&HeaderFooter> {
        self.footer.as_ref()
    }

    pub(crate) fn set_footer(&mut self, footer: Option<HeaderFooter>) {
        self.footer = footer;
        self.footer_reference_id = None;
    }

    pub(crate) fn page_numbering(&self) -> Option<&PageNumbering> {
        self.page_numbering.as_ref()
    }

    pub(crate) fn set_page_numbering(&mut self, page_numbering: Option<PageNumbering>) {
        self.page_numbering = page_numbering;
    }
}

pub(crate) struct ParsedDocument {
    pub(crate) body: Vec<BodyBlock>,
    pub(crate) section_properties: SectionProperties,
}

type NumberingDefinitions = BTreeMap<u32, ParagraphListKind>;
type VisualRelationships = BTreeMap<String, EmbeddedVisualPart>;

#[derive(Debug, Clone)]
pub(crate) struct DocxVisual {
    pub(crate) relation_id: String,
    pub(crate) width_emu: u32,
    pub(crate) height_emu: u32,
    pub(crate) doc_pr_id: u32,
    pub(crate) name: String,
    pub(crate) alt_text: Option<String>,
}

#[derive(Debug, Clone)]
struct EmbeddedVisualPart {
    format: VisualFormat,
    bytes: Vec<u8>,
}

#[derive(Debug, Default)]
struct ParsedDrawing {
    relation_id: Option<String>,
    width_emu: Option<u32>,
    height_emu: Option<u32>,
    name: Option<String>,
    alt_text: Option<String>,
}

#[derive(Debug, Default)]
struct ParsedRun {
    run: Option<Run>,
    drawing: Option<ParsedDrawing>,
}

pub(crate) fn parse_document_xml(
    xml: &[u8],
    numbering: Option<&NumberingDefinitions>,
    package_parts: &BTreeMap<String, Vec<u8>>,
) -> Result<ParsedDocument> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let visuals = parse_visual_relationships(package_parts)?;

    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if local_name(start.name().as_ref()) == b"body" => {
                let parsed = parse_body(&mut reader, numbering, &visuals)?;
                return Ok(parsed);
            }
            Event::Empty(start) if local_name(start.name().as_ref()) == b"body" => {
                return Ok(ParsedDocument {
                    body: Vec::new(),
                    section_properties: SectionProperties::default(),
                });
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: missing w:body element",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
}

pub(crate) fn write_document_xml(
    body: &[BodyBlock],
    section_properties: &SectionProperties,
    visuals: &[DocxVisual],
) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut document = BytesStart::new("w:document");
    document.push_attribute(("xmlns:w", WORD_NS));
    document.push_attribute(("xmlns:r", REL_NS));
    document.push_attribute(("xmlns:wp", WP_NS));
    document.push_attribute(("xmlns:a", A_NS));
    document.push_attribute(("xmlns:pic", PIC_NS));
    writer.write_event(Event::Start(document))?;

    writer.write_event(Event::Start(BytesStart::new("w:body")))?;
    let mut visual_iter = visuals.iter();
    for block in body {
        match block {
            BodyBlock::Paragraph(paragraph) => write_paragraph(&mut writer, paragraph)?,
            BodyBlock::Table(table) => write_table(&mut writer, table)?,
            BodyBlock::Visual(visual) => {
                let docx_visual = visual_iter.next().ok_or_else(|| {
                    DocxError::parse("missing visual metadata while writing OOXML document")
                })?;
                write_visual_paragraph(&mut writer, visual, docx_visual)?;
            }
        }
    }
    write_section_properties(&mut writer, section_properties)?;
    writer.write_event(Event::End(BytesEnd::new("w:body")))?;
    writer.write_event(Event::End(BytesEnd::new("w:document")))?;
    Ok(writer.into_inner())
}

pub(crate) fn parse_numbering_xml(xml: &[u8]) -> Result<NumberingDefinitions> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut abstract_definitions = BTreeMap::new();
    let mut numbering = BTreeMap::new();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"numbering" => {}
                b"abstractNum" => {
                    let Some(abstract_id) = attribute_value(&start, b"abstractNumId")
                        .and_then(|value| value.parse::<u32>().ok())
                    else {
                        skip_current_element(&mut reader)?;
                        buffer.clear();
                        continue;
                    };

                    if let Some(kind) = parse_abstract_numbering(&mut reader)? {
                        abstract_definitions.insert(abstract_id, kind);
                    }
                }
                b"num" => {
                    let Some(num_id) = attribute_value(&start, b"numId")
                        .and_then(|value| value.parse::<u32>().ok())
                    else {
                        skip_current_element(&mut reader)?;
                        buffer.clear();
                        continue;
                    };

                    if let Some(abstract_id) = parse_numbering_instance(&mut reader)? {
                        if let Some(kind) = abstract_definitions.get(&abstract_id).copied() {
                            numbering.insert(num_id, kind);
                        }
                    }
                }
                _ => {}
            },
            Event::Empty(_) => {}
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(numbering)
}

pub(crate) fn write_numbering_xml(body: &[BodyBlock]) -> Result<Vec<u8>> {
    let numbering = collect_numbering_instances(body)?;
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut numbering_root = BytesStart::new("w:numbering");
    numbering_root.push_attribute(("xmlns:w", WORD_NS));
    writer.write_event(Event::Start(numbering_root))?;

    write_abstract_numbering(&mut writer, 1, ParagraphListKind::Bullet)?;
    write_abstract_numbering(&mut writer, 2, ParagraphListKind::Decimal)?;

    for (num_id, kind) in numbering {
        let mut num = BytesStart::new("w:num");
        let num_id_string = num_id.to_string();
        num.push_attribute(("w:numId", num_id_string.as_str()));
        writer.write_event(Event::Start(num))?;

        let mut abstract_num = BytesStart::new("w:abstractNumId");
        abstract_num.push_attribute((
            "w:val",
            match kind {
                ParagraphListKind::Bullet => "1",
                ParagraphListKind::Decimal => "2",
            },
        ));
        writer.write_event(Event::Empty(abstract_num))?;
        writer.write_event(Event::End(BytesEnd::new("w:num")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("w:numbering")))?;
    Ok(writer.into_inner())
}

fn parse_abstract_numbering<R>(reader: &mut Reader<R>) -> Result<Option<ParagraphListKind>>
where
    R: BufRead,
{
    let mut kind = None;
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"lvl" => {
                    if kind.is_none() {
                        kind = parse_abstract_level(reader)?;
                    } else {
                        skip_current_element(reader)?;
                    }
                }
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => {
                if local_name(start.name().as_ref()) == b"numFmt" && kind.is_none() {
                    kind = attribute_value(&start, b"val")
                        .and_then(|value| ParagraphListKind::from_number_format(&value));
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"abstractNum" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:abstractNum",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(kind)
}

fn parse_abstract_level<R>(reader: &mut Reader<R>) -> Result<Option<ParagraphListKind>>
where
    R: BufRead,
{
    let mut kind = None;
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"numFmt" => {
                    if kind.is_none() {
                        kind = attribute_value(&start, b"val")
                            .and_then(|value| ParagraphListKind::from_number_format(&value));
                    }
                    skip_current_element(reader)?;
                }
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => {
                if local_name(start.name().as_ref()) == b"numFmt" && kind.is_none() {
                    kind = attribute_value(&start, b"val")
                        .and_then(|value| ParagraphListKind::from_number_format(&value));
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"lvl" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:lvl",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(kind)
}

fn parse_numbering_instance<R>(reader: &mut Reader<R>) -> Result<Option<u32>>
where
    R: BufRead,
{
    let mut abstract_id = None;
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"abstractNumId" => {
                    abstract_id =
                        attribute_value(&start, b"val").and_then(|value| value.parse::<u32>().ok());
                    skip_current_element(reader)?;
                }
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => {
                if local_name(start.name().as_ref()) == b"abstractNumId" {
                    abstract_id =
                        attribute_value(&start, b"val").and_then(|value| value.parse().ok());
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"num" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:num",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(abstract_id)
}

fn collect_numbering_instances(body: &[BodyBlock]) -> Result<BTreeMap<u32, ParagraphListKind>> {
    let mut numbering = BTreeMap::new();

    for block in body {
        match block {
            BodyBlock::Paragraph(paragraph) => {
                collect_paragraph_numbering(paragraph, &mut numbering)?
            }
            BodyBlock::Table(table) => {
                for row in table.rows() {
                    for cell in row.cells() {
                        for paragraph in cell.paragraphs() {
                            collect_paragraph_numbering(paragraph, &mut numbering)?;
                        }
                    }
                }
            }
            BodyBlock::Visual(_) => {}
        }
    }

    Ok(numbering)
}

fn collect_paragraph_numbering(
    paragraph: &Paragraph,
    numbering: &mut BTreeMap<u32, ParagraphListKind>,
) -> Result<()> {
    let Some(list) = paragraph.list() else {
        return Ok(());
    };

    if let Some(existing) = numbering.insert(list.id(), list.kind()) {
        if existing != list.kind() {
            return Err(DocxError::parse(format!(
                "conflicting paragraph list kinds for list id {}",
                list.id()
            )));
        }
    }

    Ok(())
}

fn write_abstract_numbering<W>(
    writer: &mut Writer<W>,
    abstract_num_id: u32,
    kind: ParagraphListKind,
) -> Result<()>
where
    W: Write,
{
    let mut abstract_num = BytesStart::new("w:abstractNum");
    let abstract_num_id_string = abstract_num_id.to_string();
    abstract_num.push_attribute(("w:abstractNumId", abstract_num_id_string.as_str()));
    writer.write_event(Event::Start(abstract_num))?;

    let mut multi_level_type = BytesStart::new("w:multiLevelType");
    multi_level_type.push_attribute((
        "w:val",
        match kind {
            ParagraphListKind::Bullet => "hybridMultilevel",
            ParagraphListKind::Decimal => "multilevel",
        },
    ));
    writer.write_event(Event::Empty(multi_level_type))?;

    for level in 0_u8..=8 {
        write_list_level(writer, kind, level)?;
    }

    writer.write_event(Event::End(BytesEnd::new("w:abstractNum")))?;
    Ok(())
}

fn write_list_level<W>(writer: &mut Writer<W>, kind: ParagraphListKind, level: u8) -> Result<()>
where
    W: Write,
{
    let mut level_start = BytesStart::new("w:lvl");
    let level_string = level.to_string();
    level_start.push_attribute(("w:ilvl", level_string.as_str()));
    writer.write_event(Event::Start(level_start))?;

    let mut start = BytesStart::new("w:start");
    start.push_attribute(("w:val", "1"));
    writer.write_event(Event::Empty(start))?;

    let mut num_fmt = BytesStart::new("w:numFmt");
    num_fmt.push_attribute(("w:val", kind.as_number_format()));
    writer.write_event(Event::Empty(num_fmt))?;

    let level_text = list_level_text(kind, level);
    let mut lvl_text = BytesStart::new("w:lvlText");
    lvl_text.push_attribute(("w:val", level_text.as_str()));
    writer.write_event(Event::Empty(lvl_text))?;

    let mut level_jc = BytesStart::new("w:lvlJc");
    level_jc.push_attribute(("w:val", "left"));
    writer.write_event(Event::Empty(level_jc))?;

    writer.write_event(Event::Start(BytesStart::new("w:pPr")))?;
    let mut ind = BytesStart::new("w:ind");
    let left = (720_u32 * (u32::from(level) + 1)).to_string();
    ind.push_attribute(("w:left", left.as_str()));
    ind.push_attribute(("w:hanging", "360"));
    writer.write_event(Event::Empty(ind))?;
    writer.write_event(Event::End(BytesEnd::new("w:pPr")))?;

    writer.write_event(Event::End(BytesEnd::new("w:lvl")))?;
    Ok(())
}

fn list_level_text(kind: ParagraphListKind, level: u8) -> String {
    match kind {
        ParagraphListKind::Bullet => match level % 3 {
            0 => "\u{2022}".to_string(),
            1 => "o".to_string(),
            _ => "\u{25A0}".to_string(),
        },
        ParagraphListKind::Decimal => {
            let mut text = String::new();
            for index in 1..=usize::from(level) + 1 {
                if !text.is_empty() {
                    text.push('.');
                }
                text.push('%');
                text.push_str(&index.to_string());
            }
            text.push('.');
            text
        }
    }
}

fn parse_body<R>(
    reader: &mut Reader<R>,
    numbering: Option<&NumberingDefinitions>,
    visuals: &VisualRelationships,
) -> Result<ParsedDocument>
where
    R: BufRead,
{
    let mut body = Vec::new();
    let mut section_properties = SectionProperties::default();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"p" => body.push(parse_body_paragraph(reader, numbering, visuals)?),
                b"tbl" => body.push(BodyBlock::Table(Box::new(parse_table(reader, numbering)?))),
                b"sectPr" => section_properties = parse_section_properties(reader)?,
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => match local_name(start.name().as_ref()) {
                b"p" => body.push(BodyBlock::Paragraph(Paragraph::new())),
                b"tbl" => body.push(BodyBlock::Table(Box::new(Table::new()))),
                b"sectPr" => {}
                _ => {}
            },
            Event::End(end) if local_name(end.name().as_ref()) == b"body" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:body",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(ParsedDocument {
        body,
        section_properties,
    })
}

fn parse_paragraph<R>(
    reader: &mut Reader<R>,
    numbering: Option<&NumberingDefinitions>,
) -> Result<Paragraph>
where
    R: BufRead,
{
    let (paragraph, _) = parse_paragraph_content(reader, numbering)?;
    Ok(paragraph)
}

fn parse_body_paragraph<R>(
    reader: &mut Reader<R>,
    numbering: Option<&NumberingDefinitions>,
    visuals: &VisualRelationships,
) -> Result<BodyBlock>
where
    R: BufRead,
{
    let (paragraph, drawing) = parse_paragraph_content(reader, numbering)?;
    if paragraph.text().is_empty() {
        if let Some(visual) =
            drawing.and_then(|drawing| visual_from_parsed_drawing(&paragraph, drawing, visuals))
        {
            return Ok(BodyBlock::Visual(visual));
        }
    }

    Ok(BodyBlock::Paragraph(paragraph))
}

fn parse_paragraph_content<R>(
    reader: &mut Reader<R>,
    numbering: Option<&NumberingDefinitions>,
) -> Result<(Paragraph, Option<ParsedDrawing>)>
where
    R: BufRead,
{
    let mut runs = Vec::new();
    let mut drawings = Vec::new();
    let mut list = None;
    let mut alignment = None;
    let mut spacing_before = None;
    let mut spacing_after = None;
    let mut keep_next = false;
    let mut page_break_before = false;
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"r" => {
                    let parsed = parse_run(reader)?;
                    if let Some(run) = parsed.run {
                        runs.push(run);
                    }
                    if let Some(drawing) = parsed.drawing {
                        drawings.push(drawing);
                    }
                }
                b"pPr" => {
                    let properties = parse_paragraph_properties(reader, numbering)?;
                    list = properties.list;
                    alignment = properties.alignment;
                    spacing_before = properties.spacing_before;
                    spacing_after = properties.spacing_after;
                    keep_next = properties.keep_next;
                    page_break_before = properties.page_break_before;
                }
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => {
                if local_name(start.name().as_ref()) == b"r" {
                    runs.push(Run::new());
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"p" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:p",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok((
        Paragraph::from_parts(
            runs,
            list,
            alignment,
            spacing_before,
            spacing_after,
            keep_next,
            page_break_before,
        ),
        if drawings.len() == 1 {
            drawings.into_iter().next()
        } else {
            None
        },
    ))
}

#[derive(Default)]
struct ParsedParagraphProperties {
    list: Option<ParagraphList>,
    alignment: Option<ParagraphAlignment>,
    spacing_before: Option<u32>,
    spacing_after: Option<u32>,
    keep_next: bool,
    page_break_before: bool,
}

fn parse_paragraph_properties<R>(
    reader: &mut Reader<R>,
    numbering: Option<&NumberingDefinitions>,
) -> Result<ParsedParagraphProperties>
where
    R: BufRead,
{
    let mut properties = ParsedParagraphProperties::default();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"numPr" => {
                    properties.list = parse_numbering_properties(reader, numbering)?;
                }
                b"jc" => {
                    if let Some(value) = attribute_value(&start, b"val") {
                        properties.alignment = Some(ParagraphAlignment::from_xml(&value));
                    }
                    skip_current_element(reader)?;
                }
                b"spacing" => {
                    properties.spacing_before = attribute_value(&start, b"before")
                        .and_then(|value| value.parse::<u32>().ok());
                    properties.spacing_after = attribute_value(&start, b"after")
                        .and_then(|value| value.parse::<u32>().ok());
                    skip_current_element(reader)?;
                }
                b"keepNext" => {
                    properties.keep_next = truthy_attribute(&start, b"val").unwrap_or(true);
                    skip_current_element(reader)?;
                }
                b"pageBreakBefore" => {
                    properties.page_break_before = truthy_attribute(&start, b"val").unwrap_or(true);
                    skip_current_element(reader)?;
                }
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => match local_name(start.name().as_ref()) {
                b"numPr" => {
                    properties.list = None;
                }
                b"jc" => {
                    if let Some(value) = attribute_value(&start, b"val") {
                        properties.alignment = Some(ParagraphAlignment::from_xml(&value));
                    }
                }
                b"spacing" => {
                    properties.spacing_before = attribute_value(&start, b"before")
                        .and_then(|value| value.parse::<u32>().ok());
                    properties.spacing_after = attribute_value(&start, b"after")
                        .and_then(|value| value.parse::<u32>().ok());
                }
                b"keepNext" => {
                    properties.keep_next = truthy_attribute(&start, b"val").unwrap_or(true);
                }
                b"pageBreakBefore" => {
                    properties.page_break_before = truthy_attribute(&start, b"val").unwrap_or(true);
                }
                _ => {}
            },
            Event::End(end) if local_name(end.name().as_ref()) == b"pPr" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:pPr",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(properties)
}

fn parse_numbering_properties<R>(
    reader: &mut Reader<R>,
    numbering: Option<&NumberingDefinitions>,
) -> Result<Option<ParagraphList>>
where
    R: BufRead,
{
    let mut level = 0_u8;
    let mut num_id = None;
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"ilvl" => {
                    level = attribute_value(&start, b"val")
                        .and_then(|value| value.parse::<u8>().ok())
                        .unwrap_or(0);
                    skip_current_element(reader)?;
                }
                b"numId" => {
                    num_id = attribute_value(&start, b"val").and_then(|value| value.parse().ok());
                    skip_current_element(reader)?;
                }
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => match local_name(start.name().as_ref()) {
                b"ilvl" => {
                    level = attribute_value(&start, b"val")
                        .and_then(|value| value.parse::<u8>().ok())
                        .unwrap_or(0);
                }
                b"numId" => {
                    num_id = attribute_value(&start, b"val").and_then(|value| value.parse().ok());
                }
                _ => {}
            },
            Event::End(end) if local_name(end.name().as_ref()) == b"numPr" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:numPr",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(num_id.and_then(|id| {
        numbering
            .and_then(|definitions| definitions.get(&id).copied())
            .map(|kind| ParagraphList::from_parts(kind, id, level))
    }))
}

fn parse_run<R>(reader: &mut Reader<R>) -> Result<ParsedRun>
where
    R: BufRead,
{
    let mut text = String::new();
    let mut properties = RunProperties::default();
    let mut drawing = None;
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"rPr" => properties = parse_run_properties(reader)?,
                b"t" => text.push_str(&parse_text_element(reader)?),
                b"tab" => {
                    text.push('\t');
                    skip_current_element(reader)?;
                }
                b"br" => {
                    text.push('\n');
                    skip_current_element(reader)?;
                }
                b"drawing" => drawing = parse_drawing(reader)?,
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => match local_name(start.name().as_ref()) {
                b"tab" => text.push('\t'),
                b"br" => text.push('\n'),
                _ => {}
            },
            Event::End(end) if local_name(end.name().as_ref()) == b"r" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:r",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(ParsedRun {
        run: Some(Run::from_parts(text, properties)),
        drawing,
    })
}

fn parse_drawing<R>(reader: &mut Reader<R>) -> Result<Option<ParsedDrawing>>
where
    R: BufRead,
{
    let mut drawing = ParsedDrawing::default();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"inline" | b"anchor" => parse_inline_drawing(reader, &mut drawing)?,
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => match local_name(start.name().as_ref()) {
                b"inline" | b"anchor" => {}
                _ => {}
            },
            Event::End(end) if local_name(end.name().as_ref()) == b"drawing" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:drawing",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(drawing.relation_id.is_some().then_some(drawing))
}

fn parse_inline_drawing<R>(reader: &mut Reader<R>, drawing: &mut ParsedDrawing) -> Result<()>
where
    R: BufRead,
{
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"extent" => {
                    drawing.width_emu =
                        attribute_value(&start, b"cx").and_then(|value| value.parse().ok());
                    drawing.height_emu =
                        attribute_value(&start, b"cy").and_then(|value| value.parse().ok());
                    skip_current_element(reader)?;
                }
                b"docPr" => {
                    drawing.name = attribute_value(&start, b"name");
                    drawing.alt_text = attribute_value(&start, b"descr");
                    skip_current_element(reader)?;
                }
                b"blip" => {
                    drawing.relation_id = attribute_value(&start, b"embed");
                    skip_current_element(reader)?;
                }
                _ => {}
            },
            Event::Empty(start) => match local_name(start.name().as_ref()) {
                b"extent" => {
                    drawing.width_emu =
                        attribute_value(&start, b"cx").and_then(|value| value.parse().ok());
                    drawing.height_emu =
                        attribute_value(&start, b"cy").and_then(|value| value.parse().ok());
                }
                b"docPr" => {
                    drawing.name = attribute_value(&start, b"name");
                    drawing.alt_text = attribute_value(&start, b"descr");
                }
                b"blip" => {
                    drawing.relation_id = attribute_value(&start, b"embed");
                }
                _ => {}
            },
            Event::End(end) if matches!(local_name(end.name().as_ref()), b"inline" | b"anchor") => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in inline drawing",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(())
}

fn parse_run_properties<R>(reader: &mut Reader<R>) -> Result<RunProperties>
where
    R: BufRead,
{
    let mut properties = RunProperties::default();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                apply_run_property(&mut properties, &start);
                skip_current_element(reader)?;
            }
            Event::Empty(start) => apply_run_property(&mut properties, &start),
            Event::End(end) if local_name(end.name().as_ref()) == b"rPr" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:rPr",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(properties)
}

fn apply_run_property(properties: &mut RunProperties, start: &BytesStart<'_>) {
    match local_name(start.name().as_ref()) {
        b"b" => properties.bold = truthy_attribute(start, b"val").unwrap_or(true),
        b"i" => properties.italic = truthy_attribute(start, b"val").unwrap_or(true),
        b"u" => {
            let style = attribute_value(start, b"val")
                .map(|value| UnderlineStyle::from_xml(&value))
                .unwrap_or(UnderlineStyle::Single);
            properties.underline = Some(style);
        }
        b"strike" => properties.strikethrough = truthy_attribute(start, b"val").unwrap_or(true),
        b"smallCaps" => properties.small_caps = truthy_attribute(start, b"val").unwrap_or(true),
        b"shadow" => properties.shadow = truthy_attribute(start, b"val").unwrap_or(true),
        b"color" => properties.color = attribute_value(start, b"val"),
        b"sz" => {
            properties.font_size =
                attribute_value(start, b"val").and_then(|value| value.parse::<u16>().ok());
        }
        b"rFonts" => {
            properties.font_family = attribute_value(start, b"ascii")
                .or_else(|| attribute_value(start, b"hAnsi"))
                .or_else(|| attribute_value(start, b"cs"));
        }
        b"vertAlign" => {
            if let Some(value) = attribute_value(start, b"val") {
                properties.vertical_align = Some(VerticalAlign::from_xml(&value));
            }
        }
        _ => {}
    }
}

fn parse_text_element<R>(reader: &mut Reader<R>) -> Result<String>
where
    R: BufRead,
{
    let mut text = String::new();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Text(value) => {
                let raw = String::from_utf8_lossy(value.as_ref());
                text.push_str(&quick_xml::escape::unescape(&raw)?);
            }
            Event::CData(value) => {
                text.push_str(&String::from_utf8_lossy(value.as_ref()));
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"t" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:t",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(text)
}

fn visual_from_parsed_drawing(
    paragraph: &Paragraph,
    drawing: ParsedDrawing,
    visuals: &VisualRelationships,
) -> Option<Visual> {
    let relation_id = drawing.relation_id?;
    let embedded = visuals.get(&relation_id)?;
    let mut visual = Visual::from_bytes(embedded.bytes.clone(), embedded.format).with_kind(
        drawing
            .name
            .as_deref()
            .map(VisualKind::from_docx_name)
            .unwrap_or(VisualKind::Image),
    );
    visual = visual.with_alignment(
        paragraph
            .alignment()
            .cloned()
            .unwrap_or(ParagraphAlignment::Left),
    );
    if let Some(alt_text) = drawing.alt_text {
        if !alt_text.trim().is_empty() {
            visual = visual.alt_text_text(alt_text);
        }
    }
    if let Some(width_emu) = drawing.width_emu {
        visual = visual.width_twips(width_emu / EMUS_PER_TWIP);
    }
    if let Some(height_emu) = drawing.height_emu {
        visual = visual.height_twips(height_emu / EMUS_PER_TWIP);
    }
    Some(visual)
}

fn parse_visual_relationships(
    package_parts: &BTreeMap<String, Vec<u8>>,
) -> Result<VisualRelationships> {
    let Some(rels_xml) = package_parts.get("word/_rels/document.xml.rels") else {
        return Ok(BTreeMap::new());
    };

    let mut reader = Reader::from_reader(Cursor::new(rels_xml));
    reader.config_mut().trim_text(false);
    let mut visuals = BTreeMap::new();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if local_name(start.name().as_ref()) == b"Relationship" => {
                maybe_insert_visual_relationship(&mut visuals, package_parts, &start)?;
                skip_current_element(&mut reader)?;
            }
            Event::Empty(start) if local_name(start.name().as_ref()) == b"Relationship" => {
                maybe_insert_visual_relationship(&mut visuals, package_parts, &start)?;
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(visuals)
}

fn relationship_target_to_part_name(target: &str) -> String {
    if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else if target.starts_with("word/") {
        target.to_string()
    } else {
        format!("word/{target}")
    }
}

fn maybe_insert_visual_relationship(
    visuals: &mut VisualRelationships,
    package_parts: &BTreeMap<String, Vec<u8>>,
    start: &BytesStart<'_>,
) -> Result<()> {
    let Some(relation_type) = attribute_value(start, b"Type") else {
        return Ok(());
    };
    if relation_type != IMAGE_REL_TYPE {
        return Ok(());
    }

    let Some(relation_id) = attribute_value(start, b"Id") else {
        return Ok(());
    };
    let Some(target) = attribute_value(start, b"Target") else {
        return Ok(());
    };

    let part_name = relationship_target_to_part_name(&target);
    let Some(bytes) = package_parts.get(&part_name) else {
        return Ok(());
    };
    let format = VisualFormat::from_path(Path::new(&part_name))
        .or_else(|| VisualFormat::guess(bytes))
        .ok_or_else(|| {
            DocxError::parse(format!(
                "unsupported embedded image format in OOXML part {part_name}"
            ))
        })?;
    visuals.insert(
        relation_id,
        EmbeddedVisualPart {
            format,
            bytes: bytes.clone(),
        },
    );
    Ok(())
}

fn parse_table<R>(reader: &mut Reader<R>, numbering: Option<&NumberingDefinitions>) -> Result<Table>
where
    R: BufRead,
{
    let mut rows = Vec::new();
    let mut properties = TableProperties::default();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"tblPr" => properties = parse_table_properties(reader)?,
                b"tr" => rows.push(parse_table_row(reader, numbering)?),
                _ => skip_current_element(reader)?,
            },
            Event::End(end) if local_name(end.name().as_ref()) == b"tbl" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:tbl",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(Table::from_parts(rows, properties))
}

fn parse_table_properties<R>(reader: &mut Reader<R>) -> Result<TableProperties>
where
    R: BufRead,
{
    let mut properties = TableProperties::default();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"tblW" => {
                    properties.width =
                        attribute_value(&start, b"w").and_then(|value| value.parse::<u32>().ok());
                    skip_current_element(reader)?;
                }
                b"tblBorders" => properties.borders = Some(parse_borders(reader, b"tblBorders")?),
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => match local_name(start.name().as_ref()) {
                b"tblW" => {
                    properties.width =
                        attribute_value(&start, b"w").and_then(|value| value.parse::<u32>().ok());
                }
                b"tblBorders" => properties.borders = Some(TableBorders::default()),
                _ => {}
            },
            Event::End(end) if local_name(end.name().as_ref()) == b"tblPr" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:tblPr",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(properties)
}

fn parse_table_row<R>(
    reader: &mut Reader<R>,
    numbering: Option<&NumberingDefinitions>,
) -> Result<TableRow>
where
    R: BufRead,
{
    let mut cells = Vec::new();
    let mut properties = TableRowProperties::default();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"trPr" => properties = parse_table_row_properties(reader)?,
                b"tc" => cells.push(parse_table_cell(reader, numbering)?),
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => {
                if local_name(start.name().as_ref()) == b"trPr" {
                    properties = TableRowProperties::default();
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"tr" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:tr",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(TableRow::from_parts(cells, properties))
}

fn parse_table_row_properties<R>(reader: &mut Reader<R>) -> Result<TableRowProperties>
where
    R: BufRead,
{
    let mut properties = TableRowProperties::default();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"tblHeader" => {
                    properties.repeat_as_header = true;
                    skip_current_element(reader)?;
                }
                b"cantSplit" => {
                    properties.allow_split_across_pages = false;
                    skip_current_element(reader)?;
                }
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => match local_name(start.name().as_ref()) {
                b"tblHeader" => properties.repeat_as_header = true,
                b"cantSplit" => properties.allow_split_across_pages = false,
                _ => {}
            },
            Event::End(end) if local_name(end.name().as_ref()) == b"trPr" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:trPr",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(properties)
}

fn parse_table_cell<R>(
    reader: &mut Reader<R>,
    numbering: Option<&NumberingDefinitions>,
) -> Result<TableCell>
where
    R: BufRead,
{
    let mut paragraphs = Vec::new();
    let mut properties = TableCellProperties::default();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"tcPr" => properties = parse_table_cell_properties(reader)?,
                b"p" => paragraphs.push(parse_paragraph(reader, numbering)?),
                b"tbl" => {
                    skip_current_element(reader)?;
                }
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => {
                if local_name(start.name().as_ref()) == b"p" {
                    paragraphs.push(Paragraph::new());
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"tc" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:tc",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(TableCell::from_parts(paragraphs, properties))
}

fn parse_table_cell_properties<R>(reader: &mut Reader<R>) -> Result<TableCellProperties>
where
    R: BufRead,
{
    let mut properties = TableCellProperties::default();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"tcW" => {
                    properties.width =
                        attribute_value(&start, b"w").and_then(|value| value.parse::<u32>().ok());
                    skip_current_element(reader)?;
                }
                b"gridSpan" => {
                    properties.grid_span =
                        attribute_value(&start, b"val").and_then(|value| value.parse::<u32>().ok());
                    skip_current_element(reader)?;
                }
                b"tcBorders" => properties.borders = Some(parse_borders(reader, b"tcBorders")?),
                b"shd" => {
                    properties.background_color = attribute_value(&start, b"fill");
                    skip_current_element(reader)?;
                }
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => match local_name(start.name().as_ref()) {
                b"tcW" => {
                    properties.width =
                        attribute_value(&start, b"w").and_then(|value| value.parse::<u32>().ok());
                }
                b"gridSpan" => {
                    properties.grid_span =
                        attribute_value(&start, b"val").and_then(|value| value.parse::<u32>().ok());
                }
                b"tcBorders" => properties.borders = Some(TableBorders::default()),
                b"shd" => {
                    properties.background_color = attribute_value(&start, b"fill");
                }
                _ => {}
            },
            Event::End(end) if local_name(end.name().as_ref()) == b"tcPr" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:tcPr",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(properties)
}

fn parse_borders<R>(reader: &mut Reader<R>, end_tag: &[u8]) -> Result<TableBorders>
where
    R: BufRead,
{
    let mut borders = TableBorders::default();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                assign_border(&mut borders, &start);
                skip_current_element(reader)?;
            }
            Event::Empty(start) => assign_border(&mut borders, &start),
            Event::End(end) if local_name(end.name().as_ref()) == end_tag => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in border properties",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(borders)
}

fn assign_border(borders: &mut TableBorders, start: &BytesStart<'_>) {
    let border = Border {
        style: attribute_value(start, b"val")
            .map(|value| BorderStyle::from_xml(&value))
            .unwrap_or(BorderStyle::Single),
        size: attribute_value(start, b"sz").and_then(|value| value.parse::<u16>().ok()),
        color: attribute_value(start, b"color"),
    };

    match local_name(start.name().as_ref()) {
        b"top" => borders.top = Some(border),
        b"bottom" => borders.bottom = Some(border),
        b"left" => borders.left = Some(border),
        b"right" => borders.right = Some(border),
        b"insideH" => borders.inside_horizontal = Some(border),
        b"insideV" => borders.inside_vertical = Some(border),
        _ => {}
    }
}

fn parse_section_properties<R>(reader: &mut Reader<R>) -> Result<SectionProperties>
where
    R: BufRead,
{
    let mut section = SectionProperties::default();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"type" => {
                    if let Some(value) = attribute_value(&start, b"val") {
                        section.section_type = value;
                    }
                    skip_current_element(reader)?;
                }
                b"headerReference" => {
                    if let Some(value) = attribute_value(&start, b"id") {
                        section.header_reference_id = Some(value);
                    }
                    skip_current_element(reader)?;
                }
                b"footerReference" => {
                    if let Some(value) = attribute_value(&start, b"id") {
                        section.footer_reference_id = Some(value);
                    }
                    skip_current_element(reader)?;
                }
                b"pgSz" => {
                    if let Some(value) =
                        attribute_value(&start, b"w").and_then(|value| value.parse().ok())
                    {
                        section.page_setup.width_twips = value;
                    }
                    if let Some(value) =
                        attribute_value(&start, b"h").and_then(|value| value.parse().ok())
                    {
                        section.page_setup.height_twips = value;
                    }
                    skip_current_element(reader)?;
                }
                b"pgMar" => {
                    parse_page_margins(&mut section, &start);
                    skip_current_element(reader)?;
                }
                b"pgNumType" => {
                    section.page_numbering = Some(parse_page_numbering(&start));
                    skip_current_element(reader)?;
                }
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => match local_name(start.name().as_ref()) {
                b"type" => {
                    if let Some(value) = attribute_value(&start, b"val") {
                        section.section_type = value;
                    }
                }
                b"headerReference" => {
                    if let Some(value) = attribute_value(&start, b"id") {
                        section.header_reference_id = Some(value);
                    }
                }
                b"footerReference" => {
                    if let Some(value) = attribute_value(&start, b"id") {
                        section.footer_reference_id = Some(value);
                    }
                }
                b"pgSz" => {
                    if let Some(value) =
                        attribute_value(&start, b"w").and_then(|value| value.parse().ok())
                    {
                        section.page_setup.width_twips = value;
                    }
                    if let Some(value) =
                        attribute_value(&start, b"h").and_then(|value| value.parse().ok())
                    {
                        section.page_setup.height_twips = value;
                    }
                }
                b"pgMar" => parse_page_margins(&mut section, &start),
                b"pgNumType" => section.page_numbering = Some(parse_page_numbering(&start)),
                _ => {}
            },
            Event::End(end) if local_name(end.name().as_ref()) == b"sectPr" => {
                break;
            }
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file in w:sectPr",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(section)
}

fn parse_page_margins(section: &mut SectionProperties, start: &BytesStart<'_>) {
    if let Some(value) = attribute_value(start, b"top").and_then(|value| value.parse().ok()) {
        section.page_setup.margin_top_twips = value;
    }
    if let Some(value) = attribute_value(start, b"right").and_then(|value| value.parse().ok()) {
        section.page_setup.margin_right_twips = value;
    }
    if let Some(value) = attribute_value(start, b"bottom").and_then(|value| value.parse().ok()) {
        section.page_setup.margin_bottom_twips = value;
    }
    if let Some(value) = attribute_value(start, b"left").and_then(|value| value.parse().ok()) {
        section.page_setup.margin_left_twips = value;
    }
    if let Some(value) = attribute_value(start, b"header").and_then(|value| value.parse().ok()) {
        section.page_setup.header_twips = value;
    }
    if let Some(value) = attribute_value(start, b"footer").and_then(|value| value.parse().ok()) {
        section.page_setup.footer_twips = value;
    }
    if let Some(value) = attribute_value(start, b"gutter").and_then(|value| value.parse().ok()) {
        section.page_setup.gutter_twips = value;
    }
}

fn parse_page_numbering(start: &BytesStart<'_>) -> PageNumbering {
    PageNumbering {
        start_at: attribute_value(start, b"start").and_then(|value| value.parse().ok()),
        format: attribute_value(start, b"fmt")
            .map(|value| PageNumberFormat::from_xml(&value))
            .unwrap_or(PageNumberFormat::Decimal),
    }
}

fn write_paragraph<W>(writer: &mut Writer<W>, paragraph: &Paragraph) -> Result<()>
where
    W: Write,
{
    writer.write_event(Event::Start(BytesStart::new("w:p")))?;
    if paragraph.has_properties() {
        writer.write_event(Event::Start(BytesStart::new("w:pPr")))?;
        if let Some(list) = paragraph.list() {
            writer.write_event(Event::Start(BytesStart::new("w:numPr")))?;

            let mut level = BytesStart::new("w:ilvl");
            let level_string = list.level().to_string();
            level.push_attribute(("w:val", level_string.as_str()));
            writer.write_event(Event::Empty(level))?;

            let mut num_id = BytesStart::new("w:numId");
            let num_id_string = list.id().to_string();
            num_id.push_attribute(("w:val", num_id_string.as_str()));
            writer.write_event(Event::Empty(num_id))?;

            writer.write_event(Event::End(BytesEnd::new("w:numPr")))?;
        }
        if let Some(alignment) = paragraph.alignment() {
            let mut start = BytesStart::new("w:jc");
            start.push_attribute(("w:val", alignment.as_xml_value()));
            writer.write_event(Event::Empty(start))?;
        }
        if paragraph.spacing_before_value().is_some() || paragraph.spacing_after_value().is_some() {
            let mut start = BytesStart::new("w:spacing");
            let spacing_before = paragraph
                .spacing_before_value()
                .map(|value| value.to_string());
            let spacing_after = paragraph
                .spacing_after_value()
                .map(|value| value.to_string());
            if let Some(spacing_before) = spacing_before.as_deref() {
                start.push_attribute(("w:before", spacing_before));
            }
            if let Some(spacing_after) = spacing_after.as_deref() {
                start.push_attribute(("w:after", spacing_after));
            }
            writer.write_event(Event::Empty(start))?;
        }
        if paragraph.has_keep_next() {
            writer.write_event(Event::Empty(BytesStart::new("w:keepNext")))?;
        }
        if paragraph.has_page_break_before() {
            writer.write_event(Event::Empty(BytesStart::new("w:pageBreakBefore")))?;
        }
        writer.write_event(Event::End(BytesEnd::new("w:pPr")))?;
    }
    for run in paragraph.runs() {
        write_run(writer, run)?;
    }
    writer.write_event(Event::End(BytesEnd::new("w:p")))?;
    Ok(())
}

fn write_visual_paragraph<W>(
    writer: &mut Writer<W>,
    visual: &Visual,
    docx_visual: &DocxVisual,
) -> Result<()>
where
    W: Write,
{
    writer.write_event(Event::Start(BytesStart::new("w:p")))?;
    if !matches!(visual.alignment(), ParagraphAlignment::Left) {
        writer.write_event(Event::Start(BytesStart::new("w:pPr")))?;
        let mut alignment = BytesStart::new("w:jc");
        alignment.push_attribute(("w:val", visual.alignment().as_xml_value()));
        writer.write_event(Event::Empty(alignment))?;
        writer.write_event(Event::End(BytesEnd::new("w:pPr")))?;
    }

    writer.write_event(Event::Start(BytesStart::new("w:r")))?;
    writer.write_event(Event::Start(BytesStart::new("w:drawing")))?;

    let mut inline = BytesStart::new("wp:inline");
    inline.push_attribute(("distT", "0"));
    inline.push_attribute(("distB", "0"));
    inline.push_attribute(("distL", "0"));
    inline.push_attribute(("distR", "0"));
    writer.write_event(Event::Start(inline))?;

    let mut extent = BytesStart::new("wp:extent");
    let width = docx_visual.width_emu.to_string();
    let height = docx_visual.height_emu.to_string();
    extent.push_attribute(("cx", width.as_str()));
    extent.push_attribute(("cy", height.as_str()));
    writer.write_event(Event::Empty(extent))?;

    let mut doc_pr = BytesStart::new("wp:docPr");
    let doc_pr_id = docx_visual.doc_pr_id.to_string();
    doc_pr.push_attribute(("id", doc_pr_id.as_str()));
    doc_pr.push_attribute(("name", docx_visual.name.as_str()));
    if let Some(alt_text) = docx_visual.alt_text.as_deref() {
        doc_pr.push_attribute(("descr", alt_text));
    }
    writer.write_event(Event::Empty(doc_pr))?;

    writer.write_event(Event::Start(BytesStart::new("wp:cNvGraphicFramePr")))?;
    let mut locks = BytesStart::new("a:graphicFrameLocks");
    locks.push_attribute(("noChangeAspect", "1"));
    writer.write_event(Event::Empty(locks))?;
    writer.write_event(Event::End(BytesEnd::new("wp:cNvGraphicFramePr")))?;

    writer.write_event(Event::Start(BytesStart::new("a:graphic")))?;
    let mut graphic_data = BytesStart::new("a:graphicData");
    graphic_data.push_attribute(("uri", PIC_NS));
    writer.write_event(Event::Start(graphic_data))?;

    writer.write_event(Event::Start(BytesStart::new("pic:pic")))?;
    writer.write_event(Event::Start(BytesStart::new("pic:nvPicPr")))?;

    let mut c_nv_pr = BytesStart::new("pic:cNvPr");
    c_nv_pr.push_attribute(("id", "0"));
    c_nv_pr.push_attribute(("name", docx_visual.name.as_str()));
    if let Some(alt_text) = docx_visual.alt_text.as_deref() {
        c_nv_pr.push_attribute(("descr", alt_text));
    }
    writer.write_event(Event::Empty(c_nv_pr))?;
    writer.write_event(Event::Empty(BytesStart::new("pic:cNvPicPr")))?;
    writer.write_event(Event::End(BytesEnd::new("pic:nvPicPr")))?;

    writer.write_event(Event::Start(BytesStart::new("pic:blipFill")))?;
    let mut blip = BytesStart::new("a:blip");
    blip.push_attribute(("r:embed", docx_visual.relation_id.as_str()));
    writer.write_event(Event::Empty(blip))?;
    writer.write_event(Event::Start(BytesStart::new("a:stretch")))?;
    writer.write_event(Event::Empty(BytesStart::new("a:fillRect")))?;
    writer.write_event(Event::End(BytesEnd::new("a:stretch")))?;
    writer.write_event(Event::End(BytesEnd::new("pic:blipFill")))?;

    writer.write_event(Event::Start(BytesStart::new("pic:spPr")))?;
    writer.write_event(Event::Start(BytesStart::new("a:xfrm")))?;
    let mut offset = BytesStart::new("a:off");
    offset.push_attribute(("x", "0"));
    offset.push_attribute(("y", "0"));
    writer.write_event(Event::Empty(offset))?;
    let mut ext = BytesStart::new("a:ext");
    ext.push_attribute(("cx", width.as_str()));
    ext.push_attribute(("cy", height.as_str()));
    writer.write_event(Event::Empty(ext))?;
    writer.write_event(Event::End(BytesEnd::new("a:xfrm")))?;
    let mut geometry = BytesStart::new("a:prstGeom");
    geometry.push_attribute(("prst", "rect"));
    writer.write_event(Event::Start(geometry))?;
    writer.write_event(Event::Empty(BytesStart::new("a:avLst")))?;
    writer.write_event(Event::End(BytesEnd::new("a:prstGeom")))?;
    writer.write_event(Event::End(BytesEnd::new("pic:spPr")))?;

    writer.write_event(Event::End(BytesEnd::new("pic:pic")))?;
    writer.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
    writer.write_event(Event::End(BytesEnd::new("a:graphic")))?;
    writer.write_event(Event::End(BytesEnd::new("wp:inline")))?;
    writer.write_event(Event::End(BytesEnd::new("w:drawing")))?;
    writer.write_event(Event::End(BytesEnd::new("w:r")))?;
    writer.write_event(Event::End(BytesEnd::new("w:p")))?;
    Ok(())
}

fn write_run<W>(writer: &mut Writer<W>, run: &Run) -> Result<()>
where
    W: Write,
{
    writer.write_event(Event::Start(BytesStart::new("w:r")))?;
    if run.properties().has_serialized_content() {
        writer.write_event(Event::Start(BytesStart::new("w:rPr")))?;
        if run.properties().bold {
            writer.write_event(Event::Empty(BytesStart::new("w:b")))?;
        }
        if run.properties().italic {
            writer.write_event(Event::Empty(BytesStart::new("w:i")))?;
        }
        if let Some(underline) = &run.properties().underline {
            let mut start = BytesStart::new("w:u");
            start.push_attribute(("w:val", underline.as_xml_value()));
            writer.write_event(Event::Empty(start))?;
        }
        if run.properties().strikethrough {
            writer.write_event(Event::Empty(BytesStart::new("w:strike")))?;
        }
        if run.properties().small_caps {
            writer.write_event(Event::Empty(BytesStart::new("w:smallCaps")))?;
        }
        if run.properties().shadow {
            writer.write_event(Event::Empty(BytesStart::new("w:shadow")))?;
        }
        if let Some(color) = &run.properties().color {
            let mut start = BytesStart::new("w:color");
            start.push_attribute(("w:val", color.as_str()));
            writer.write_event(Event::Empty(start))?;
        }
        if let Some(font_size) = run.properties().font_size {
            let font_size = font_size.to_string();
            let mut start = BytesStart::new("w:sz");
            start.push_attribute(("w:val", font_size.as_str()));
            writer.write_event(Event::Empty(start))?;
            let mut complex_start = BytesStart::new("w:szCs");
            complex_start.push_attribute(("w:val", font_size.as_str()));
            writer.write_event(Event::Empty(complex_start))?;
        }
        if let Some(font_family) = &run.properties().font_family {
            let mut start = BytesStart::new("w:rFonts");
            start.push_attribute(("w:ascii", font_family.as_str()));
            start.push_attribute(("w:hAnsi", font_family.as_str()));
            start.push_attribute(("w:cs", font_family.as_str()));
            writer.write_event(Event::Empty(start))?;
        }
        if let Some(vertical_align) = run.properties().vertical_align {
            let mut start = BytesStart::new("w:vertAlign");
            start.push_attribute(("w:val", vertical_align.as_xml_value()));
            writer.write_event(Event::Empty(start))?;
        }
        writer.write_event(Event::End(BytesEnd::new("w:rPr")))?;
    }

    write_run_content(writer, run)?;
    writer.write_event(Event::End(BytesEnd::new("w:r")))?;
    Ok(())
}

fn write_run_content<W>(writer: &mut Writer<W>, run: &Run) -> Result<()>
where
    W: Write,
{
    let text = run.text();
    if !text.contains(['\t', '\n', '\r']) {
        if text.is_empty() {
            writer.write_event(Event::Start(BytesStart::new("w:t")))?;
            writer.write_event(Event::End(BytesEnd::new("w:t")))?;
        } else {
            write_text_segment_with_preserve(writer, text, run.needs_space_preserve())?;
        }
        return Ok(());
    }

    let mut emitted_content = false;
    let mut segment_start = 0usize;
    let mut chars = text.char_indices().peekable();

    while let Some((index, ch)) = chars.next() {
        match ch {
            '\t' | '\n' | '\r' => {
                if segment_start < index {
                    write_text_segment(writer, &text[segment_start..index])?;
                }

                match ch {
                    '\t' => writer.write_event(Event::Empty(BytesStart::new("w:tab")))?,
                    '\n' | '\r' => {
                        writer.write_event(Event::Empty(BytesStart::new("w:br")))?;
                        if ch == '\r' && chars.next_if(|(_, next)| *next == '\n').is_some() {
                            // Treat CRLF as a single OOXML break.
                        }
                    }
                    _ => {}
                }

                emitted_content = true;
                segment_start = match ch {
                    '\r' => chars
                        .peek()
                        .map(|(next_index, _)| *next_index)
                        .unwrap_or(text.len()),
                    _ => index + ch.len_utf8(),
                };
            }
            _ => {}
        }
    }

    if segment_start < text.len() {
        write_text_segment(writer, &text[segment_start..])?;
        emitted_content = true;
    }

    if !emitted_content {
        writer.write_event(Event::Start(BytesStart::new("w:t")))?;
        writer.write_event(Event::End(BytesEnd::new("w:t")))?;
    }

    Ok(())
}

fn write_text_segment<W>(writer: &mut Writer<W>, text: &str) -> Result<()>
where
    W: Write,
{
    write_text_segment_with_preserve(writer, text, text_segment_needs_space_preserve(text))
}

fn write_text_segment_with_preserve<W>(
    writer: &mut Writer<W>,
    text: &str,
    preserve_space: bool,
) -> Result<()>
where
    W: Write,
{
    let mut element = BytesStart::new("w:t");
    if preserve_space {
        element.push_attribute(("xml:space", "preserve"));
    }
    writer.write_event(Event::Start(element))?;
    writer.write_event(Event::Text(BytesText::new(text)))?;
    writer.write_event(Event::End(BytesEnd::new("w:t")))?;
    Ok(())
}

fn text_segment_needs_space_preserve(text: &str) -> bool {
    let starts_with_ws = text.chars().next().is_some_and(char::is_whitespace);
    let ends_with_ws = text.chars().last().is_some_and(char::is_whitespace);
    starts_with_ws || ends_with_ws
}

fn write_table<W>(writer: &mut Writer<W>, table: &Table) -> Result<()>
where
    W: Write,
{
    writer.write_event(Event::Start(BytesStart::new("w:tbl")))?;
    if table.properties().width.is_some() || table.properties().borders.is_some() {
        writer.write_event(Event::Start(BytesStart::new("w:tblPr")))?;
        if let Some(width) = table.properties().width {
            let mut start = BytesStart::new("w:tblW");
            let width_string = width.to_string();
            start.push_attribute(("w:type", "dxa"));
            start.push_attribute(("w:w", width_string.as_str()));
            writer.write_event(Event::Empty(start))?;
        }
        if let Some(borders) = &table.properties().borders {
            write_borders(writer, "w:tblBorders", borders)?;
        }
        writer.write_event(Event::End(BytesEnd::new("w:tblPr")))?;
    }
    for row in table.rows() {
        writer.write_event(Event::Start(BytesStart::new("w:tr")))?;
        if row.properties().has_serialized_content() {
            writer.write_event(Event::Start(BytesStart::new("w:trPr")))?;
            if row.properties().repeat_as_header {
                writer.write_event(Event::Empty(BytesStart::new("w:tblHeader")))?;
            }
            if !row.properties().allow_split_across_pages {
                writer.write_event(Event::Empty(BytesStart::new("w:cantSplit")))?;
            }
            writer.write_event(Event::End(BytesEnd::new("w:trPr")))?;
        }
        for cell in row.cells() {
            write_table_cell(writer, cell)?;
        }
        writer.write_event(Event::End(BytesEnd::new("w:tr")))?;
    }
    writer.write_event(Event::End(BytesEnd::new("w:tbl")))?;
    Ok(())
}

fn write_table_cell<W>(writer: &mut Writer<W>, cell: &TableCell) -> Result<()>
where
    W: Write,
{
    writer.write_event(Event::Start(BytesStart::new("w:tc")))?;
    if cell.properties().width.is_some()
        || cell.properties().grid_span.is_some()
        || cell.properties().borders.is_some()
        || cell.properties().background_color.is_some()
    {
        writer.write_event(Event::Start(BytesStart::new("w:tcPr")))?;
        if let Some(width) = cell.properties().width {
            let mut start = BytesStart::new("w:tcW");
            let width_string = width.to_string();
            start.push_attribute(("w:type", "dxa"));
            start.push_attribute(("w:w", width_string.as_str()));
            writer.write_event(Event::Empty(start))?;
        }
        if let Some(grid_span) = cell.properties().grid_span {
            let mut start = BytesStart::new("w:gridSpan");
            let grid_span_string = grid_span.to_string();
            start.push_attribute(("w:val", grid_span_string.as_str()));
            writer.write_event(Event::Empty(start))?;
        }
        if let Some(borders) = &cell.properties().borders {
            write_borders(writer, "w:tcBorders", borders)?;
        }
        if let Some(background_color) = &cell.properties().background_color {
            let mut start = BytesStart::new("w:shd");
            start.push_attribute(("w:val", "clear"));
            start.push_attribute(("w:color", "auto"));
            start.push_attribute(("w:fill", background_color.as_str()));
            writer.write_event(Event::Empty(start))?;
        }
        writer.write_event(Event::End(BytesEnd::new("w:tcPr")))?;
    }

    if cell.paragraphs().len() == 0 {
        write_paragraph(writer, &Paragraph::new())?;
    } else {
        for paragraph in cell.paragraphs() {
            write_paragraph(writer, paragraph)?;
        }
    }
    writer.write_event(Event::End(BytesEnd::new("w:tc")))?;
    Ok(())
}

fn write_borders<W>(writer: &mut Writer<W>, tag_name: &str, borders: &TableBorders) -> Result<()>
where
    W: Write,
{
    if !borders.has_serialized_content() {
        return Ok(());
    }

    writer.write_event(Event::Start(BytesStart::new(tag_name)))?;
    write_border(writer, "w:top", borders.top.as_ref())?;
    write_border(writer, "w:bottom", borders.bottom.as_ref())?;
    write_border(writer, "w:left", borders.left.as_ref())?;
    write_border(writer, "w:right", borders.right.as_ref())?;
    write_border(writer, "w:insideH", borders.inside_horizontal.as_ref())?;
    write_border(writer, "w:insideV", borders.inside_vertical.as_ref())?;
    writer.write_event(Event::End(BytesEnd::new(tag_name)))?;
    Ok(())
}

fn write_border<W>(writer: &mut Writer<W>, tag_name: &str, border: Option<&Border>) -> Result<()>
where
    W: Write,
{
    if let Some(border) = border {
        let mut start = BytesStart::new(tag_name);
        start.push_attribute(("w:val", border.style.as_xml_value()));
        let size_string = border.size.map(|value| value.to_string());
        if let Some(size_string) = size_string.as_deref() {
            start.push_attribute(("w:sz", size_string));
        }
        if let Some(color) = border.color.as_deref() {
            start.push_attribute(("w:color", color));
        }
        writer.write_event(Event::Empty(start))?;
    }
    Ok(())
}

fn write_section_properties<W>(
    writer: &mut Writer<W>,
    section_properties: &SectionProperties,
) -> Result<()>
where
    W: Write,
{
    writer.write_event(Event::Start(BytesStart::new("w:sectPr")))?;

    if section_properties.header.is_some() {
        let mut header = BytesStart::new("w:headerReference");
        header.push_attribute(("w:type", "default"));
        header.push_attribute(("r:id", DEFAULT_HEADER_REL_ID));
        writer.write_event(Event::Empty(header))?;
    } else if let Some(reference_id) = section_properties.header_reference_id.as_deref() {
        let mut header = BytesStart::new("w:headerReference");
        header.push_attribute(("w:type", "default"));
        header.push_attribute(("r:id", reference_id));
        writer.write_event(Event::Empty(header))?;
    }

    if section_properties.footer.is_some() {
        let mut footer = BytesStart::new("w:footerReference");
        footer.push_attribute(("w:type", "default"));
        footer.push_attribute(("r:id", DEFAULT_FOOTER_REL_ID));
        writer.write_event(Event::Empty(footer))?;
    } else if let Some(reference_id) = section_properties.footer_reference_id.as_deref() {
        let mut footer = BytesStart::new("w:footerReference");
        footer.push_attribute(("w:type", "default"));
        footer.push_attribute(("r:id", reference_id));
        writer.write_event(Event::Empty(footer))?;
    }

    let mut section_type = BytesStart::new("w:type");
    section_type.push_attribute(("w:val", section_properties.section_type.as_str()));
    writer.write_event(Event::Empty(section_type))?;

    let mut page_size = BytesStart::new("w:pgSz");
    let page_width = section_properties.page_setup.width_twips.to_string();
    let page_height = section_properties.page_setup.height_twips.to_string();
    page_size.push_attribute(("w:w", page_width.as_str()));
    page_size.push_attribute(("w:h", page_height.as_str()));
    writer.write_event(Event::Empty(page_size))?;

    let mut page_margins = BytesStart::new("w:pgMar");
    let top = section_properties.page_setup.margin_top_twips.to_string();
    let right = section_properties.page_setup.margin_right_twips.to_string();
    let bottom = section_properties
        .page_setup
        .margin_bottom_twips
        .to_string();
    let left = section_properties.page_setup.margin_left_twips.to_string();
    let header = section_properties.page_setup.header_twips.to_string();
    let footer = section_properties.page_setup.footer_twips.to_string();
    let gutter = section_properties.page_setup.gutter_twips.to_string();
    page_margins.push_attribute(("w:top", top.as_str()));
    page_margins.push_attribute(("w:right", right.as_str()));
    page_margins.push_attribute(("w:bottom", bottom.as_str()));
    page_margins.push_attribute(("w:left", left.as_str()));
    page_margins.push_attribute(("w:header", header.as_str()));
    page_margins.push_attribute(("w:footer", footer.as_str()));
    page_margins.push_attribute(("w:gutter", gutter.as_str()));
    writer.write_event(Event::Empty(page_margins))?;

    if let Some(page_numbering) = &section_properties.page_numbering {
        let mut start = BytesStart::new("w:pgNumType");
        start.push_attribute(("w:fmt", page_numbering.format.as_xml_value()));
        let start_at = page_numbering.start_at.map(|value| value.to_string());
        if let Some(start_at) = start_at.as_deref() {
            start.push_attribute(("w:start", start_at));
        }
        writer.write_event(Event::Empty(start))?;
    }

    writer.write_event(Event::End(BytesEnd::new("w:sectPr")))?;
    Ok(())
}

pub(crate) fn hydrate_section_from_package_parts(
    section_properties: &mut SectionProperties,
    package_parts: &BTreeMap<String, Vec<u8>>,
) -> Result<()> {
    let Some(relationships_xml) = package_parts.get("word/_rels/document.xml.rels") else {
        return Ok(());
    };
    let relationships = parse_relationship_targets(relationships_xml)?;

    if let Some(reference_id) = section_properties.header_reference_id.as_deref() {
        if let Some(target) = relationships.get(reference_id) {
            let part_name = word_part_name_from_target(target);
            if let Some(xml) = package_parts.get(&part_name) {
                if let Ok(header) = parse_header_footer_xml(xml, b"hdr") {
                    section_properties.header = Some(header);
                }
            }
        }
    }

    if let Some(reference_id) = section_properties.footer_reference_id.as_deref() {
        if let Some(target) = relationships.get(reference_id) {
            let part_name = word_part_name_from_target(target);
            if let Some(xml) = package_parts.get(&part_name) {
                if let Ok(footer) = parse_header_footer_xml(xml, b"ftr") {
                    section_properties.footer = Some(footer);
                }
            }
        }
    }

    Ok(())
}

pub(crate) fn render_header_footer_xml(
    header_footer: &HeaderFooter,
    root_tag: &str,
) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut root = BytesStart::new(root_tag);
    root.push_attribute(("xmlns:w", WORD_NS));
    writer.write_event(Event::Start(root))?;
    write_header_footer_paragraph(&mut writer, header_footer)?;
    writer.write_event(Event::End(BytesEnd::new(root_tag)))?;
    Ok(writer.into_inner())
}

fn parse_relationship_targets(xml: &[u8]) -> Result<BTreeMap<String, String>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut relationships = BTreeMap::new();
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if local_name(start.name().as_ref()) == b"Relationship" => {
                let Some(id) = attribute_value(&start, b"Id") else {
                    buffer.clear();
                    continue;
                };
                let Some(target) = attribute_value(&start, b"Target") else {
                    buffer.clear();
                    continue;
                };
                relationships.insert(id, target);
                skip_current_element(&mut reader)?;
            }
            Event::Empty(start) if local_name(start.name().as_ref()) == b"Relationship" => {
                let Some(id) = attribute_value(&start, b"Id") else {
                    buffer.clear();
                    continue;
                };
                let Some(target) = attribute_value(&start, b"Target") else {
                    buffer.clear();
                    continue;
                };
                relationships.insert(id, target);
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(relationships)
}

fn word_part_name_from_target(target: &str) -> String {
    if target.starts_with("word/") {
        target.to_string()
    } else if target.starts_with("/word/") {
        target.trim_start_matches('/').to_string()
    } else {
        format!("word/{target}")
    }
}

fn parse_header_footer_xml(xml: &[u8], root_local_name: &[u8]) -> Result<HeaderFooter> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut text = String::new();
    let mut alignment = ParagraphAlignment::Left;
    let mut saw_paragraph = false;
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if local_name(start.name().as_ref()) == root_local_name => {}
            Event::Start(start) if local_name(start.name().as_ref()) == b"p" => {
                let (paragraph_alignment, paragraph_text) =
                    parse_header_footer_paragraph(&mut reader)?;
                if !saw_paragraph {
                    alignment = paragraph_alignment;
                }
                if !paragraph_text.is_empty() {
                    if !text.is_empty() {
                        text.push('\n');
                    }
                    text.push_str(&paragraph_text);
                }
                saw_paragraph = true;
            }
            Event::Empty(start) if local_name(start.name().as_ref()) == b"p" => {
                saw_paragraph = true;
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(HeaderFooter::new(text).with_alignment(alignment))
}

#[derive(Default)]
struct HeaderFooterFieldState {
    skip_text_until_field_end: bool,
}

fn parse_header_footer_paragraph<R>(reader: &mut Reader<R>) -> Result<(ParagraphAlignment, String)>
where
    R: BufRead,
{
    let mut alignment = ParagraphAlignment::Left;
    let mut text = String::new();
    let mut field_state = HeaderFooterFieldState::default();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"pPr" => {
                    let properties = parse_paragraph_properties(reader, None)?;
                    alignment = properties.alignment.unwrap_or(ParagraphAlignment::Left);
                }
                b"r" => text.push_str(&parse_header_footer_run(reader, &mut field_state)?),
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => {
                if local_name(start.name().as_ref()) == b"r" {
                    // Ignore empty runs.
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"p" => break,
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML header/footer: unexpected end of file in w:p",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok((alignment, text))
}

fn parse_header_footer_run<R>(
    reader: &mut Reader<R>,
    field_state: &mut HeaderFooterFieldState,
) -> Result<String>
where
    R: BufRead,
{
    let mut text = String::new();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"rPr" => skip_current_element(reader)?,
                b"t" => {
                    if !field_state.skip_text_until_field_end {
                        text.push_str(&parse_text_element(reader)?);
                    } else {
                        skip_current_element(reader)?;
                    }
                }
                b"instrText" => {
                    if let Some(token) = parse_instruction_placeholder(reader)? {
                        field_state.skip_text_until_field_end = true;
                        text.push_str(token);
                    }
                }
                b"fldChar" => {
                    apply_field_char_state(field_state, &start);
                    skip_current_element(reader)?;
                }
                b"tab" => {
                    if !field_state.skip_text_until_field_end {
                        text.push('\t');
                    }
                    skip_current_element(reader)?;
                }
                b"br" => {
                    if !field_state.skip_text_until_field_end {
                        text.push('\n');
                    }
                    skip_current_element(reader)?;
                }
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => match local_name(start.name().as_ref()) {
                b"fldChar" => apply_field_char_state(field_state, &start),
                b"tab" if !field_state.skip_text_until_field_end => text.push('\t'),
                b"br" if !field_state.skip_text_until_field_end => text.push('\n'),
                _ => {}
            },
            Event::End(end) if local_name(end.name().as_ref()) == b"r" => break,
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML header/footer: unexpected end of file in w:r",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(text)
}

fn parse_instruction_placeholder<R>(reader: &mut Reader<R>) -> Result<Option<&'static str>>
where
    R: BufRead,
{
    let mut instruction = String::new();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Text(value) => {
                let raw = String::from_utf8_lossy(value.as_ref());
                instruction.push_str(&quick_xml::escape::unescape(&raw)?);
            }
            Event::CData(value) => {
                instruction.push_str(&String::from_utf8_lossy(value.as_ref()));
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"instrText" => break,
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML header/footer: unexpected end of file in w:instrText",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }

    let uppercase = instruction.to_ascii_uppercase();
    if uppercase.contains("NUMPAGES") {
        Ok(Some("{pages}"))
    } else if uppercase.contains("PAGE") {
        Ok(Some("{page}"))
    } else {
        Ok(None)
    }
}

fn apply_field_char_state(field_state: &mut HeaderFooterFieldState, start: &BytesStart<'_>) {
    match attribute_value(start, b"fldCharType").as_deref() {
        Some("begin") => field_state.skip_text_until_field_end = false,
        Some("separate") => field_state.skip_text_until_field_end = true,
        Some("end") => field_state.skip_text_until_field_end = false,
        _ => {}
    }
}

fn write_header_footer_paragraph<W>(
    writer: &mut Writer<W>,
    header_footer: &HeaderFooter,
) -> Result<()>
where
    W: Write,
{
    writer.write_event(Event::Start(BytesStart::new("w:p")))?;

    if header_footer.alignment != ParagraphAlignment::Left {
        writer.write_event(Event::Start(BytesStart::new("w:pPr")))?;
        let mut alignment = BytesStart::new("w:jc");
        alignment.push_attribute(("w:val", header_footer.alignment.as_xml_value()));
        writer.write_event(Event::Empty(alignment))?;
        writer.write_event(Event::End(BytesEnd::new("w:pPr")))?;
    }

    write_header_footer_template_runs(writer, &header_footer.text)?;
    writer.write_event(Event::End(BytesEnd::new("w:p")))?;
    Ok(())
}

fn write_header_footer_template_runs<W>(writer: &mut Writer<W>, template: &str) -> Result<()>
where
    W: Write,
{
    if template.is_empty() {
        write_plain_header_footer_run(writer, "")?;
        return Ok(());
    }

    let mut cursor = 0usize;
    while cursor < template.len() {
        let remaining = &template[cursor..];
        if remaining.starts_with("{pages}") {
            write_field_runs(writer, "NUMPAGES", "1")?;
            cursor += "{pages}".len();
            continue;
        }
        if remaining.starts_with("{page}") {
            write_field_runs(writer, "PAGE", "1")?;
            cursor += "{page}".len();
            continue;
        }

        let next_page = remaining.find("{page}");
        let next_pages = remaining.find("{pages}");
        let next_token = match (next_page, next_pages) {
            (Some(page), Some(pages)) => page.min(pages),
            (Some(page), None) => page,
            (None, Some(pages)) => pages,
            (None, None) => remaining.len(),
        };
        write_plain_header_footer_run(writer, &remaining[..next_token])?;
        cursor += next_token;
    }

    Ok(())
}

fn write_plain_header_footer_run<W>(writer: &mut Writer<W>, text: &str) -> Result<()>
where
    W: Write,
{
    writer.write_event(Event::Start(BytesStart::new("w:r")))?;
    write_run_content(writer, &Run::from_text(text))?;
    writer.write_event(Event::End(BytesEnd::new("w:r")))?;
    Ok(())
}

fn write_field_runs<W>(writer: &mut Writer<W>, instruction: &str, fallback_text: &str) -> Result<()>
where
    W: Write,
{
    writer.write_event(Event::Start(BytesStart::new("w:r")))?;
    let mut begin = BytesStart::new("w:fldChar");
    begin.push_attribute(("w:fldCharType", "begin"));
    writer.write_event(Event::Empty(begin))?;
    writer.write_event(Event::End(BytesEnd::new("w:r")))?;

    writer.write_event(Event::Start(BytesStart::new("w:r")))?;
    let mut instr_text = BytesStart::new("w:instrText");
    instr_text.push_attribute(("xml:space", "preserve"));
    writer.write_event(Event::Start(instr_text))?;
    writer.write_event(Event::Text(BytesText::new(&format!(" {instruction} "))))?;
    writer.write_event(Event::End(BytesEnd::new("w:instrText")))?;
    writer.write_event(Event::End(BytesEnd::new("w:r")))?;

    writer.write_event(Event::Start(BytesStart::new("w:r")))?;
    let mut separate = BytesStart::new("w:fldChar");
    separate.push_attribute(("w:fldCharType", "separate"));
    writer.write_event(Event::Empty(separate))?;
    writer.write_event(Event::End(BytesEnd::new("w:r")))?;

    write_plain_header_footer_run(writer, fallback_text)?;

    writer.write_event(Event::Start(BytesStart::new("w:r")))?;
    let mut end = BytesStart::new("w:fldChar");
    end.push_attribute(("w:fldCharType", "end"));
    writer.write_event(Event::Empty(end))?;
    writer.write_event(Event::End(BytesEnd::new("w:r")))?;
    Ok(())
}

fn attribute_value(start: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    for attribute in start.attributes().with_checks(false) {
        let Ok(attribute) = attribute else {
            continue;
        };
        if local_name(attribute.key.as_ref()) == key {
            return Some(String::from_utf8_lossy(attribute.value.as_ref()).into_owned());
        }
    }
    None
}

fn truthy_attribute(start: &BytesStart<'_>, key: &[u8]) -> Option<bool> {
    attribute_value(start, key).map(|value| matches!(value.as_str(), "true" | "1" | "on"))
}

fn skip_current_element<R>(reader: &mut Reader<R>) -> Result<()>
where
    R: BufRead,
{
    let mut depth = 1usize;
    let mut buffer = Vec::new();
    while depth > 0 {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(_) => depth += 1,
            Event::End(_) => depth -= 1,
            Event::Eof => {
                return Err(DocxError::parse(
                    "malformed OOXML document: unexpected end of file while skipping element",
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
    Ok(())
}

fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|byte| *byte == b':') {
        Some(index) => &name[index + 1..],
        None => name,
    }
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeMap;

    use super::{parse_document_xml, parse_numbering_xml, write_document_xml, write_numbering_xml};
    use crate::document::BodyBlock;
    use crate::{
        HeaderFooter, Paragraph, ParagraphAlignment, ParagraphList, Run, Table, TableCell, TableRow,
    };

    #[test]
    fn writer_emits_space_preserve_for_boundary_whitespace() {
        let paragraph = Paragraph::new().add_run(Run::from_text(" leading "));
        let document_xml = write_document_xml(
            &[BodyBlock::Paragraph(paragraph)],
            &super::SectionProperties::default(),
            &[],
        )
        .unwrap_or_else(|error| panic!("writer should succeed in test: {error}"));
        let xml = String::from_utf8_lossy(&document_xml);
        assert!(xml.contains(r#"xml:space="preserve""#));
    }

    #[test]
    fn writer_emits_tabs_and_breaks_as_ooxml_run_content() {
        let paragraph = Paragraph::new().add_run(Run::from_text("A\t B\r\nC\nD"));
        let document_xml = write_document_xml(
            &[BodyBlock::Paragraph(paragraph)],
            &super::SectionProperties::default(),
            &[],
        )
        .unwrap_or_else(|error| panic!("writer should succeed in test: {error}"));
        let xml = String::from_utf8_lossy(&document_xml);

        assert!(xml.contains("<w:t>A</w:t>"));
        assert!(xml.contains("<w:tab/>"));
        assert!(xml.contains(r#"<w:t xml:space="preserve"> B</w:t>"#));
        assert!(xml.matches("<w:br/>").count() >= 2);
        assert!(!xml.contains("A\t B"));
    }

    #[test]
    fn writer_emits_table_cell_paragraphs() {
        let table = Table::new().add_row(TableRow::new().add_cell(
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("x"))),
        ));
        let document_xml = write_document_xml(
            &[BodyBlock::Table(Box::new(table))],
            &super::SectionProperties::default(),
            &[],
        )
        .unwrap_or_else(|error| panic!("writer should succeed in test: {error}"));
        let xml = String::from_utf8_lossy(&document_xml);
        assert!(xml.contains("<w:tbl>"));
        assert!(xml.contains("<w:tc>"));
    }

    #[test]
    fn writer_emits_rich_presentation_properties() {
        let paragraph = Paragraph::new()
            .spacing_before(120)
            .spacing_after(240)
            .keep_next()
            .page_break_before()
            .add_run(
                Run::from_text("Styled")
                    .font("Arial")
                    .size_points(18)
                    .color("0F172A"),
            );
        let table = Table::new().add_row(
            TableRow::new().add_cell(
                TableCell::new()
                    .background("E2E8F0")
                    .add_paragraph(Paragraph::new().add_run(Run::from_text("Cell"))),
            ),
        );
        let document_xml = write_document_xml(
            &[
                BodyBlock::Paragraph(paragraph),
                BodyBlock::Table(Box::new(table)),
            ],
            &super::SectionProperties::default(),
            &[],
        )
        .unwrap_or_else(|error| panic!("writer should succeed in test: {error}"));
        let xml = String::from_utf8_lossy(&document_xml);
        assert!(xml.contains(r#"<w:spacing w:before="120" w:after="240""#));
        assert!(xml.contains("<w:keepNext/>"));
        assert!(xml.contains("<w:pageBreakBefore/>"));
        assert!(xml.contains(r#"<w:rFonts w:ascii="Arial""#));
        assert!(xml.contains(r#"<w:sz w:val="36""#));
        assert!(xml.contains(r#"<w:shd w:val="clear" w:color="auto" w:fill="E2E8F0""#));
    }

    #[test]
    fn header_footer_templates_emit_fields_and_round_trip_placeholders() {
        let footer =
            HeaderFooter::new("Page {page} of {pages}").with_alignment(ParagraphAlignment::Right);
        let footer_xml = super::render_header_footer_xml(&footer, "w:ftr")
            .unwrap_or_else(|error| panic!("render footer: {error}"));
        let xml = String::from_utf8_lossy(&footer_xml);

        assert!(xml.contains(r#"<w:jc w:val="right"/>"#));
        assert!(xml.contains(" PAGE "));
        assert!(xml.contains(" NUMPAGES "));

        let parsed = super::parse_header_footer_xml(&footer_xml, b"ftr")
            .unwrap_or_else(|error| panic!("parse footer: {error}"));
        assert_eq!(parsed, footer);
    }

    #[test]
    fn writer_emits_semantic_numbering_properties() {
        let paragraph = Paragraph::new()
            .with_list(ParagraphList::bullet_with_id(7).with_level(2))
            .add_run(Run::from_text("Item"));
        let document_xml = write_document_xml(
            &[BodyBlock::Paragraph(paragraph)],
            &super::SectionProperties::default(),
            &[],
        )
        .unwrap_or_else(|error| panic!("writer should succeed in test: {error}"));
        let xml = String::from_utf8_lossy(&document_xml);

        assert!(xml.contains("<w:numPr>"));
        assert!(xml.contains(r#"<w:ilvl w:val="2"/>"#));
        assert!(xml.contains(r#"<w:numId w:val="7"/>"#));
        assert!(!xml.contains("• Item"));
    }

    #[test]
    fn numbering_part_defines_bullet_and_decimal_instances() {
        let body = [
            BodyBlock::Paragraph(
                Paragraph::new()
                    .with_list(ParagraphList::bullet_with_id(3))
                    .add_run(Run::from_text("Bullet")),
            ),
            BodyBlock::Paragraph(
                Paragraph::new()
                    .with_list(ParagraphList::numbered_with_id(9).with_level(1))
                    .add_run(Run::from_text("Number")),
            ),
        ];
        let numbering_xml =
            write_numbering_xml(&body).unwrap_or_else(|error| panic!("write numbering: {error}"));
        let xml = String::from_utf8_lossy(&numbering_xml);

        assert!(xml.contains(r#"<w:abstractNum w:abstractNumId="1">"#));
        assert!(xml.contains(r#"<w:abstractNum w:abstractNumId="2">"#));
        assert!(xml.contains(r#"<w:num w:numId="3">"#));
        assert!(xml.contains(r#"<w:num w:numId="9">"#));
        assert!(xml.contains(r#"<w:numFmt w:val="bullet"/>"#));
        assert!(xml.contains(r#"<w:numFmt w:val="decimal"/>"#));
    }

    #[test]
    fn parser_restores_semantic_numbering_from_document_and_numbering_parts() {
        let body = [BodyBlock::Paragraph(
            Paragraph::new()
                .with_list(ParagraphList::numbered_with_id(4).with_level(1))
                .add_run(Run::from_text("Alpha")),
        )];
        let numbering_xml =
            write_numbering_xml(&body).unwrap_or_else(|error| panic!("write numbering: {error}"));
        let numbering = parse_numbering_xml(&numbering_xml)
            .unwrap_or_else(|error| panic!("parse numbering: {error}"));
        let document_xml = write_document_xml(&body, &super::SectionProperties::default(), &[])
            .unwrap_or_else(|error| panic!("write document: {error}"));
        let parsed = parse_document_xml(&document_xml, Some(&numbering), &BTreeMap::new())
            .unwrap_or_else(|error| panic!("parse document: {error}"));

        let paragraph = match parsed.body.first() {
            Some(BodyBlock::Paragraph(paragraph)) => paragraph,
            other => panic!("expected paragraph body block, got {other:?}"),
        };

        assert_eq!(
            paragraph.list(),
            Some(&ParagraphList::numbered_with_id(4).with_level(1))
        );
        assert_eq!(paragraph.text(), "Alpha");
    }

    #[test]
    fn parser_restores_keep_next_from_document_xml() {
        let body = [BodyBlock::Paragraph(
            Paragraph::new()
                .keep_next()
                .add_run(Run::from_text("Keep me with next")),
        )];
        let document_xml = write_document_xml(&body, &super::SectionProperties::default(), &[])
            .unwrap_or_else(|error| panic!("write document: {error}"));
        let parsed = parse_document_xml(&document_xml, None, &BTreeMap::new())
            .unwrap_or_else(|error| panic!("parse document: {error}"));

        let paragraph = match parsed.body.first() {
            Some(BodyBlock::Paragraph(paragraph)) => paragraph,
            other => panic!("expected paragraph body block, got {other:?}"),
        };

        assert!(paragraph.has_keep_next());
    }
}
