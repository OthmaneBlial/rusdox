use std::io::{BufRead, Cursor, Write};

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use crate::document::BodyBlock;
use crate::error::{DocxError, Result};
use crate::paragraph::{Paragraph, ParagraphAlignment};
use crate::run::{Run, RunProperties, UnderlineStyle, VerticalAlign};
use crate::table::{
    Border, BorderStyle, Table, TableBorders, TableCell, TableCellProperties, TableProperties,
    TableRow,
};

const WORD_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Internal section settings preserved for the generated document body.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct SectionProperties {
    section_type: String,
    page_width: u32,
    page_height: u32,
    margin_top: u32,
    margin_right: u32,
    margin_bottom: u32,
    margin_left: u32,
    header: u32,
    footer: u32,
    gutter: u32,
}

impl Default for SectionProperties {
    fn default() -> Self {
        Self {
            section_type: "nextPage".to_string(),
            page_width: 12_240,
            page_height: 15_840,
            margin_top: 1_440,
            margin_right: 1_440,
            margin_bottom: 1_440,
            margin_left: 1_440,
            header: 720,
            footer: 720,
            gutter: 0,
        }
    }
}

pub(crate) struct ParsedDocument {
    pub(crate) body: Vec<BodyBlock>,
    pub(crate) section_properties: SectionProperties,
}

pub(crate) fn parse_document_xml(xml: &[u8]) -> Result<ParsedDocument> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if local_name(start.name().as_ref()) == b"body" => {
                let parsed = parse_body(&mut reader)?;
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
    writer.write_event(Event::Start(document))?;

    writer.write_event(Event::Start(BytesStart::new("w:body")))?;
    for block in body {
        match block {
            BodyBlock::Paragraph(paragraph) => write_paragraph(&mut writer, paragraph)?,
            BodyBlock::Table(table) => write_table(&mut writer, table)?,
        }
    }
    write_section_properties(&mut writer, section_properties)?;
    writer.write_event(Event::End(BytesEnd::new("w:body")))?;
    writer.write_event(Event::End(BytesEnd::new("w:document")))?;
    Ok(writer.into_inner())
}

fn parse_body<R>(reader: &mut Reader<R>) -> Result<ParsedDocument>
where
    R: BufRead,
{
    let mut body = Vec::new();
    let mut section_properties = SectionProperties::default();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"p" => body.push(BodyBlock::Paragraph(parse_paragraph(reader)?)),
                b"tbl" => body.push(BodyBlock::Table(Box::new(parse_table(reader)?))),
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

fn parse_paragraph<R>(reader: &mut Reader<R>) -> Result<Paragraph>
where
    R: BufRead,
{
    let mut runs = Vec::new();
    let mut alignment = None;
    let mut spacing_before = None;
    let mut spacing_after = None;
    let mut page_break_before = false;
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"r" => runs.push(parse_run(reader)?),
                b"pPr" => {
                    let properties = parse_paragraph_properties(reader)?;
                    alignment = properties.alignment;
                    spacing_before = properties.spacing_before;
                    spacing_after = properties.spacing_after;
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

    Ok(Paragraph::from_parts(
        runs,
        alignment,
        spacing_before,
        spacing_after,
        page_break_before,
    ))
}

#[derive(Default)]
struct ParsedParagraphProperties {
    alignment: Option<ParagraphAlignment>,
    spacing_before: Option<u32>,
    spacing_after: Option<u32>,
    page_break_before: bool,
}

fn parse_paragraph_properties<R>(reader: &mut Reader<R>) -> Result<ParsedParagraphProperties>
where
    R: BufRead,
{
    let mut properties = ParsedParagraphProperties::default();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
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
                b"pageBreakBefore" => {
                    properties.page_break_before = truthy_attribute(&start, b"val").unwrap_or(true);
                    skip_current_element(reader)?;
                }
                _ => skip_current_element(reader)?,
            },
            Event::Empty(start) => match local_name(start.name().as_ref()) {
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

fn parse_run<R>(reader: &mut Reader<R>) -> Result<Run>
where
    R: BufRead,
{
    let mut text = String::new();
    let mut properties = RunProperties::default();
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

    Ok(Run::from_parts(text, properties))
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

fn parse_table<R>(reader: &mut Reader<R>) -> Result<Table>
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
                b"tr" => rows.push(parse_table_row(reader)?),
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

fn parse_table_row<R>(reader: &mut Reader<R>) -> Result<TableRow>
where
    R: BufRead,
{
    let mut cells = Vec::new();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"tc" => cells.push(parse_table_cell(reader)?),
                _ => skip_current_element(reader)?,
            },
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

    Ok(TableRow::from_parts(cells))
}

fn parse_table_cell<R>(reader: &mut Reader<R>) -> Result<TableCell>
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
                b"p" => paragraphs.push(parse_paragraph(reader)?),
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
                b"pgSz" => {
                    if let Some(value) =
                        attribute_value(&start, b"w").and_then(|value| value.parse().ok())
                    {
                        section.page_width = value;
                    }
                    if let Some(value) =
                        attribute_value(&start, b"h").and_then(|value| value.parse().ok())
                    {
                        section.page_height = value;
                    }
                    skip_current_element(reader)?;
                }
                b"pgMar" => {
                    parse_page_margins(&mut section, &start);
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
                b"pgSz" => {
                    if let Some(value) =
                        attribute_value(&start, b"w").and_then(|value| value.parse().ok())
                    {
                        section.page_width = value;
                    }
                    if let Some(value) =
                        attribute_value(&start, b"h").and_then(|value| value.parse().ok())
                    {
                        section.page_height = value;
                    }
                }
                b"pgMar" => parse_page_margins(&mut section, &start),
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
        section.margin_top = value;
    }
    if let Some(value) = attribute_value(start, b"right").and_then(|value| value.parse().ok()) {
        section.margin_right = value;
    }
    if let Some(value) = attribute_value(start, b"bottom").and_then(|value| value.parse().ok()) {
        section.margin_bottom = value;
    }
    if let Some(value) = attribute_value(start, b"left").and_then(|value| value.parse().ok()) {
        section.margin_left = value;
    }
    if let Some(value) = attribute_value(start, b"header").and_then(|value| value.parse().ok()) {
        section.header = value;
    }
    if let Some(value) = attribute_value(start, b"footer").and_then(|value| value.parse().ok()) {
        section.footer = value;
    }
    if let Some(value) = attribute_value(start, b"gutter").and_then(|value| value.parse().ok()) {
        section.gutter = value;
    }
}

fn write_paragraph<W>(writer: &mut Writer<W>, paragraph: &Paragraph) -> Result<()>
where
    W: Write,
{
    writer.write_event(Event::Start(BytesStart::new("w:p")))?;
    if paragraph.has_properties() {
        writer.write_event(Event::Start(BytesStart::new("w:pPr")))?;
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

    let mut text = BytesStart::new("w:t");
    if run.needs_space_preserve() {
        text.push_attribute(("xml:space", "preserve"));
    }
    writer.write_event(Event::Start(text))?;
    writer.write_event(Event::Text(BytesText::new(run.text())))?;
    writer.write_event(Event::End(BytesEnd::new("w:t")))?;
    writer.write_event(Event::End(BytesEnd::new("w:r")))?;
    Ok(())
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

    let mut section_type = BytesStart::new("w:type");
    section_type.push_attribute(("w:val", section_properties.section_type.as_str()));
    writer.write_event(Event::Empty(section_type))?;

    let mut page_size = BytesStart::new("w:pgSz");
    let page_width = section_properties.page_width.to_string();
    let page_height = section_properties.page_height.to_string();
    page_size.push_attribute(("w:w", page_width.as_str()));
    page_size.push_attribute(("w:h", page_height.as_str()));
    writer.write_event(Event::Empty(page_size))?;

    let mut page_margins = BytesStart::new("w:pgMar");
    let top = section_properties.margin_top.to_string();
    let right = section_properties.margin_right.to_string();
    let bottom = section_properties.margin_bottom.to_string();
    let left = section_properties.margin_left.to_string();
    let header = section_properties.header.to_string();
    let footer = section_properties.footer.to_string();
    let gutter = section_properties.gutter.to_string();
    page_margins.push_attribute(("w:top", top.as_str()));
    page_margins.push_attribute(("w:right", right.as_str()));
    page_margins.push_attribute(("w:bottom", bottom.as_str()));
    page_margins.push_attribute(("w:left", left.as_str()));
    page_margins.push_attribute(("w:header", header.as_str()));
    page_margins.push_attribute(("w:footer", footer.as_str()));
    page_margins.push_attribute(("w:gutter", gutter.as_str()));
    writer.write_event(Event::Empty(page_margins))?;

    writer.write_event(Event::End(BytesEnd::new("w:sectPr")))?;
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
    use super::write_document_xml;
    use crate::document::BodyBlock;
    use crate::{Paragraph, Run, Table, TableCell, TableRow};

    #[test]
    fn writer_emits_space_preserve_for_boundary_whitespace() {
        let paragraph = Paragraph::new().add_run(Run::from_text(" leading "));
        let document_xml = write_document_xml(
            &[BodyBlock::Paragraph(paragraph)],
            &super::SectionProperties::default(),
        )
        .unwrap_or_else(|error| panic!("writer should succeed in test: {error}"));
        let xml = String::from_utf8_lossy(&document_xml);
        assert!(xml.contains(r#"xml:space="preserve""#));
    }

    #[test]
    fn writer_emits_table_cell_paragraphs() {
        let table = Table::new().add_row(TableRow::new().add_cell(
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("x"))),
        ));
        let document_xml = write_document_xml(
            &[BodyBlock::Table(Box::new(table))],
            &super::SectionProperties::default(),
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
        )
        .unwrap_or_else(|error| panic!("writer should succeed in test: {error}"));
        let xml = String::from_utf8_lossy(&document_xml);
        assert!(xml.contains(r#"<w:spacing w:before="120" w:after="240""#));
        assert!(xml.contains("<w:pageBreakBefore/>"));
        assert!(xml.contains(r#"<w:rFonts w:ascii="Arial""#));
        assert!(xml.contains(r#"<w:sz w:val="36""#));
        assert!(xml.contains(r#"<w:shd w:val="clear" w:color="auto" w:fill="E2E8F0""#));
    }
}
