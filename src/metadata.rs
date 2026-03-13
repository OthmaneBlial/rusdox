use std::collections::BTreeMap;
use std::io::Cursor;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use crate::Result;

pub(crate) const DEFAULT_TITLE: &str = "RusDox Document";
pub(crate) const DEFAULT_AUTHOR: &str = "RusDox";
pub(crate) const PROPERTIES_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
pub(crate) const VT_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
pub(crate) const CUSTOM_PROPERTY_FMTID: &str = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

/// First-class document metadata used for DOCX package properties.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct DocumentMetadata {
    pub title: Option<String>,
    pub author: Option<String>,
    pub subject: Option<String>,
    pub keywords: Vec<String>,
    pub custom_properties: BTreeMap<String, String>,
}

impl DocumentMetadata {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(title.into());
        self
    }

    pub fn author(mut self, author: impl Into<String>) -> Self {
        self.author = Some(author.into());
        self
    }

    pub fn subject(mut self, subject: impl Into<String>) -> Self {
        self.subject = Some(subject.into());
        self
    }

    pub fn keyword(mut self, keyword: impl Into<String>) -> Self {
        self.keywords.push(keyword.into());
        self
    }

    pub fn custom_property(mut self, name: impl Into<String>, value: impl Into<String>) -> Self {
        self.custom_properties.insert(name.into(), value.into());
        self
    }

    pub(crate) fn resolved_title(&self) -> &str {
        normalized_text(self.title.as_deref()).unwrap_or(DEFAULT_TITLE)
    }

    pub(crate) fn resolved_author(&self) -> &str {
        normalized_text(self.author.as_deref()).unwrap_or(DEFAULT_AUTHOR)
    }
}

pub(crate) fn render_core_properties_xml(metadata: &DocumentMetadata) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut root = BytesStart::new("cp:coreProperties");
    root.push_attribute((
        "xmlns:cp",
        "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    ));
    root.push_attribute(("xmlns:dc", "http://purl.org/dc/elements/1.1/"));
    root.push_attribute(("xmlns:dcterms", "http://purl.org/dc/terms/"));
    root.push_attribute(("xmlns:dcmitype", "http://purl.org/dc/dcmitype/"));
    root.push_attribute(("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"));
    writer.write_event(Event::Start(root))?;

    write_text_element(&mut writer, "dc:title", metadata.resolved_title())?;
    write_text_element(&mut writer, "dc:creator", metadata.resolved_author())?;
    if let Some(subject) = normalized_text(metadata.subject.as_deref()) {
        write_text_element(&mut writer, "dc:subject", subject)?;
    }
    if !metadata.keywords.is_empty() {
        write_text_element(&mut writer, "cp:keywords", &metadata.keywords.join(", "))?;
    }
    write_text_element(&mut writer, "cp:lastModifiedBy", metadata.resolved_author())?;
    write_w3cdtf_element(&mut writer, "dcterms:created", "2026-03-10T00:00:00Z")?;
    write_w3cdtf_element(&mut writer, "dcterms:modified", "2026-03-10T00:00:00Z")?;

    writer.write_event(Event::End(BytesEnd::new("cp:coreProperties")))?;
    Ok(writer.into_inner())
}

pub(crate) fn render_custom_properties_xml(metadata: &DocumentMetadata) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut root = BytesStart::new("Properties");
    root.push_attribute(("xmlns", PROPERTIES_NS));
    root.push_attribute(("xmlns:vt", VT_NS));
    writer.write_event(Event::Start(root))?;

    for (index, (name, value)) in metadata.custom_properties.iter().enumerate() {
        let mut property = BytesStart::new("property");
        let pid = (index + 2).to_string();
        property.push_attribute(("fmtid", CUSTOM_PROPERTY_FMTID));
        property.push_attribute(("pid", pid.as_str()));
        property.push_attribute(("name", name.as_str()));
        writer.write_event(Event::Start(property))?;
        write_text_element(&mut writer, "vt:lpwstr", value)?;
        writer.write_event(Event::End(BytesEnd::new("property")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("Properties")))?;
    Ok(writer.into_inner())
}

pub(crate) fn parse_core_properties_xml(xml: &[u8]) -> Result<DocumentMetadata> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut metadata = DocumentMetadata::default();
    let mut current = None;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => {
                current = match local_name(start.name().as_ref()) {
                    b"title" => Some("title"),
                    b"creator" => Some("author"),
                    b"subject" => Some("subject"),
                    b"keywords" => Some("keywords"),
                    _ => None,
                };
            }
            Event::Text(text) => {
                let Some(field) = current else {
                    buffer.clear();
                    continue;
                };
                let raw = String::from_utf8_lossy(text.as_ref());
                let value = quick_xml::escape::unescape(&raw)?.into_owned();
                match field {
                    "title" => metadata.title = Some(value),
                    "author" => metadata.author = Some(value),
                    "subject" => metadata.subject = Some(value),
                    "keywords" => metadata.keywords = parse_keywords(&value),
                    _ => {}
                }
            }
            Event::End(_) => current = None,
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(metadata)
}

pub(crate) fn parse_custom_properties_xml(
    xml: &[u8],
    metadata: &mut DocumentMetadata,
) -> Result<()> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut current_name = None;
    let mut current_value = String::new();
    let mut capture_text = false;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) => match local_name(start.name().as_ref()) {
                b"property" => {
                    current_name = attribute_value(&start, b"name");
                    current_value.clear();
                    capture_text = false;
                }
                _ if current_name.is_some() => capture_text = true,
                _ => {}
            },
            Event::Text(text) if capture_text && current_name.is_some() => {
                let raw = String::from_utf8_lossy(text.as_ref());
                current_value.push_str(&quick_xml::escape::unescape(&raw)?);
            }
            Event::End(end) => match local_name(end.name().as_ref()) {
                b"property" => {
                    if let Some(name) = current_name.take() {
                        metadata
                            .custom_properties
                            .insert(name, current_value.clone());
                    }
                    current_value.clear();
                    capture_text = false;
                }
                _ if current_name.is_some() => capture_text = false,
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(())
}

fn write_text_element<W>(writer: &mut Writer<W>, name: &str, text: &str) -> Result<()>
where
    W: std::io::Write,
{
    writer.write_event(Event::Start(BytesStart::new(name)))?;
    writer.write_event(Event::Text(BytesText::new(text)))?;
    writer.write_event(Event::End(BytesEnd::new(name)))?;
    Ok(())
}

fn write_w3cdtf_element<W>(writer: &mut Writer<W>, name: &str, text: &str) -> Result<()>
where
    W: std::io::Write,
{
    let mut element = BytesStart::new(name);
    element.push_attribute(("xsi:type", "dcterms:W3CDTF"));
    writer.write_event(Event::Start(element))?;
    writer.write_event(Event::Text(BytesText::new(text)))?;
    writer.write_event(Event::End(BytesEnd::new(name)))?;
    Ok(())
}

fn parse_keywords(raw: &str) -> Vec<String> {
    raw.split([',', ';'])
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .map(str::to_string)
        .collect()
}

fn normalized_text(value: Option<&str>) -> Option<&str> {
    value.and_then(|value| {
        let trimmed = value.trim();
        if trimmed.is_empty() {
            None
        } else {
            Some(trimmed)
        }
    })
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn attribute_value(start: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    start.attributes().flatten().find_map(|attribute| {
        if local_name(attribute.key.as_ref()) == key {
            Some(String::from_utf8_lossy(attribute.value.as_ref()).into_owned())
        } else {
            None
        }
    })
}

#[cfg(test)]
mod tests {
    use super::{
        parse_core_properties_xml, parse_custom_properties_xml, render_core_properties_xml,
        render_custom_properties_xml, DocumentMetadata,
    };

    #[test]
    fn builder_methods_compose_metadata() {
        let metadata = DocumentMetadata::new()
            .title("Board Report")
            .author("RusDox")
            .subject("Quarterly review")
            .keyword("finance")
            .custom_property("Client", "Acme");

        assert_eq!(metadata.title.as_deref(), Some("Board Report"));
        assert_eq!(metadata.author.as_deref(), Some("RusDox"));
        assert_eq!(metadata.subject.as_deref(), Some("Quarterly review"));
        assert_eq!(metadata.keywords, vec!["finance"]);
        assert_eq!(
            metadata.custom_properties.get("Client").map(String::as_str),
            Some("Acme")
        );
    }

    #[test]
    fn core_properties_round_trip() {
        let metadata = DocumentMetadata::new()
            .title("Board Report")
            .author("Ops")
            .subject("Q2")
            .keyword("finance")
            .keyword("board");

        let xml = render_core_properties_xml(&metadata).expect("render core properties");
        let parsed = parse_core_properties_xml(&xml).expect("parse core properties");

        assert_eq!(parsed.title.as_deref(), Some("Board Report"));
        assert_eq!(parsed.author.as_deref(), Some("Ops"));
        assert_eq!(parsed.subject.as_deref(), Some("Q2"));
        assert_eq!(parsed.keywords, vec!["finance", "board"]);
    }

    #[test]
    fn custom_properties_round_trip() {
        let metadata = DocumentMetadata::new()
            .custom_property("Client", "Acme")
            .custom_property("Region", "EMEA");

        let xml = render_custom_properties_xml(&metadata).expect("render custom properties");
        let mut parsed = DocumentMetadata::default();
        parse_custom_properties_xml(&xml, &mut parsed).expect("parse custom properties");

        assert_eq!(
            parsed.custom_properties.get("Client").map(String::as_str),
            Some("Acme")
        );
        assert_eq!(
            parsed.custom_properties.get("Region").map(String::as_str),
            Some("EMEA")
        );
    }
}
