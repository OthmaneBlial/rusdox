use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::io::{Read, Seek, SeekFrom};
use std::path::{Component, Path, PathBuf};

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use serde::Serialize;
use zip::ZipArchive;

use crate::{DocxError, InputLimits, Result};

const MAIN_DOCUMENT_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

/// Structural OOXML package validation evidence.
#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct PackageValidationReport {
    /// Whether every package, XML, content-type, and relationship check passed.
    pub valid: bool,
    /// Number of non-directory package parts inspected.
    pub parts_checked: usize,
    /// Number of relationships inspected across all `.rels` parts.
    pub relationships_checked: usize,
    /// Actionable validation failures. Empty when `valid` is true.
    pub errors: Vec<String>,
}

/// Validates a DOCX package with the default resource ceilings.
pub fn validate_docx_package(path: impl AsRef<Path>) -> Result<PackageValidationReport> {
    validate_docx_package_with_limits(path, InputLimits::default())
}

/// Validates a DOCX package with explicit resource ceilings.
pub fn validate_docx_package_with_limits(
    path: impl AsRef<Path>,
    limits: InputLimits,
) -> Result<PackageValidationReport> {
    let file = fs::File::open(path)?;
    validate_docx_reader_with_limits(file, limits)
}

/// Validates a seekable DOCX reader with explicit resource ceilings.
pub fn validate_docx_reader_with_limits<R>(
    mut reader: R,
    limits: InputLimits,
) -> Result<PackageValidationReport>
where
    R: Read + Seek,
{
    let start = reader.stream_position()?;
    let archive_bytes = reader.seek(SeekFrom::End(0))?.saturating_sub(start);
    if archive_bytes > limits.max_docx_archive_bytes {
        return Err(DocxError::resource_limit(format!(
            "DOCX archive is {archive_bytes} bytes; limit is {} bytes",
            limits.max_docx_archive_bytes
        )));
    }
    reader.seek(SeekFrom::Start(start))?;
    let mut archive = ZipArchive::new(reader)?;
    if archive.len() > limits.max_docx_entries {
        return Err(DocxError::resource_limit(format!(
            "DOCX contains {} ZIP entries; limit is {}",
            archive.len(),
            limits.max_docx_entries
        )));
    }

    let mut parts = BTreeMap::new();
    let mut total_uncompressed = 0_u64;
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index)?;
        if entry.is_dir() {
            continue;
        }
        let name = entry.name().to_string();
        if entry.enclosed_name().is_none() || name.contains('\\') {
            return Err(DocxError::parse(format!(
                "unsafe DOCX ZIP entry path: {name}"
            )));
        }
        if parts.contains_key(&name) {
            return Err(DocxError::parse(format!(
                "duplicate DOCX ZIP entry: {name}"
            )));
        }
        let uncompressed = entry.size();
        let compressed = entry.compressed_size();
        enforce_entry_limits(&name, uncompressed, compressed, limits)?;
        total_uncompressed = total_uncompressed.saturating_add(uncompressed);
        if total_uncompressed > limits.max_docx_total_bytes {
            return Err(DocxError::resource_limit(format!(
                "DOCX expands beyond the {} byte total limit",
                limits.max_docx_total_bytes
            )));
        }
        let mut bytes = Vec::new();
        (&mut entry)
            .take(uncompressed.saturating_add(1))
            .read_to_end(&mut bytes)?;
        parts.insert(name, bytes);
    }

    Ok(validate_parts(&parts))
}

fn enforce_entry_limits(
    name: &str,
    uncompressed: u64,
    compressed: u64,
    limits: InputLimits,
) -> Result<()> {
    if uncompressed > limits.max_docx_entry_bytes {
        return Err(DocxError::resource_limit(format!(
            "DOCX part '{name}' expands to {uncompressed} bytes; per-entry limit is {} bytes",
            limits.max_docx_entry_bytes
        )));
    }
    if is_xml_part(name) && uncompressed > limits.max_xml_bytes {
        return Err(DocxError::resource_limit(format!(
            "XML part '{name}' is {uncompressed} bytes; limit is {} bytes",
            limits.max_xml_bytes
        )));
    }
    if uncompressed > 0
        && (compressed == 0
            || uncompressed > compressed.saturating_mul(limits.max_zip_compression_ratio))
    {
        return Err(DocxError::resource_limit(format!(
            "DOCX part '{name}' exceeds the {}:1 ZIP compression-ratio limit",
            limits.max_zip_compression_ratio
        )));
    }
    Ok(())
}

fn validate_parts(parts: &BTreeMap<String, Vec<u8>>) -> PackageValidationReport {
    let mut errors = Vec::new();
    for required in [
        "[Content_Types].xml",
        "_rels/.rels",
        "word/document.xml",
        "word/_rels/document.xml.rels",
    ] {
        if !parts.contains_key(required) {
            errors.push(format!("missing required OOXML part: {required}"));
        }
    }

    for (name, bytes) in parts.iter().filter(|(name, _)| is_xml_part(name)) {
        if let Err(error) = validate_xml(bytes) {
            errors.push(format!("invalid XML in '{name}': {error}"));
        }
    }

    let (defaults, overrides) = parts
        .get("[Content_Types].xml")
        .and_then(|bytes| match parse_content_types(bytes) {
            Ok(value) => Some(value),
            Err(error) => {
                errors.push(format!("invalid [Content_Types].xml: {error}"));
                None
            }
        })
        .unwrap_or_default();
    if overrides.get("word/document.xml").map(String::as_str) != Some(MAIN_DOCUMENT_CONTENT_TYPE) {
        errors.push(format!(
            "word/document.xml must declare content type '{MAIN_DOCUMENT_CONTENT_TYPE}'"
        ));
    }
    for name in parts.keys().filter(|name| *name != "[Content_Types].xml") {
        let extension = if name.ends_with(".rels") {
            "rels".to_string()
        } else {
            Path::new(name)
                .extension()
                .and_then(|value| value.to_str())
                .unwrap_or_default()
                .to_ascii_lowercase()
        };
        if !overrides.contains_key(name) && !defaults.contains_key(&extension) {
            errors.push(format!("part '{name}' has no content-type declaration"));
        }
    }

    let mut relationships_checked = 0;
    let mut root_office_document = false;
    for (rels_name, bytes) in parts.iter().filter(|(name, _)| name.ends_with(".rels")) {
        match parse_relationships(bytes) {
            Ok(relationships) => {
                let mut ids = BTreeSet::new();
                for relationship in relationships {
                    relationships_checked += 1;
                    if !ids.insert(relationship.id.clone()) {
                        errors.push(format!(
                            "relationship part '{rels_name}' repeats Id '{}'",
                            relationship.id
                        ));
                    }
                    if rels_name == "_rels/.rels" && relationship.kind.ends_with("/officeDocument")
                    {
                        root_office_document = relationship.target == "word/document.xml";
                    }
                    if relationship.external {
                        continue;
                    }
                    match resolve_relationship_target(rels_name, &relationship.target) {
                        Some(target) if parts.contains_key(&target) => {}
                        Some(target) => errors.push(format!(
                            "relationship '{}' in '{rels_name}' targets missing part '{target}'",
                            relationship.id
                        )),
                        None => errors.push(format!(
                            "relationship '{}' in '{rels_name}' has an unsafe target '{}'",
                            relationship.id, relationship.target
                        )),
                    }
                }
            }
            Err(error) => errors.push(format!("invalid relationships part '{rels_name}': {error}")),
        }
    }
    if parts.contains_key("_rels/.rels") && !root_office_document {
        errors.push(
            "root relationships must target word/document.xml as the officeDocument".to_string(),
        );
    }

    PackageValidationReport {
        valid: errors.is_empty(),
        parts_checked: parts.len(),
        relationships_checked,
        errors,
    }
}

fn validate_xml(bytes: &[u8]) -> Result<()> {
    let mut reader = Reader::from_reader(bytes);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Eof => return Ok(()),
            _ => buffer.clear(),
        }
    }
}

fn parse_content_types(
    bytes: &[u8],
) -> Result<(BTreeMap<String, String>, BTreeMap<String, String>)> {
    let mut defaults = BTreeMap::new();
    let mut overrides = BTreeMap::new();
    let mut reader = Reader::from_reader(bytes);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Empty(start) | Event::Start(start) => match start.local_name().as_ref() {
                b"Default" => {
                    if let (Some(extension), Some(content_type)) = (
                        attribute(&start, b"Extension"),
                        attribute(&start, b"ContentType"),
                    ) {
                        defaults.insert(extension.to_ascii_lowercase(), content_type);
                    }
                }
                b"Override" => {
                    if let (Some(part_name), Some(content_type)) = (
                        attribute(&start, b"PartName"),
                        attribute(&start, b"ContentType"),
                    ) {
                        overrides
                            .insert(part_name.trim_start_matches('/').to_string(), content_type);
                    }
                }
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok((defaults, overrides))
}

#[derive(Debug)]
struct Relationship {
    id: String,
    kind: String,
    target: String,
    external: bool,
}

fn parse_relationships(bytes: &[u8]) -> Result<Vec<Relationship>> {
    let mut relationships = Vec::new();
    let mut reader = Reader::from_reader(bytes);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Empty(start) | Event::Start(start)
                if start.local_name().as_ref() == b"Relationship" =>
            {
                let id = attribute(&start, b"Id")
                    .ok_or_else(|| DocxError::parse("relationship is missing Id"))?;
                let kind = attribute(&start, b"Type")
                    .ok_or_else(|| DocxError::parse("relationship is missing Type"))?;
                let target = attribute(&start, b"Target")
                    .ok_or_else(|| DocxError::parse("relationship is missing Target"))?;
                let external = attribute(&start, b"TargetMode")
                    .is_some_and(|mode| mode.eq_ignore_ascii_case("external"));
                relationships.push(Relationship {
                    id,
                    kind,
                    target,
                    external,
                });
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(relationships)
}

fn attribute(start: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    start
        .attributes()
        .with_checks(false)
        .filter_map(std::result::Result::ok)
        .find(|attribute| attribute.key.local_name().as_ref() == name)
        .and_then(|attribute| {
            attribute
                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .ok()
                .map(|value| value.into_owned())
        })
}

fn resolve_relationship_target(rels_name: &str, target: &str) -> Option<String> {
    let target = target.split(['#', '?']).next().unwrap_or_default();
    let base = if rels_name == "_rels/.rels" {
        PathBuf::new()
    } else {
        let (prefix, file) = rels_name.rsplit_once("/_rels/")?;
        let source_file = file.strip_suffix(".rels")?;
        Path::new(prefix)
            .join(source_file)
            .parent()
            .unwrap_or_else(|| Path::new(""))
            .to_path_buf()
    };
    let candidate = if target.starts_with('/') {
        PathBuf::from(target.trim_start_matches('/'))
    } else {
        base.join(target)
    };
    normalize_package_path(&candidate)
}

fn normalize_package_path(path: &Path) -> Option<String> {
    let mut components = Vec::new();
    for component in path.components() {
        match component {
            Component::Normal(value) => components.push(value.to_string_lossy().into_owned()),
            Component::CurDir => {}
            Component::ParentDir => {
                components.pop()?;
            }
            Component::RootDir | Component::Prefix(_) => return None,
        }
    }
    (!components.is_empty()).then(|| components.join("/"))
}

fn is_xml_part(name: &str) -> bool {
    name.ends_with(".xml") || name.ends_with(".rels")
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeMap;
    use std::io::{Cursor, Read};

    use tempfile::tempdir;
    use zip::ZipArchive;

    use super::{validate_docx_package, validate_parts};
    use crate::{Document, Paragraph, Run};

    #[test]
    fn generated_package_has_valid_content_types_and_relationships() {
        let temp = tempdir().expect("temp dir");
        let path = temp.path().join("package.docx");
        Document::new()
            .add_paragraph(Paragraph::new().add_run(Run::from_text("validated")))
            .save(&path)
            .expect("save package");

        let report = validate_docx_package(&path).expect("validate package");
        assert!(report.valid, "{}", report.errors.join("; "));
        assert!(report.parts_checked >= 8);
        assert!(report.relationships_checked >= 6);
    }

    #[test]
    fn dangling_internal_relationship_is_reported() {
        let document =
            Document::new().add_paragraph(Paragraph::new().add_run(Run::from_text("validated")));
        let mut buffer = Cursor::new(Vec::new());
        document.save_to_writer(&mut buffer).expect("save package");
        buffer.set_position(0);

        let mut archive = ZipArchive::new(buffer).expect("open package");
        let mut parts = BTreeMap::new();
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index).expect("read entry");
            let name = entry.name().to_string();
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes).expect("read bytes");
            parts.insert(name, bytes);
        }
        let relationships = parts
            .get_mut("word/_rels/document.xml.rels")
            .expect("document relationships");
        *relationships = String::from_utf8_lossy(relationships)
            .replace("Target=\"styles.xml\"", "Target=\"missing.xml\"")
            .into_bytes();

        let report = validate_parts(&parts);
        assert!(!report.valid);
        assert!(report
            .errors
            .iter()
            .any(|error| error.contains("missing part 'word/missing.xml'")));
    }
}
