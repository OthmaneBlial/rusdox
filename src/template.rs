//! Word-native DOCX template inspection and rendering.
//!
//! Template expansion edits only textual OOXML parts. Unchanged package parts
//! retain their exact bytes so designer-authored styles, media, relationships,
//! and unsupported package extensions survive the render.

use std::borrow::Cow;
use std::collections::BTreeMap;
use std::fs;
use std::io::{BufReader, Write};
use std::path::Path;

use quick_xml::escape::{escape, unescape};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

use crate::package_validate::read_docx_parts_with_limits;
use crate::{DocxError, InputLimits, Result};

/// Current version of the Word-native placeholder contract.
pub const TEMPLATE_SYNTAX_VERSION: &str = "1";

/// Severity of a template inspection or render diagnostic.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum TemplateDiagnosticSeverity {
    /// Selects the warning form.
    Warning,
    /// Selects the error form.
    Error,
}

/// Actionable template diagnostic with an OOXML part and document location.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct TemplateDiagnostic {
    /// Severity.
    pub severity: TemplateDiagnosticSeverity,
    /// Part.
    pub part: String,
    /// Location.
    pub location: String,
    /// Placeholder.
    pub placeholder: String,
    /// Message.
    pub message: String,
    /// Suggestion.
    pub suggestion: String,
}

/// One placeholder discovered in a Word-native template.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct TemplatePlaceholder {
    /// Part.
    pub part: String,
    /// Location.
    pub location: String,
    /// Expression.
    pub expression: String,
    /// Kind.
    pub kind: String,
}

/// Machine-readable result of `template inspect`.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct TemplateInspection {
    /// Syntax version.
    pub syntax_version: String,
    /// Ordered placeholders.
    pub placeholders: Vec<TemplatePlaceholder>,
    /// Ordered diagnostics.
    pub diagnostics: Vec<TemplateDiagnostic>,
}

impl TemplateInspection {
    /// Returns whether the value has errors.
    pub fn has_errors(&self) -> bool {
        self.diagnostics
            .iter()
            .any(|diagnostic| diagnostic.severity == TemplateDiagnosticSeverity::Error)
    }
}

/// Machine-readable result of a Word-native template render.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct TemplateRenderReport {
    /// Syntax version.
    pub syntax_version: String,
    /// Output.
    pub output: String,
    /// Whether an output package was written.
    pub written: bool,
    /// Whether strict missing-value validation was enabled.
    pub strict: bool,
    /// Replacements.
    pub replacements: usize,
    /// Expanded blocks.
    pub expanded_blocks: usize,
    /// Ordered diagnostics.
    pub diagnostics: Vec<TemplateDiagnostic>,
}

impl TemplateRenderReport {
    /// Returns whether the value has errors.
    pub fn has_errors(&self) -> bool {
        self.diagnostics
            .iter()
            .any(|diagnostic| diagnostic.severity == TemplateDiagnosticSeverity::Error)
    }
}

/// A DOCX package whose textual OOXML parts can be expanded from JSON data.
#[derive(Debug, Clone)]
pub struct DocxTemplate {
    parts: BTreeMap<String, Vec<u8>>,
}

impl DocxTemplate {
    /// Opens a DOCX template with the default untrusted-input ceilings.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_limits(path, InputLimits::default())
    }

    /// Opens a DOCX template with explicit untrusted-input ceilings.
    pub fn open_with_limits(path: impl AsRef<Path>, limits: InputLimits) -> Result<Self> {
        let file = fs::File::open(path)?;
        let parts = read_docx_parts_with_limits(BufReader::new(file), limits)?;
        if !parts.contains_key("word/document.xml") {
            return Err(DocxError::parse(
                "DOCX template is missing word/document.xml",
            ));
        }
        Ok(Self { parts })
    }

    /// Inspects every supported textual part without modifying the template.
    pub fn inspect(&self) -> TemplateInspection {
        let mut placeholders = Vec::new();
        let mut diagnostics = Vec::new();
        for (part, bytes) in self.template_parts() {
            let Ok(xml) = std::str::from_utf8(bytes) else {
                diagnostics.push(diagnostic(
                    TemplateDiagnosticSeverity::Error,
                    part,
                    "part",
                    "",
                    "template part is not UTF-8 XML",
                    "save the DOCX with UTF-8 OOXML parts",
                ));
                continue;
            };
            let segments = split_segments(xml);
            inspect_segments(part, &segments, &mut placeholders, &mut diagnostics);
        }
        TemplateInspection {
            syntax_version: TEMPLATE_SYNTAX_VERSION.to_string(),
            placeholders,
            diagnostics,
        }
    }

    /// Renders JSON data into a new DOCX while preserving untouched package parts.
    ///
    /// Missing values become empty strings with warnings by default. With
    /// `strict = true`, they are errors and no output is written.
    pub fn render_to_path(
        &self,
        data: &Value,
        output: impl AsRef<Path>,
        strict: bool,
    ) -> Result<TemplateRenderReport> {
        let output = output.as_ref();
        let mut parts = self.parts.clone();
        let mut state = RenderState {
            strict,
            replacements: 0,
            expanded_blocks: 0,
            partial_stack: Vec::new(),
            diagnostics: Vec::new(),
        };
        let context = RenderContext {
            root: data,
            current: data,
            index: None,
        };

        for (part, bytes) in self.template_parts() {
            let xml = match std::str::from_utf8(bytes) {
                Ok(xml) => xml,
                Err(_) => {
                    state.diagnostics.push(diagnostic(
                        TemplateDiagnosticSeverity::Error,
                        part,
                        "part",
                        "",
                        "template part is not UTF-8 XML",
                        "save the DOCX with UTF-8 OOXML parts",
                    ));
                    continue;
                }
            };
            let segments = split_segments(xml);
            let rendered = expand_segments(part, &segments, context, &mut state);
            parts.insert(part.to_string(), join_segments(&rendered).into_bytes());
        }

        let written = !state
            .diagnostics
            .iter()
            .any(|diagnostic| diagnostic.severity == TemplateDiagnosticSeverity::Error);
        if written {
            crate::io_utils::atomic_write_with(output, |file| write_package(file, &parts))?;
        }

        Ok(TemplateRenderReport {
            syntax_version: TEMPLATE_SYNTAX_VERSION.to_string(),
            output: output.display().to_string(),
            written,
            strict,
            replacements: state.replacements,
            expanded_blocks: state.expanded_blocks,
            diagnostics: state.diagnostics,
        })
    }

    fn template_parts(&self) -> impl Iterator<Item = (&str, &[u8])> {
        self.parts.iter().filter_map(|(name, bytes)| {
            is_template_part(name).then_some((name.as_str(), bytes.as_slice()))
        })
    }
}

fn write_package(writer: &mut fs::File, parts: &BTreeMap<String, Vec<u8>>) -> Result<()> {
    let mut archive = ZipWriter::new(writer);
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    for (name, bytes) in parts {
        archive.start_file(name, options)?;
        archive.write_all(bytes)?;
    }
    archive.finish()?;
    Ok(())
}

fn is_template_part(name: &str) -> bool {
    name == "word/document.xml"
        || (name.starts_with("word/header") && name.ends_with(".xml"))
        || (name.starts_with("word/footer") && name.ends_with(".xml"))
        || matches!(
            name,
            "word/footnotes.xml"
                | "word/endnotes.xml"
                | "word/comments.xml"
                | "word/glossary/document.xml"
        )
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SegmentKind {
    Raw,
    Paragraph,
    TableRow,
}

#[derive(Debug, Clone)]
struct Segment {
    kind: SegmentKind,
    xml: String,
    location: String,
}

fn split_segments(xml: &str) -> Vec<Segment> {
    let mut segments = Vec::new();
    let mut cursor = 0;
    let mut paragraph = 0usize;
    let mut row = 0usize;

    while cursor < xml.len() {
        let next_row = find_element_start(xml, cursor, "w:tr");
        let next_paragraph = find_element_start(xml, cursor, "w:p");
        let next = match (next_row, next_paragraph) {
            (Some(row), Some(paragraph)) if row <= paragraph => Some((row, SegmentKind::TableRow)),
            (Some(_), Some(paragraph)) => Some((paragraph, SegmentKind::Paragraph)),
            (Some(row), None) => Some((row, SegmentKind::TableRow)),
            (None, Some(paragraph)) => Some((paragraph, SegmentKind::Paragraph)),
            (None, None) => None,
        };
        let Some((start, kind)) = next else {
            push_raw(&mut segments, &xml[cursor..]);
            break;
        };
        if start > cursor {
            push_raw(&mut segments, &xml[cursor..start]);
        }
        let tag = if kind == SegmentKind::TableRow {
            "w:tr"
        } else {
            "w:p"
        };
        let Some(end) = find_element_end(xml, start, tag) else {
            push_raw(&mut segments, &xml[start..]);
            break;
        };
        let location = match kind {
            SegmentKind::Paragraph => {
                paragraph += 1;
                format!("paragraph {paragraph}")
            }
            SegmentKind::TableRow => {
                row += 1;
                format!("table row {row}")
            }
            SegmentKind::Raw => unreachable!(),
        };
        segments.push(Segment {
            kind,
            xml: xml[start..end].to_string(),
            location,
        });
        cursor = end;
    }
    segments
}

fn push_raw(segments: &mut Vec<Segment>, xml: &str) {
    if !xml.is_empty() {
        segments.push(Segment {
            kind: SegmentKind::Raw,
            xml: xml.to_string(),
            location: "part".to_string(),
        });
    }
}

fn find_element_start(xml: &str, mut cursor: usize, tag: &str) -> Option<usize> {
    let needle = format!("<{tag}");
    while let Some(relative) = xml[cursor..].find(&needle) {
        let start = cursor + relative;
        let boundary = xml.as_bytes().get(start + needle.len()).copied();
        if boundary.is_some_and(|value| value == b'>' || value.is_ascii_whitespace()) {
            return Some(start);
        }
        cursor = start + needle.len();
    }
    None
}

fn find_element_end(xml: &str, start: usize, tag: &str) -> Option<usize> {
    let open_end = xml[start..].find('>')? + start + 1;
    if xml.as_bytes().get(open_end.saturating_sub(2)) == Some(&b'/') {
        return Some(open_end);
    }
    let close = format!("</{tag}>");
    let mut depth = 1usize;
    let mut cursor = open_end;
    while cursor < xml.len() {
        let next_open = find_element_start(xml, cursor, tag);
        let next_close = xml[cursor..].find(&close).map(|value| cursor + value);
        match (next_open, next_close) {
            (_, Some(close_start)) if next_open.is_none_or(|open| close_start < open) => {
                depth -= 1;
                let end = close_start + close.len();
                if depth == 0 {
                    return Some(end);
                }
                cursor = end;
            }
            (None, Some(close_start)) => {
                depth -= 1;
                let end = close_start + close.len();
                if depth == 0 {
                    return Some(end);
                }
                cursor = end;
            }
            (Some(open), _) => {
                depth += 1;
                cursor = xml[open..].find('>')? + open + 1;
            }
            (None, None) => return None,
        }
    }
    None
}

fn inspect_segments(
    part: &str,
    segments: &[Segment],
    placeholders: &mut Vec<TemplatePlaceholder>,
    diagnostics: &mut Vec<TemplateDiagnostic>,
) {
    let mut markers = Vec::new();
    for segment in segments
        .iter()
        .filter(|segment| segment.kind != SegmentKind::Raw)
    {
        let text = visible_text(&segment.xml).unwrap_or_default();
        let spans = placeholder_spans(&text);
        if text.contains("{{") && spans.is_empty() {
            diagnostics.push(diagnostic(
                TemplateDiagnosticSeverity::Error,
                part,
                &segment.location,
                text.trim(),
                "placeholder is not closed with `}}`",
                "keep the complete placeholder inside one paragraph or table row",
            ));
        }
        for span in spans {
            let expression = text[span.inner_start..span.inner_end].trim().to_string();
            let marker = BlockMarker::parse(&expression);
            let kind = marker
                .as_ref()
                .map(BlockMarker::kind)
                .unwrap_or("value")
                .to_string();
            placeholders.push(TemplatePlaceholder {
                part: part.to_string(),
                location: segment.location.clone(),
                expression: expression.clone(),
                kind,
            });
            if let Some(marker) = marker {
                if text.trim() != format!("{{{{{expression}}}}}") {
                    diagnostics.push(diagnostic(
                        TemplateDiagnosticSeverity::Error,
                        part,
                        &segment.location,
                        &expression,
                        "block marker must occupy a complete paragraph or table row",
                        "move the marker into its own paragraph or row",
                    ));
                }
                markers.push((marker, segment.location.clone(), expression));
            }
        }
    }
    validate_marker_stack(part, &markers, diagnostics);
}

fn validate_marker_stack(
    part: &str,
    markers: &[(BlockMarker, String, String)],
    diagnostics: &mut Vec<TemplateDiagnostic>,
) {
    let mut stack = Vec::new();
    for (marker, location, expression) in markers {
        match marker {
            BlockMarker::Each(_) => stack.push(MarkerType::Each),
            BlockMarker::If(_) => stack.push(MarkerType::If),
            BlockMarker::Else if stack.last() != Some(&MarkerType::If) => {
                diagnostics.push(diagnostic(
                    TemplateDiagnosticSeverity::Error,
                    part,
                    location,
                    expression,
                    "`else` must be inside an `if` block",
                    "add a matching `{{#if path}}` before this marker",
                ))
            }
            BlockMarker::Else => {}
            BlockMarker::CloseEach => close_marker(
                part,
                location,
                expression,
                MarkerType::Each,
                &mut stack,
                diagnostics,
            ),
            BlockMarker::CloseIf => close_marker(
                part,
                location,
                expression,
                MarkerType::If,
                &mut stack,
                diagnostics,
            ),
        }
    }
    for marker in stack.into_iter().rev() {
        diagnostics.push(diagnostic(
            TemplateDiagnosticSeverity::Error,
            part,
            "part",
            marker.open_expression(),
            "block marker is not closed",
            marker.close_suggestion(),
        ));
    }
}

fn close_marker(
    part: &str,
    location: &str,
    expression: &str,
    expected: MarkerType,
    stack: &mut Vec<MarkerType>,
    diagnostics: &mut Vec<TemplateDiagnostic>,
) {
    if stack.pop() != Some(expected) {
        diagnostics.push(diagnostic(
            TemplateDiagnosticSeverity::Error,
            part,
            location,
            expression,
            "closing marker does not match the active block",
            expected.open_suggestion(),
        ));
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum BlockMarker {
    Each(String),
    If(String),
    Else,
    CloseEach,
    CloseIf,
}

impl BlockMarker {
    fn parse(expression: &str) -> Option<Self> {
        if let Some(path) = expression.strip_prefix("#each ") {
            Some(Self::Each(path.trim().to_string()))
        } else if let Some(path) = expression.strip_prefix("#if ") {
            Some(Self::If(path.trim().to_string()))
        } else {
            match expression {
                "else" => Some(Self::Else),
                "/each" => Some(Self::CloseEach),
                "/if" => Some(Self::CloseIf),
                _ => None,
            }
        }
    }

    fn kind(&self) -> &'static str {
        match self {
            Self::Each(_) => "each_start",
            Self::If(_) => "if_start",
            Self::Else => "else",
            Self::CloseEach => "each_end",
            Self::CloseIf => "if_end",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum MarkerType {
    Each,
    If,
}

impl MarkerType {
    fn open_expression(self) -> &'static str {
        match self {
            Self::Each => "#each",
            Self::If => "#if",
        }
    }

    fn close_suggestion(self) -> &'static str {
        match self {
            Self::Each => "add `{{/each}}` in a matching paragraph or row",
            Self::If => "add `{{/if}}` in a matching paragraph or row",
        }
    }

    fn open_suggestion(self) -> &'static str {
        match self {
            Self::Each => "add a matching `{{#each path}}` before this marker",
            Self::If => "add a matching `{{#if path}}` before this marker",
        }
    }
}

#[derive(Clone, Copy)]
struct RenderContext<'a> {
    root: &'a Value,
    current: &'a Value,
    index: Option<usize>,
}

struct RenderState {
    strict: bool,
    replacements: usize,
    expanded_blocks: usize,
    partial_stack: Vec<String>,
    diagnostics: Vec<TemplateDiagnostic>,
}

fn expand_segments(
    part: &str,
    segments: &[Segment],
    context: RenderContext<'_>,
    state: &mut RenderState,
) -> Vec<Segment> {
    let mut output = Vec::new();
    let mut index = 0usize;
    while index < segments.len() {
        let segment = &segments[index];
        let marker = segment_marker(segment);
        match marker {
            Some(BlockMarker::Each(path)) => {
                let Some(block) = matching_block(segments, index, MarkerType::Each) else {
                    state.diagnostics.push(diagnostic(
                        TemplateDiagnosticSeverity::Error,
                        part,
                        &segment.location,
                        &format!("#each {path}"),
                        "loop block is not closed",
                        "add `{{/each}}` in a matching paragraph or row",
                    ));
                    index += 1;
                    continue;
                };
                match resolve_path(context, &path) {
                    Some(Value::Array(items)) => {
                        for (item_index, item) in items.iter().enumerate() {
                            let child = RenderContext {
                                root: context.root,
                                current: item,
                                index: Some(item_index),
                            };
                            output.extend(expand_segments(
                                part,
                                &segments[index + 1..block.end],
                                child,
                                state,
                            ));
                            state.expanded_blocks += 1;
                        }
                    }
                    Some(_) => state.diagnostics.push(diagnostic(
                        TemplateDiagnosticSeverity::Error,
                        part,
                        &segment.location,
                        &path,
                        "loop value is not an array",
                        "provide a JSON array or remove the `#each` block",
                    )),
                    None => missing_value(part, segment, &path, state),
                }
                index = block.end + 1;
            }
            Some(BlockMarker::If(path)) => {
                let Some(block) = matching_block(segments, index, MarkerType::If) else {
                    state.diagnostics.push(diagnostic(
                        TemplateDiagnosticSeverity::Error,
                        part,
                        &segment.location,
                        &format!("#if {path}"),
                        "condition block is not closed",
                        "add `{{/if}}` in a matching paragraph or row",
                    ));
                    index += 1;
                    continue;
                };
                let condition = match resolve_path(context, &path) {
                    Some(value) => truthy(value),
                    None => {
                        missing_value(part, segment, &path, state);
                        false
                    }
                };
                let (start, end) = if condition {
                    (index + 1, block.alternative.unwrap_or(block.end))
                } else {
                    (
                        block.alternative.map_or(block.end, |value| value + 1),
                        block.end,
                    )
                };
                output.extend(expand_segments(part, &segments[start..end], context, state));
                state.expanded_blocks += 1;
                index = block.end + 1;
            }
            Some(BlockMarker::Else | BlockMarker::CloseEach | BlockMarker::CloseIf) => {
                state.diagnostics.push(diagnostic(
                    TemplateDiagnosticSeverity::Error,
                    part,
                    &segment.location,
                    &visible_text(&segment.xml).unwrap_or_default(),
                    "unexpected block marker",
                    "check the nesting and matching opening marker",
                ));
                index += 1;
            }
            None if segment.kind == SegmentKind::Raw => {
                output.push(segment.clone());
                index += 1;
            }
            None => {
                let mut rendered = segment.clone();
                rendered.xml = render_segment(part, segment, context, state);
                output.push(rendered);
                index += 1;
            }
        }
    }
    output
}

struct MatchingBlock {
    end: usize,
    alternative: Option<usize>,
}

fn matching_block(
    segments: &[Segment],
    opening: usize,
    expected: MarkerType,
) -> Option<MatchingBlock> {
    let mut stack = vec![expected];
    let mut alternative = None;
    let opening_kind = segments.get(opening)?.kind;
    for (index, segment) in segments.iter().enumerate().skip(opening + 1) {
        match segment_marker(segment) {
            Some(BlockMarker::Each(_)) => stack.push(MarkerType::Each),
            Some(BlockMarker::If(_)) => stack.push(MarkerType::If),
            Some(BlockMarker::Else) if stack == [MarkerType::If] => alternative = Some(index),
            Some(BlockMarker::CloseEach) if stack.last() == Some(&MarkerType::Each) => {
                stack.pop();
            }
            Some(BlockMarker::CloseIf) if stack.last() == Some(&MarkerType::If) => {
                stack.pop();
            }
            _ => {}
        }
        if stack.is_empty() {
            if segment.kind != opening_kind {
                return None;
            }
            return Some(MatchingBlock {
                end: index,
                alternative,
            });
        }
    }
    None
}

fn segment_marker(segment: &Segment) -> Option<BlockMarker> {
    if segment.kind == SegmentKind::Raw {
        return None;
    }
    let text = visible_text(&segment.xml).ok()?;
    let trimmed = text.trim();
    let expression = trimmed.strip_prefix("{{")?.strip_suffix("}}")?.trim();
    BlockMarker::parse(expression)
}

fn render_segment(
    part: &str,
    segment: &Segment,
    context: RenderContext<'_>,
    state: &mut RenderState,
) -> String {
    let Ok(nodes) = text_nodes(&segment.xml) else {
        state.diagnostics.push(diagnostic(
            TemplateDiagnosticSeverity::Error,
            part,
            &segment.location,
            "",
            "text node contains invalid XML escaping",
            "repair the Word text in this paragraph or row",
        ));
        return segment.xml.clone();
    };
    let plain = nodes
        .iter()
        .map(|node| node.text.as_str())
        .collect::<String>();
    let spans = placeholder_spans(&plain);
    if spans.is_empty() {
        if plain.contains("{{") {
            state.diagnostics.push(diagnostic(
                TemplateDiagnosticSeverity::Error,
                part,
                &segment.location,
                plain.trim(),
                "placeholder is not closed with `}}`",
                "close the placeholder or remove the opening braces",
            ));
        }
        return segment.xml.clone();
    }

    let mut output_nodes = vec![String::new(); nodes.len()];
    let mut cursor = 0usize;
    for span in spans {
        distribute_original(&plain, &nodes, cursor, span.start, &mut output_nodes);
        let expression = plain[span.inner_start..span.inner_end].trim();
        if BlockMarker::parse(expression).is_some() {
            state.diagnostics.push(diagnostic(
                TemplateDiagnosticSeverity::Error,
                part,
                &segment.location,
                expression,
                "block marker must occupy a complete paragraph or table row",
                "move the marker into its own paragraph or row",
            ));
        } else {
            let replacement = evaluate_expression(part, segment, expression, context, state);
            if let Some(node_index) = node_for_offset(&nodes, span.start) {
                output_nodes[node_index].push_str(&replacement);
                state.replacements += 1;
            }
        }
        cursor = span.end;
    }
    distribute_original(&plain, &nodes, cursor, plain.len(), &mut output_nodes);

    let mut rendered = segment.xml.clone();
    for (node, value) in nodes.iter().zip(output_nodes).rev() {
        let escaped: Cow<'_, str> = escape(&value);
        rendered.replace_range(node.source_start..node.source_end, &escaped);
    }
    rendered
}

#[derive(Debug)]
struct TextNode {
    source_start: usize,
    source_end: usize,
    plain_start: usize,
    plain_end: usize,
    text: String,
}

fn text_nodes(xml: &str) -> Result<Vec<TextNode>> {
    let mut nodes = Vec::new();
    let mut cursor = 0usize;
    let mut plain_cursor = 0usize;
    while let Some(start) = find_element_start(xml, cursor, "w:t") {
        let content_start = xml[start..]
            .find('>')
            .map(|value| start + value + 1)
            .ok_or_else(|| DocxError::parse("invalid w:t element"))?;
        if xml.as_bytes().get(content_start.saturating_sub(2)) == Some(&b'/') {
            cursor = content_start;
            continue;
        }
        let close = "</w:t>";
        let content_end = xml[content_start..]
            .find(close)
            .map(|value| content_start + value)
            .ok_or_else(|| DocxError::parse("unclosed w:t element"))?;
        let text = unescape(&xml[content_start..content_end])?.into_owned();
        let plain_end = plain_cursor + text.len();
        nodes.push(TextNode {
            source_start: content_start,
            source_end: content_end,
            plain_start: plain_cursor,
            plain_end,
            text,
        });
        plain_cursor = plain_end;
        cursor = content_end + close.len();
    }
    Ok(nodes)
}

fn visible_text(xml: &str) -> Result<String> {
    Ok(text_nodes(xml)?.into_iter().map(|node| node.text).collect())
}

fn distribute_original(
    plain: &str,
    nodes: &[TextNode],
    start: usize,
    end: usize,
    outputs: &mut [String],
) {
    if start >= end {
        return;
    }
    for (index, node) in nodes.iter().enumerate() {
        let overlap_start = start.max(node.plain_start);
        let overlap_end = end.min(node.plain_end);
        if overlap_start < overlap_end {
            outputs[index].push_str(&plain[overlap_start..overlap_end]);
        }
    }
}

fn node_for_offset(nodes: &[TextNode], offset: usize) -> Option<usize> {
    nodes
        .iter()
        .position(|node| node.plain_start <= offset && offset < node.plain_end)
        .or_else(|| {
            (!nodes.is_empty() && offset == nodes.last()?.plain_end).then_some(nodes.len() - 1)
        })
}

#[derive(Debug)]
struct PlaceholderSpan {
    start: usize,
    end: usize,
    inner_start: usize,
    inner_end: usize,
}

fn placeholder_spans(text: &str) -> Vec<PlaceholderSpan> {
    let mut spans = Vec::new();
    let mut cursor = 0usize;
    while let Some(open) = text[cursor..].find("{{").map(|value| cursor + value) {
        let Some(close) = text[open + 2..].find("}}").map(|value| open + 2 + value) else {
            break;
        };
        spans.push(PlaceholderSpan {
            start: open,
            end: close + 2,
            inner_start: open + 2,
            inner_end: close,
        });
        cursor = close + 2;
    }
    spans
}

fn evaluate_expression(
    part: &str,
    segment: &Segment,
    expression: &str,
    context: RenderContext<'_>,
    state: &mut RenderState,
) -> String {
    if let Some(name) = expression.strip_prefix('>') {
        let name = name.trim();
        let partial = context
            .root
            .get("$partials")
            .and_then(|partials| partials.get(name))
            .and_then(Value::as_str)
            .map(ToString::to_string);
        let Some(partial) = partial else {
            missing_value(part, segment, &format!("> {name}"), state);
            return String::new();
        };
        if state.partial_stack.iter().any(|active| active == name) {
            state.diagnostics.push(diagnostic(
                TemplateDiagnosticSeverity::Error,
                part,
                &segment.location,
                expression,
                "recursive partial reference detected",
                "remove the partial cycle; partial expansion must be finite",
            ));
            return String::new();
        }
        state.partial_stack.push(name.to_string());
        let rendered = render_inline_text(part, segment, &partial, context, state);
        state.partial_stack.pop();
        return rendered;
    }

    let mut pieces = expression.split('|').map(str::trim);
    let path = pieces.next().unwrap_or_default();
    let mut value = if path == "@index" {
        context.index.map(|index| Value::from(index + 1))
    } else {
        resolve_path(context, path).cloned()
    };
    for filter in pieces {
        if !is_known_filter(filter) {
            state.diagnostics.push(diagnostic(
                TemplateDiagnosticSeverity::Error,
                part,
                &segment.location,
                expression,
                &format!("unknown template filter `{filter}`"),
                "use `upper`, `lower`, `title`, `trim`, or `default(\"text\")`",
            ));
            return String::new();
        }
        value = apply_filter(value, filter);
    }
    match value {
        Some(Value::String(value)) => value,
        Some(Value::Number(value)) => value.to_string(),
        Some(Value::Bool(value)) => value.to_string(),
        Some(Value::Null) | None => {
            missing_value(part, segment, path, state);
            String::new()
        }
        Some(Value::Array(_) | Value::Object(_)) => {
            state.diagnostics.push(diagnostic(
                TemplateDiagnosticSeverity::Error,
                part,
                &segment.location,
                expression,
                "value is structured and cannot be inserted as text",
                "select a nested scalar path or use an `#each` block",
            ));
            String::new()
        }
    }
}

fn render_inline_text(
    part: &str,
    segment: &Segment,
    template: &str,
    context: RenderContext<'_>,
    state: &mut RenderState,
) -> String {
    let spans = placeholder_spans(template);
    if spans.is_empty() {
        return template.to_string();
    }
    let mut rendered = String::new();
    let mut cursor = 0usize;
    for span in spans {
        rendered.push_str(&template[cursor..span.start]);
        let expression = template[span.inner_start..span.inner_end].trim();
        if BlockMarker::parse(expression).is_some() {
            state.diagnostics.push(diagnostic(
                TemplateDiagnosticSeverity::Error,
                part,
                &segment.location,
                expression,
                "partials support inline values, not block markers",
                "move loops and conditions into complete Word paragraphs or rows",
            ));
        } else {
            rendered.push_str(&evaluate_expression(
                part, segment, expression, context, state,
            ));
            state.replacements += 1;
        }
        cursor = span.end;
    }
    rendered.push_str(&template[cursor..]);
    rendered
}

fn is_known_filter(filter: &str) -> bool {
    matches!(filter, "upper" | "lower" | "trim" | "title")
        || (filter.starts_with("default(") && filter.ends_with(')'))
}

fn apply_filter(value: Option<Value>, filter: &str) -> Option<Value> {
    if let Some(argument) = filter
        .strip_prefix("default(")
        .and_then(|value| value.strip_suffix(')'))
    {
        if value.as_ref().is_none_or(|value| value.is_null()) {
            return Some(Value::String(
                argument.trim().trim_matches(['\'', '"']).to_string(),
            ));
        }
        return value;
    }
    let text = match value {
        Some(Value::String(value)) => value,
        Some(Value::Number(value)) => value.to_string(),
        Some(Value::Bool(value)) => value.to_string(),
        other => return other,
    };
    Some(Value::String(match filter {
        "upper" => text.to_uppercase(),
        "lower" => text.to_lowercase(),
        "trim" => text.trim().to_string(),
        "title" => text
            .split_whitespace()
            .map(|word| {
                let mut chars = word.chars();
                chars
                    .next()
                    .map(|first| first.to_uppercase().collect::<String>() + chars.as_str())
                    .unwrap_or_default()
            })
            .collect::<Vec<_>>()
            .join(" "),
        _ => text,
    }))
}

fn resolve_path<'a>(context: RenderContext<'a>, path: &str) -> Option<&'a Value> {
    if path == "this" || path == "." {
        return Some(context.current);
    }
    resolve_from(context.current, path).or_else(|| resolve_from(context.root, path))
}

fn resolve_from<'a>(value: &'a Value, path: &str) -> Option<&'a Value> {
    path.split('.')
        .filter(|segment| !segment.is_empty())
        .try_fold(value, |current, segment| match current {
            Value::Object(map) => map.get(segment),
            Value::Array(items) => segment
                .parse::<usize>()
                .ok()
                .and_then(|index| items.get(index)),
            _ => None,
        })
}

fn truthy(value: &Value) -> bool {
    match value {
        Value::Null => false,
        Value::Bool(value) => *value,
        Value::Number(value) => value.as_f64().is_some_and(|value| value != 0.0),
        Value::String(value) => !value.is_empty(),
        Value::Array(value) => !value.is_empty(),
        Value::Object(value) => !value.is_empty(),
    }
}

fn missing_value(part: &str, segment: &Segment, path: &str, state: &mut RenderState) {
    state.diagnostics.push(diagnostic(
        if state.strict {
            TemplateDiagnosticSeverity::Error
        } else {
            TemplateDiagnosticSeverity::Warning
        },
        part,
        &segment.location,
        path,
        "value is missing or null",
        "provide the JSON value, use `default(\"text\")`, or disable strict mode",
    ));
}

fn join_segments(segments: &[Segment]) -> String {
    segments
        .iter()
        .map(|segment| segment.xml.as_str())
        .collect()
}

fn diagnostic(
    severity: TemplateDiagnosticSeverity,
    part: &str,
    location: &str,
    placeholder: &str,
    message: &str,
    suggestion: &str,
) -> TemplateDiagnostic {
    TemplateDiagnostic {
        severity,
        part: part.to_string(),
        location: location.to_string(),
        placeholder: placeholder.to_string(),
        message: message.to_string(),
        suggestion: suggestion.to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::{
        apply_filter, placeholder_spans, render_segment, split_segments, visible_text,
        RenderContext, RenderState,
    };
    use serde_json::json;

    #[test]
    fn placeholders_can_span_word_text_runs() {
        let xml =
            r#"<w:p><w:r><w:t>Hello {{ custo</w:t></w:r><w:r><w:t>mer.name }}</w:t></w:r></w:p>"#;
        let segments = split_segments(xml);
        assert_eq!(segments.len(), 1);
        let text = visible_text(&segments[0].xml).expect("visible text");
        let spans = placeholder_spans(&text);
        assert_eq!(spans.len(), 1);
        assert_eq!(
            &text[spans[0].inner_start..spans[0].inner_end],
            " customer.name "
        );
    }

    #[test]
    fn default_and_case_filters_are_deterministic() {
        assert_eq!(
            apply_filter(None, "default(\"unknown\")"),
            Some(json!("unknown"))
        );
        assert_eq!(
            apply_filter(Some(json!("hello")), "upper"),
            Some(json!("HELLO"))
        );
    }

    #[test]
    fn rendering_across_runs_preserves_surrounding_run_nodes() {
        let xml = r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hello {{ custo</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>mer.name }}!</w:t></w:r></w:p>"#;
        let segment = split_segments(xml).remove(0);
        let data = json!({"customer": {"name": "Ada"}});
        let mut state = RenderState {
            strict: true,
            replacements: 0,
            expanded_blocks: 0,
            partial_stack: Vec::new(),
            diagnostics: Vec::new(),
        };
        let rendered = render_segment(
            "word/document.xml",
            &segment,
            RenderContext {
                root: &data,
                current: &data,
                index: None,
            },
            &mut state,
        );
        assert_eq!(visible_text(&rendered).expect("visible text"), "Hello Ada!");
        assert!(rendered.contains("<w:b/>"));
        assert!(rendered.contains("<w:i/>"));
        assert_eq!(state.replacements, 1);
        assert!(state.diagnostics.is_empty());
    }
}
