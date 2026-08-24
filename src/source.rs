//! Lightweight source-location recovery for semantic spec diagnostics.

use std::fs;
use std::path::Path;

use crate::{Result, SourceSpan, ValidationReport};

/// Attaches best-effort line and column spans to semantic validation issues.
///
/// The locator deliberately avoids a second parser-specific AST. It anchors a
/// semantic path to its block and terminal key, which keeps JSON, YAML, and
/// TOML diagnostics consistent while the serde model remains authoritative.
pub fn attach_source_spans(path: impl AsRef<Path>, report: &mut ValidationReport) -> Result<()> {
    let path = path.as_ref();
    let source = fs::read_to_string(path)?;
    let format = path
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or("yaml")
        .to_ascii_lowercase();
    attach_source_spans_from_str(&source, &format, report);
    Ok(())
}

/// Attaches spans from in-memory source text.
pub fn attach_source_spans_from_str(source: &str, format: &str, report: &mut ValidationReport) {
    let lines = source.lines().collect::<Vec<_>>();
    let block_anchors = collect_block_anchors(&lines, format);
    for issue in &mut report.issues {
        issue.source = locate_path(&lines, &block_anchors, &issue.path, format);
    }
}

fn locate_path(
    lines: &[&str],
    block_anchors: &[usize],
    path: &str,
    format: &str,
) -> Option<SourceSpan> {
    let key = terminal_key(path);
    let range = block_index(path).and_then(|index| {
        let start = *block_anchors.get(index)?;
        let end = block_anchors.get(index + 1).copied().unwrap_or(lines.len());
        Some(start..end)
    });

    if let Some(range) = range {
        if let Some(found) = find_key(lines, range.clone(), key, format) {
            return Some(found);
        }
        return line_span(lines, range.start, "type", format);
    }

    find_key(lines, 0..lines.len(), key, format)
}

fn collect_block_anchors(lines: &[&str], format: &str) -> Vec<usize> {
    lines
        .iter()
        .enumerate()
        .filter_map(|(index, line)| {
            let trimmed = line.trim_start();
            let is_anchor = match format {
                "toml" => trimmed == "[[blocks]]" || trimmed.starts_with("[[blocks."),
                "json" => trimmed.starts_with("\"type\"") || trimmed.contains("\"type\":"),
                _ => trimmed.starts_with("- type:") || trimmed.starts_with("type:"),
            };
            is_anchor.then_some(index)
        })
        .collect()
}

fn find_key(
    lines: &[&str],
    range: std::ops::Range<usize>,
    key: &str,
    format: &str,
) -> Option<SourceSpan> {
    for index in range {
        let line = *lines.get(index)?;
        if key_column(line, key, format).is_some() {
            return line_span(lines, index, key, format);
        }
    }
    None
}

fn line_span(lines: &[&str], index: usize, key: &str, format: &str) -> Option<SourceSpan> {
    let line = *lines.get(index)?;
    let start =
        key_column(line, key, format).unwrap_or_else(|| line.len() - line.trim_start().len());
    Some(SourceSpan {
        line: index + 1,
        column: start + 1,
        end_line: index + 1,
        end_column: start + key.len() + 1,
    })
}

fn key_column(line: &str, key: &str, format: &str) -> Option<usize> {
    match format {
        "json" => line.find(&format!("\"{key}\"")),
        "toml" => {
            let trimmed = line.trim_start();
            (trimmed.starts_with(&format!("{key} ")) || trimmed.starts_with(&format!("{key}=")))
                .then_some(line.len() - trimmed.len())
        }
        _ => {
            let trimmed = line.trim_start().trim_start_matches("- ");
            trimmed
                .starts_with(&format!("{key}:"))
                .then(|| line.find(key).unwrap_or(0))
        }
    }
}

fn block_index(path: &str) -> Option<usize> {
    let tail = path.strip_prefix("blocks[")?;
    tail.split(']').next()?.parse().ok()
}

fn terminal_key(path: &str) -> &str {
    let terminal = path.rsplit('.').next().unwrap_or(path);
    terminal.split('[').next().unwrap_or(terminal)
}

#[cfg(test)]
mod tests {
    use crate::{SourceSpan, ValidationReport};

    use super::attach_source_spans_from_str;

    #[test]
    fn locates_yaml_block_issue_by_semantic_path() {
        let source = "version: 1\nblocks:\n  - type: body\n    text: hello\n  - type: image\n    path: missing.png\n";
        let mut report = ValidationReport::default();
        report.push_error("blocks[1].path", "missing");
        attach_source_spans_from_str(source, "yaml", &mut report);
        assert_eq!(
            report.issues[0].source,
            Some(SourceSpan {
                line: 6,
                column: 5,
                end_line: 6,
                end_column: 9,
            })
        );
    }
}
