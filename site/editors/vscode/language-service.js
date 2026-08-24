"use strict";

const BLOCK_TYPES = [
  "body",
  "bullets",
  "centered_note",
  "chart",
  "cover_title",
  "hero",
  "image",
  "include",
  "label_values",
  "logo",
  "metrics",
  "numbered",
  "page_break",
  "page_heading",
  "paragraph",
  "repeat",
  "section",
  "section_break",
  "signature",
  "spacer",
  "subtitle",
  "table",
  "table_of_contents",
  "tagline",
  "title",
  "when"
];

const HOVERS = {
  version: "Required spec contract version. RusDox currently supports version: 1.",
  variables: "Local values for nested paths, deterministic filters, repeat blocks, and when blocks.",
  blocks: "Ordered document content. Native blocks and the bounded include, repeat, and when helpers are supported.",
  output_name: "Optional artifact stem. The input file stem is used when omitted.",
  type: "Block discriminator. Completion lists every supported native and authoring block.",
  path: "A local asset/include path, or a nested variable path inside a when block.",
  equals: "Optional scalar equality used by a when block. Without it, RusDox applies documented truthiness.",
  otherwise: "Optional deterministic fallback branch for a when block."
};

const FILTERS = new Set(["upper", "lower", "title", "trim"]);

function analyze(text) {
  const diagnostics = [];
  const lines = text.split(/\r?\n/);
  const versionLine = lines.findIndex((line) => /^\s*["']?version["']?\s*[:=]/.test(line));
  if (versionLine < 0) {
    diagnostics.push({
      severity: "warning",
      line: 0,
      start: 0,
      end: Math.min((lines[0] || "").length || 1, 1),
      message: "Add version: 1; run rusdox migrate <file> --in-place for legacy specs."
    });
  } else {
    const match = lines[versionLine].match(/version["']?\s*[:=]\s*["']?(\d+)/);
    if (match && match[1] !== "1") {
      diagnostics.push({
        severity: "error",
        line: versionLine,
        start: lines[versionLine].indexOf(match[1]),
        end: lines[versionLine].indexOf(match[1]) + match[1].length,
        message: "Unsupported RusDox spec version " + match[1] + "; this extension supports version 1."
      });
    }
  }

  lines.forEach((line, lineIndex) => {
    const type = line.match(/["']?type["']?\s*[:=]\s*["']?([a-z_]+)/);
    if (type && !BLOCK_TYPES.includes(type[1])) {
      const start = line.indexOf(type[1]);
      diagnostics.push({
        severity: "error",
        line: lineIndex,
        start,
        end: start + type[1].length,
        message: "Unknown block type '" + type[1] + "'."
      });
    }

    let cursor = 0;
    while ((cursor = line.indexOf("{{", cursor)) >= 0) {
      if (line.startsWith("{{{{", cursor)) {
        cursor += 4;
        continue;
      }
      const end = line.indexOf("}}", cursor + 2);
      if (end < 0) {
        diagnostics.push({
          severity: "error",
          line: lineIndex,
          start: cursor,
          end: line.length,
          message: "Unterminated RusDox expression; close it with }}."
        });
        break;
      }
      const expression = line.slice(cursor + 2, end);
      expression.split("|").slice(1).map((part) => part.trim()).forEach((filter) => {
        if (!FILTERS.has(filter) && !/^default\((['"]).*\1\)$/.test(filter)) {
          const start = line.indexOf(filter, cursor);
          diagnostics.push({
            severity: "error",
            line: lineIndex,
            start,
            end: start + filter.length,
            message: "Unknown deterministic filter '" + filter + "'."
          });
        }
      });
      cursor = end + 2;
    }
  });
  return diagnostics;
}

function completionItems(linePrefix) {
  if (/type\s*:\s*$/.test(linePrefix) || /"type"\s*:\s*"?$/.test(linePrefix)) {
    return BLOCK_TYPES.map((type) => ({label: type, detail: "RusDox block type"}));
  }
  return Object.keys(HOVERS).map((label) => ({label, detail: HOVERS[label]}));
}

function hoverFor(word) {
  return HOVERS[word] || null;
}

module.exports = {analyze, completionItems, hoverFor, BLOCK_TYPES};
