#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const read = (relative) => fs.readFileSync(path.join(root, relative), "utf8");
const source = walk(path.join(root, "src"))
  .filter((file) => file.endsWith(".rs"))
  .map((file) => fs.readFileSync(file, "utf8"))
  .join("\n");
const review = read("docs/security-review-v1.md");

assert(!/\bunsafe\s*\{/.test(source), "src/ must not introduce unsafe blocks");
assert(read("src/package_validate.rs").includes("Event::DocType(_)"), "OOXML DTD rejection is missing");
assert(read("src/visual.rs").includes("resolve_string: Box::new(|_, _| None)"), "SVG file/URL resolution is enabled");
for (const field of [
  "max_template_expansions",
  "max_template_output_xml_bytes",
  "max_template_partial_depth",
]) {
  assert(read("src/limits.rs").includes(field), `missing template limit ${field}`);
  assert(read("src/template.rs").includes(field), `template renderer does not enforce ${field}`);
}
assert(read("src/bin/rusdox.rs").includes("default_value_t = ServiceLimitProfile::Hosted"), "serve must default to hosted limits");
assert(read("src/batch.rs").includes("max_total_source_bytes"), "batch aggregate-memory preflight is missing");
for (const phrase of [
  "Threat model and trust boundaries",
  "ZIP and OOXML",
  "Images and SVG",
  "Template expansion",
  "Protocol, hosted use, and batch work",
  "Residual risks and non-claims",
]) {
  assert(review.includes(phrase), `security review is missing '${phrase}'`);
}

console.log("v1 security source and review contract passed.");

function walk(directory) {
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const absolute = path.join(directory, entry.name);
    return entry.isDirectory() ? walk(absolute) : [absolute];
  });
}

function assert(condition, message) {
  if (!condition) {
    console.error(message);
    process.exit(1);
  }
}
