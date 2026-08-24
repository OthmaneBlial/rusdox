#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const read = (relative) => fs.readFileSync(path.join(root, relative), "utf8");
const readiness = JSON.parse(read("compatibility/international-readiness.json"));
const parity = JSON.parse(read("reports/gallery/international-scripts-parity.json"));

assert(readiness.schema_version === 1, "international readiness contract must be v1");
assert(/^\d{4}-\d{2}-\d{2}$/.test(readiness.reviewed_at), "readiness review date is invalid");
assert(Array.isArray(readiness.tracks) && readiness.tracks.length === 2, "RTL and CJK tracks are required");
for (const track of readiness.tracks) {
  assert(["rtl", "cjk"].includes(track.id), `unknown track ${track.id}`);
  assert(["experimental", "supported"].includes(track.status), `${track.id}: invalid status`);
  assert(fs.existsSync(path.join(root, track.fixture)), `${track.id}: fixture is missing`);
  assert(fs.existsSync(path.join(root, track.parity_report)), `${track.id}: parity report is missing`);
  assert(Array.isArray(track.graduation_gates) && track.graduation_gates.length >= 5, `${track.id}: incomplete graduation gates`);
  for (const gate of track.graduation_gates) {
    assert(typeof gate.passed === "boolean", `${track.id}/${gate.id}: gate result must be boolean`);
    assert(fs.existsSync(path.join(root, gate.evidence)), `${track.id}/${gate.id}: evidence is missing`);
  }
  if (track.status === "supported") {
    assert(track.graduation_gates.every((gate) => gate.passed), `${track.id}: cannot graduate while a gate fails`);
  } else {
    assert(track.graduation_gates.some((gate) => !gate.passed), `${track.id}: experimental status needs an explicit unmet gate`);
  }
}

const checks = new Map(parity.checks.map((check) => [check.id, check]));
assert(checks.get("normalized_text")?.status === "passed", "international semantic round-trip must pass");
assert(checks.get("document_language")?.status === "passed", "international language metadata must pass");
assert(parity.expected.metadata.language === "mul", "international fixture must declare mul language metadata");

const metadata = read("src/metadata.rs");
const studio = read("src/studio.rs");
const xml = read("src/xml_utils.rs");
assert(metadata.includes('write_text_element(&mut writer, "dc:language"'), "DOCX dc:language output is missing");
assert(studio.includes("catalog.lang(TextStr(language))"), "PDF catalog /Lang output is missing");
assert(xml.match(/push_attribute\(\(\"descr\", alt_text\)\)/g)?.length >= 2, "DOCX alt text must reach both drawing properties");
assert(studio.includes("is_outline_embedding_allowed"), "font embedding permissions are not enforced");

const research = read("docs/pdf-conformance-research.md");
for (const phrase of ["not tagged PDF", "not PDF/UA", "not PDF/A", "veraPDF"]) {
  assert(research.includes(phrase), `PDF conformance research is missing '${phrase}'`);
}

console.log("International graduation and accessibility contract passed (RTL/CJK remain experimental).");

function assert(condition, message) {
  if (!condition) {
    console.error(message);
    process.exit(1);
  }
}
