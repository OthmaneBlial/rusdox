#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const contract = JSON.parse(fs.readFileSync(path.join(root, "compatibility/v1-feature-contract.json"), "utf8"));
const spec = fs.readFileSync(path.join(root, "src/spec.rs"), "utf8");
const docs = fs.readFileSync(path.join(root, "docs/compatibility.md"), "utf8");
const formats = new Set(["supported", "bounded", "fallback", "unsupported", "not-applicable"]);
const states = new Set(["supported", "partial", "experimental", "unsupported"]);
const parity = new Set(["automated", "viewer", "not-applicable"]);

assert(contract.schema_version === 1, "compatibility schema must be v1");
assert(contract.document_spec_version === 1, "compatibility contract must target spec v1");
assert(/^\d{4}-\d{2}-\d{2}$/.test(contract.reviewed_at), "reviewed_at must be an ISO date");
assert(Array.isArray(contract.features) && contract.features.length >= 30, "v1 contract must enumerate at least 30 capabilities");

const ids = new Set();
const mappedVariants = new Set();
for (const feature of contract.features) {
  assert(!ids.has(feature.id), `duplicate feature ${feature.id}`);
  ids.add(feature.id);
  assert(formats.has(feature.docx), `${feature.id}: invalid DOCX status`);
  assert(formats.has(feature.pdf), `${feature.id}: invalid PDF status`);
  assert(states.has(feature.contract), `${feature.id}: invalid contract status`);
  assert(parity.has(feature.parity), `${feature.id}: invalid parity status`);
  assert(Array.isArray(feature.evidence) && feature.evidence.length > 0, `${feature.id}: evidence is required`);
  for (const evidence of feature.evidence) {
    assert(fs.existsSync(path.join(root, evidence)), `${feature.id}: missing evidence ${evidence}`);
  }
  for (const variant of feature.spec_variants || []) {
    assert(!mappedVariants.has(variant), `spec variant ${variant} is mapped more than once`);
    mappedVariants.add(variant);
  }
  if (feature.contract === "unsupported") {
    assert(feature.docx === "unsupported" || feature.pdf === "unsupported", `${feature.id}: unsupported status must be explicit`);
  }
}

const blockSection = spec.split("pub enum BlockSpec {")[1]?.split("/// A fully specified paragraph block.")[0] || "";
const blockVariants = [...blockSection.matchAll(/^    ([A-Z][A-Za-z0-9]+)(?:\s|\{|,)/gm)].map((match) => match[1]);
assert(blockVariants.length > 0, "could not extract BlockSpec variants");
for (const variant of blockVariants) assert(mappedVariants.has(variant), `unmapped BlockSpec variant ${variant}`);
for (const variant of mappedVariants) assert(blockVariants.includes(variant), `unknown mapped BlockSpec variant ${variant}`);
assert(docs.includes("v1-feature-contract.json"), "compatibility docs must link the machine contract");

console.log(`Compatibility contract passed (${contract.features.length} capabilities, ${blockVariants.length} BlockSpec variants).`);

function assert(condition, message) {
  if (!condition) {
    console.error(message);
    process.exit(1);
  }
}
