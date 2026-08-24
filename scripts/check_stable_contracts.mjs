#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const read = (relative) => fs.readFileSync(path.join(root, relative), "utf8");
const cargo = read("Cargo.toml");
const spec = read("src/spec.rs");
const renderer = read("src/renderer.rs");
const protocol = read("src/protocol.rs");
const template = read("src/template.rs");
const stability = read("docs/stability.md");
const schema = JSON.parse(read("schema/rusdox-spec-v1.schema.json"));
const editorSchema = JSON.parse(read("editors/vscode/schema/rusdox-spec-v1.schema.json"));

assert(/rust-version\s*=\s*"1\.88"/.test(cargo), "Cargo MSRV must remain 1.88");
assert(/pub const SPEC_VERSION: u32 = 1;/.test(spec), "spec contract must remain v1");
assert(/pub const RENDERER_API_VERSION: u32 = 1;/.test(renderer), "renderer API must remain v1");
assert(/pub const PROTOCOL_VERSION: u32 = 1;/.test(protocol), "protocol must remain v1");
assert(/pub const TEMPLATE_SYNTAX_VERSION: &str = "1";/.test(template), "template syntax must remain v1");
assert(schema["x-rusdox-spec-version"] === 1, "generated schema must declare spec v1");
assert(schema.$id?.endsWith("/schema/rusdox-spec-v1.schema.json"), "schema ID must be canonical v1 URL");
assert(JSON.stringify(schema) === JSON.stringify(editorSchema), "editor and canonical schemas must match");
assert(schema.properties?.version?.const === 1, "schema must require version 1");
assert(schema.required?.includes("version"), "schema version field must be required");

for (const phrase of [
  "Rust crate API",
  "Deprecation window",
  "Supported releases",
  "Minimum supported Rust version",
  "cargo-semver-checks",
]) {
  assert(stability.includes(phrase), `stability policy is missing ${phrase}`);
}

console.log("Stable contract gate passed (spec/template/renderer/protocol v1, MSRV 1.88, canonical schema).");

function assert(condition, message) {
  if (!condition) {
    console.error(message);
    process.exit(1);
  }
}
