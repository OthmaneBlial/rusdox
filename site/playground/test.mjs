import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { cliCommand, downloadableName, parseSpecPreview } from "./core.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixtures = [
  ["board_report.yaml", "board-report"],
  ["executive_dashboard.yaml", "executive-dashboard"],
  ["product_launch_brief.yaml", "product-launch-brief"],
  ["talent_profile.yaml", "talent-profile"],
  ["invoice.yaml", "invoice"],
  ["meeting_notes.yaml", "meeting-notes"],
];

for (const [file, outputName] of fixtures) {
  const source = fs.readFileSync(path.join(root, "examples", file), "utf8");
  const parsed = parseSpecPreview(source);
  assert.equal(parsed.version, 1, file);
  assert.equal(parsed.outputName, outputName, file);
  assert.ok(parsed.blocks.length > 1, file);
  assert.deepEqual(parsed.diagnostics, [], file);
  assert.match(cliCommand("examples/" + file, outputName), new RegExp("rusdox examples/"));
  assert.equal(downloadableName(outputName), outputName + ".yaml");
}

const invalid = parseSpecPreview("version: 2\noutput_name: test\nblocks:\n");
assert.equal(invalid.diagnostics.length, 2);
assert.equal(downloadableName("../bad name"), "..-bad-name.yaml");
console.log("Playground core: 6 published examples and safety helpers passed.");
